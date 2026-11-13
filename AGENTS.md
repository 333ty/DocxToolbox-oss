# AGENTS.md

给在本仓库工作的编码 agent 的项目说明。面向用户的用法文档见 [README.md](README.md)，格式支持边界、能力矩阵与扩展新格式时的分层契约见 [FORMAT-SPEC.md](FORMAT-SPEC.md)，同类工具调研结论与功能吸收清单见 [PRIOR-ART.md](PRIOR-ART.md)。

## 项目概览

单个 .NET 8 WinForms 工程（`DocxToolbox.csproj`，**没有 .sln，也没有测试工程**）。程序集名 `DocxToolbox`，根目录交付的 `Word 批处理工具箱.exe` 一个可执行文件两种模式：无参数启动图形界面，带参数走命令行。用途是批量清理 `.docx` 的页眉页脚 / 末尾空白节 / 文档属性、按哈希删除图片，以及按正则批量重命名文件名。

目标框架 `net8.0-windows`，`OutputType` 是 `WinExe`，依赖 `DocumentFormat.OpenXml` 3.0.1。

## 命令（均已实际验证）

```powershell
# 构建 → bin\Release\net8.0-windows\win-x64\DocxToolbox.exe
dotnet build .\DocxToolbox.csproj -c Release -r win-x64

# 发布单文件版（框架依赖，约 6.99 MB，目标机需 .NET 8 Desktop Runtime）
dotnet publish .\DocxToolbox.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true

# 发布产物在 bin\...\publish\DocxToolbox.exe；根目录那份 exe 不进版本库，需要手动覆盖
Copy-Item '.\bin\Release\net8.0-windows\win-x64\publish\DocxToolbox.exe' '.\Word 批处理工具箱.exe' -Force

# CLI 自检（PowerShell 不会等待 WinExe，必须走管道）
.\Word 批处理工具箱.exe --help | Out-Host
```

## 代码地图

| 文件 | 职责 |
| --- | --- |
| `program1.cs` | `Program.Main`（按 `args.Length` 分流 GUI/CLI）、`MainForm` 图形界面、`AppSettings` + `SettingsStore` 设置持久化 |
| `cli.cs` | `Cli.Run`：参数解析、路径收集、输出与退出码、`AttachConsole` 附加父控制台 |
| `wordcleaner.cs` | `WordCleaner`：GUI 与 CLI 共用的全部文档处理逻辑 |
| `renamer.cs` | `RenameRule` / `RenameItem` / `RenamePlanner`：按正则重命名文件名，GUI 与 CLI 共用。不引用 OpenXml，只做文件系统操作 |
| `renameform.cs` | `RenamePreviewForm`：重命名预览对话框，逐条列出原/新文件名与状态，勾选确认后才动文件 |
| `ocr.cs` | `LocalOcr`：调用本机 `ppocr.exe`（PaddleOCR 编译版）对**一个目录**的图片批量 OCR，返回「文件名 → 文字」。只在 `--watermark --ocr` 时用到，除 `Process` 外不依赖任何库 |

## 约定

- 文档处理逻辑只写在 `wordcleaner.cs`。GUI 和 CLI 都调用 `WordCleaner`，不要在 `cli.cs` 里再复制一份实现。
- 文件名重命名只写在 `renamer.cs`。**内容的归 `wordcleaner.cs`，名字的归 `renamer.cs`**，两边不要互相掺（`renamer.cs` 不应引用 `DocumentFormat.OpenXml`）。
- 重命名必须先整批出计划（`RenamePlanner.BuildPlan`）再执行（`RenamePlanner.Apply`）：重名判断要跨文件看，在单文件循环里逐个判是判不出来的。执行结果用 `RenamedMap` 回写 GUI 的文件列表，否则后续清理会指向已被改名的旧路径。
- GUI 的耗时操作一律 `await Task.Run(...)`，日志用 `IProgress<string>` 投回 UI 线程；工作线程里不要碰控件。运行时用 `_cts` 标记"忙"，`FormClosing` 里据此拦住关窗（避免 `.wtmp` 临时文件留在用户目录）。
- 增删 CLI 选项时，同步更新 `Cli.PrintUsage()` 和 README.md 的选项表。
- 退出码语义固定：`0` 全部成功、`1` 有文件失败、`2` 参数错误。
- 分节的"认法"只有一份：探测（`AnalyzeSections`）和删除（`RemoveLastSection`）都走 `ScanSections`，不要各写一套扫描逻辑。改分节判定时改这一个地方，两边的口径自动一致。
- 输出格式保持 `[成功]` / `[跳过]` / `[异常]` 前缀加末尾一行汇总，GUI 日志区与 CLI 共用同一批字符串。
- 输出流用 `Console.OutputEncoding` 构造：控制台下是系统代码页（中文 Windows 为 GBK），重定向时是 UTF-8，两边都能正常显示中文——不要硬编码成 UTF-8 或给控制台设 `chcp`。
- 设置文件在 `%AppData%\DocxToolbox\settings.json`。`SettingsStore.LegacyFilePath` 指向改名前的 `%AppData%\WordTool\settings.json`，只在读取时兜底一次，别当死代码删掉。
- 动文件格式支持前先读 [FORMAT-SPEC.md](FORMAT-SPEC.md)：处理层按格式各自实现、禁止共享；写入层（预检 → 内存校验 → `.wtmp` → 复校验 → 原子替换）与格式无关，必须泛化复用而不是复制一份。
- 水印/来源检测（`RemoveWatermarks`）只在**命中关键词**时动手，默认只删水印对象和命中文字，`--wm-drop-all-images` 才删全部图片。执行顺序固定：**探测 → 清水印对象与命中文字 → 对剩余图片做 OCR → 孤儿媒体收尾**；不要把 OCR 提到清理之前，否则既浪费识别时间，判断的也不是「清完水印后剩下的图」。
- 命中文字用 `RemoveKeywordText` 删**关键词子串**，不要整段/整个 run 删——正文里「本题由某某网整理」要变成「本题由整理」。跨 run 拆写（`某公`+`司`）要能删：按段落把 `w:t` 拼起来定位，再逐字回写。
- 判定一律按 `LocalName` + 属性名（`alt`/`string`/`descr`/`title`），不要按前缀写死命名空间：VML 的 `alt`、DrawingML 的 `descr` 分属不同命名空间，`GetAttribute(name, ns)` 猜错一次就静默失效。
- OCR 一律「一次进程跑一个目录」（`LocalOcr.Recognize`），不要按图片循环起进程——每次都要重新加载模型，20 张图从 14 秒变成 100 秒以上。OCR 失败/超时只记警告，绝不抛出去让整个文件处理失败。

## 已知陷阱

1. `OpenXmlPartContainer.DataPartReferenceRelationships` 在 SDK 3.0.1 下对主文档的图片关系返回空集合，用它取 rId 会静默失效（图片被删、引用残留，Word 报文档损坏）。要用 `Parts`（关系 ID → 部件映射）或 `GetIdOfPart()`，见 `WordCleaner.GetRelationshipIdsFor`。
2. 删除 `<a:blip>` 不能只删 `Inline`/`Anchor`：会留下 schema 不允许的空 `<w:drawing/>`。要连 `w:drawing` 一起删；VML 路径删承载的 `shape`。
3. 程序是 WinExe。从 PowerShell 直接调用不等待、可能丢输出，CLI 内部用 `AttachConsole(-1)` 附加父控制台；自动化测试要用管道 `| Out-Host`、`Start-Process -Wait` 或重定向到文件。
4. 处理是就地改写，没有备份机制。写盘统一走 `WordCleaner.ProcessFile`（预检 → 内存校验取基线 → 同目录 `.wtmp` 临时文件上改 → 再校验比对 → `File.Replace` 原子替换），不要绕开它直接在 `WordprocessingDocument.Open(path, true)` 上改原文件，也不要擅自加"自动另存一份"这类隐式行为。
5. `OpenXmlValidator` 校验必须走内存流（`MemoryStream`）而不是文件路径：SDK 打开损坏的包失败时会漏出文件句柄，导致临时文件删不掉、后续复用失败。
6. 全部源码与 csproj 都是 **UTF-8 无 BOM**，含中文注释、中文输出字符串。PowerShell 5.1 的 `Set-Content -Encoding UTF8` 会写入 BOM，改文件请用 `[System.IO.File]::WriteAllText($p, $t, (New-Object System.Text.UTF8Encoding($false)))`。行尾目前不统一（`cli.cs`、`README.md`、`AGENTS.md` 是 LF，`program1.cs`、`wordcleaner.cs`、`.csproj` 是 CRLF），不要为了统一而整文件重写行尾，那会让 diff 无法审查。
7. `.gitignore` 忽略 `*.exe`、`bin/`、`obj/`、`*.rar`、`*.zip`、`.reasonix/`、`.workbuddy/`。发布产物和构建产物都不进版本库，历史里也不要有二进制。
8. SDK 不认识的关系（典型是 `customXml` 的 `itemProps`）既不会出现在 `doc.GetAllParts()` 里，`DeletePart` 也删不掉，会在 zip 里变成不可达条目。这类残留只能在包关闭后用 `System.IO.Compression` 从 `_rels/.rels` 出发做可达性分析再删（见 `RemoveUnreachablePackageEntries`）。注意 `DataPart` 不是 `OpenXmlPart`，做可达性遍历时不要混着递归。
9. 清理项的执行顺序固定：页眉页脚 → 末尾空白节 → 属性 → 隐私 → 修订/批注（`--track-changes`）→ 图片 → 媒体/包级清扫。包级清扫必须放在最后，而且只能在整个包关闭之后做。
10. 重命名的正则只作用于**主文件名**（`Path.GetFileNameWithoutExtension`），扩展名由程序拼回去，不要拿整个文件名跑正则——否则 `.docx` 会被改掉。新名还要额外拦三种 Windows 特有的坑：结尾的空格/句点会被系统自动砍掉、CON/NUL 这类保留设备名建不出来、`Path.GetInvalidFileNameChars()` 之外的链式改名（A→B、B→C）执行顺序敏感一律判冲突。
11. CLI 里 `--rename` 的正则如果以 `-` 开头会被参数解析当成选项，必须写成 `--rename=-?\d+`；PowerShell 里替换模板的 `$1` 在双引号中会被展开成空串，只能用单引号或 `` `$1 `` 转义。
12. `w:sectPr` 只能出现在 `w:pPr` 里或 `body` 末尾。手工拼夹具时写成 `<w:p><w:sectPr>...` 是非法结构，SDK 会把它当未知元素，`Descendants<SectionProperties>()` 一个都扫不到——现象是"三节的文档被识别成 1 节"。另外分节符可能藏在表格单元格、内容控件里，定位必须用 `Descendants` 再回溯到顶层块，不能只看 `body.Elements<Paragraph>()`。
13. 修订/批注处理（`RemoveTrackChangesAndComments`）有两个必须成对的动作：① 批注锚点（`w:commentRangeStart/End`、`w:commentReference`）与 `comments.xml` 部件必须同批删除，只删一边 Word 都会报文档损坏——锚点散在正文里，部件挂在关系图上，漏了哪个都验证不过；② `w:ins`/`w:moveTo` 是解包保留内容，`w:del`/`w:moveFrom` 是连内容删除（删 `del` 时其 `delText` 随元素消失），不能搞反，否则可见内容会丢。修订标记可能藏在页眉页脚、脚注、尾注里，扫描走 `CollectContentParts`，不能只看 `document.xml`。匹配按 `LocalName`（Word 命名空间内），不依赖前缀。

## 验证方式

没有测试工程，改动文档处理逻辑后用这条链路验证：

1. 用 PowerShell 手工拼一个最小 `.docx`：`[Content_Types].xml`、`_rels/.rels`、`word/document.xml`（两个分节 + 页眉引用 + 一个 `<w:drawing>` 内嵌图）、`word/_rels/document.xml.rels`、`word/header1.xml`、`word/media/image1.png`、`docProps/core.xml` + `app.xml`，再用 `ZipFile.CreateEntryFromFile` 打包，**条目名必须用 `/` 分隔**。
2. 用 CLI 处理副本：先 `--dry-run` 并比对文件哈希确认没被改动，再做实际处理。
3. 解包检查条目增删、`document.xml` 里 `blip`/`drawing` 是否清干净、`rels` 里对应关系是否消失。
4. 用 `DocumentFormat.OpenXml.Validation.OpenXmlValidator` 校验处理后的文档，要求与原始文件同为 0 错误——这是判断"文档有没有被写坏"最可靠的信号。
5. 改分节删除逻辑时，额外准备一个三节夹具（第 1 节有正文 A4、第 2/3 节空白 A5/LETTER），验证四点：① 末尾连续空白节一次删完，输出「删除末尾空白节 2 个」；② 正文文字还在；③ 剩下的 `w:pgSz` 是第 1 节的 A4（说明页面设置是逐级继承、不是丢失）；④ 再跑一次全部 `[跳过]`、文件哈希不变。
6. 改重命名逻辑时，准备一批文件名同时覆盖五种情形：未匹配、目标已存在、替换后为空、含非法字符、同批次撞名。先 `--rename=... -n` 干跑确认一个文件都没动，再实跑核对改名结果与每条跳过原因；非法正则要返回退出码 2。
7. 自动化跑 CLI 时，用 `subprocess.run(..., creationflags=0x8)`（DETACHED_PROCESS）绕开 `AttachConsole(-1)` 对控制台句柄的劫持，输出才进得来管道；此时编码是 GBK，按 `gbk` 解码。
8. 改修订/批注逻辑时，准备一个含 `w:ins`、`w:del`、`w:rPrChange`、批注锚点 + `comments.xml` 部件的夹具，验证四点：① `ins` 的文字保留（解包）、`del` 的文字消失；② `comments.xml` 及其关系和 `[Content_Types].xml` Override 一起消失；③ `OpenXmlValidator` 错误数不增；④ 连跑两次第二次全部 `[跳过]` 且哈希不变（幂等）。批注锚点与部件不同批的损坏场景也要能复现「校验后放弃写入、原文件不动」。
9. 改水印检测逻辑时，用一份带品牌水印的测试文档（触发词自定，不要预置任何真实品牌词）验证六点：① 页眉页脚里 `v:shape/@alt`、`v:textpath/@string` 命中；② 正文 `v:group alt="品牌词"` 连空组一起清（只剩空组也算法不对）；③ `wp:docPr/@descr` 的 logo 图按整块 `w:drawing`（`mc:AlternateContent` 时连外壳）删除；④ 0.05pt/0.75pt 埋点图消失；⑤ **正常插图数量不变**（拿 `--validate` 前后的 `<a:blip>` / `<v:imagedata>` 计数对比）；⑥ `--validate` 错误数不高于原文件。要准备一份「没有水印」的文档跑一遍，确认输出 `[跳过] 未检出指定文字` 且文件哈希不变。
10. 改 OCR 调用时先单独跑 `D:\software\ppocr\ppocr.exe --help` 确认参数还在；再用 2~3 张图测耗时。注意 ppocr 的输出是 UTF-8（`cmd` 默认 GBK 会乱码），解析走 stdout 的 `predict img:` / `rec text:` 两行，不要依赖它落盘的 `.txt` 目录结构。
