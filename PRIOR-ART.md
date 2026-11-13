# PRIOR-ART.md — 同类项目调研与吸收清单

调研时间：2026-10-25。数据来源：GitHub Search API（按 star 排序）与各项目源码/README 直读。
本文件记录**别人做了什么、我们缺什么、哪些值得吸收**。格式支持边界见 [FORMAT-SPEC.md](FORMAT-SPEC.md)，代码约定见 [AGENTS.md](AGENTS.md)。

---

## 1. 赛道结论

这类需求**没有事实标准**。docx 专用清理工具的 star 普遍在 0–7 之间，元数据清理类最高的也只有 59。star 少不是需求不存在，而是需求分散在各人的临时脚本里。

更关键的一条：**没有项目同时满足「Windows 原生 + 免安装单文件 + GUI/CLI 双模式 + 不依赖 Office」**。这是当前的空位。

---

## 2. 竞品清单

### A. 隐私与元数据方向（占九成）

| 项目 | 形态 | 关键点 |
| --- | --- | --- |
| **mat2** | Python CLI，Whonix 预装、Debian 官方打包 | 最成熟。支持 OOXML 三件套，清理深度远超同类（见 §3） |
| **dmeta**（openscilab） | Python CLI + API | 专注 Office 元数据，支持就地改写，另提供 pre-commit hook |
| **sandy-sp/metadata-cleaner** | Python CLI | 59★，多格式元数据，带独立可执行文件 |
| **00DaEdRa00/OfficeMetaCleaner** | C# WPF + CLI | 与本项目最同构的 Windows 竞品，额外覆盖旧版 `.doc/.xls/.ppt` 与 Access |
| **xhu96/docuclean** | 浏览器端 | 清完**重新解析并校验**，是少数有「改完要验证」意识的实现 |
| **sergiomarotco/word-metadata-cleaner** | PowerShell | Word 专用元数据，14★ |

### B. 结构清理方向（很少）

| 项目 | 形态 | 关键点 |
| --- | --- | --- |
| **tuhalang/RemoveHeaderAndFooterInDocx** | Java | 只删页眉页脚 |
| **nowscott/DocxHeaderFooterCleaner** | Python + python-docx | 删页眉页脚**并在页脚插页码**，带备份与递归批量 |
| **justou 的 gist** | win32com | 调 Word 删页眉页脚，**必须装 Office** |
| **Hefaistos68/docxCleaner** | C# CLI | 命令行 docx 清理 |
| **zwt0204/standardize-docx-format** | Python，中文 | 论文排版流水线，审计 → 应用 → 验收三段式，处理分节与页眉页脚 |

---

## 3. 与 mat2 的逐项对照

读完 mat2 的 `libmat2/office.py`（`MSOfficeParser`）后逐项比对：

| 能力 | mat2 | 本项目 |
| --- | --- | --- |
| 清空 `docProps/core.xml`、`app.xml`，删 `custom.xml` 与缩略图 | ✅ | ✅ |
| 删 `rsid`（元素与属性） | ✅ | ✅ |
| 删 `customXml/` 整目录 | ✅ | ✅ |
| 删悬空 `Override` / `Relationship` | ✅ | ✅ |
| **删页眉页脚** | ❌ 保留，只清其中的修订与批注 | ✅ |
| **删末尾空白节** | ❌ | ✅ 独有 |
| **按哈希删指定图片** | ❌ | ✅ 独有 |
| **交互式分节探测** | ❌ | ✅ 独有 |
| **修订痕迹**（`ins`/`del`/`move` + `author`/`date`/`userId`） | ✅ 彻底 | ✅ `--track-changes`：ins/moveTo 解包保留、del/moveFrom 连内容删、格式修订与身份属性全清 |
| **批注锚点**（`commentRangeStart/End`、`commentReference`） | ✅ | ✅ 锚点与 `comments.xml` 部件同批整删 |
| `nsid` | ✅ | ❌ |
| 属性排序规范化 | ✅ | ❌ |
| 就地改写 | ❌ 生成 `.cleaned` 新文件 | ✅ 原子替换 |
| GUI | 仅第三方前端 | ✅ 原生 |

**最刺眼的一行是修订与批注**：本项主打「清除隐私残留」，却把最容易暴露身份的两类痕迹留在了文档里——属性清得再干净，批注里仍挂着「张三 2024-03-15」，修订记录里仍有每个人每次改动的姓名与时间。

---

## 4. 本项目的位置

差异点不在「功能更多」，而在工程约束：

1. **Windows 原生单文件 exe**，双击即用，不装 Python、不装 Office。
2. **GUI 与 CLI 同一份实现**，参数化脚本与手工操作都能覆盖。
3. **写入前先建校验基线**：只有 `OpenXmlValidator` 错误数不变差才落盘，否则原文件不动。竞品里只有 docuclean 做了类似的事。
4. **干跑可哈希比对、重复运行幂等**：第二次跑全 `[跳过]` 且文件哈希不变。

第 3、4 条比「多支持一种格式」更值得写进 README 首屏。

---

## 5. 真正的空白

搜索 "remove image from docx batch" 仅返回 1 条无关结果，基本可判定无现成轮子：

- **按图片哈希批量删图**：无。中文办公场景里「投标文件批量去公司 logo」是真实需求。
- **删末尾空白节 + 页面设置逐级提升**：无。`standardize-docx-format` 有 audit，但目的是排版验收，不删节。

这两项是差异化资产，优先级高于「支持 xlsx」。

---

## 6. 吸收清单

状态标记：⬜ 未实现 ｜ 🟡 部分 ｜ ✅ 已实现

### P0 — 修订痕迹与批注锚点 ✅

**实现**：`--track-changes`（`-tc`），2026-10-25 落地。独立选项，不并入 `--privacy`/`--all`；GUI 勾选与 CLI 文案均写明「会接受所有修订、改变可见内容」。批注锚点与 `comments.xml`（含 commentsExtended/commentsIds/people）同批整删，修订标记扫描覆盖页眉页脚、脚注、尾注。

**目标**：把「谁在什么时候改了什么」从文档里抹掉。

**元素处理规则**（WordprocessingML 命名空间，按本地名精确匹配）：

| 元素 | 处理 |
| --- | --- |
| `w:del`、`w:cellDel`、`w:cellMerge` | 连内容一起删（被删文字在 `w:delText` 里，随元素删除） |
| `w:moveFrom`、`w:moveFromRangeStart/End`、`w:moveToRangeStart/End` | 删 |
| `w:ins`、`w:moveTo`、`w:cellIns` | **解包保留内容**：把子元素按原位置插入父节点，只去掉包裹层 |
| `w:rPrChange`、`w:pPrChange`、`w:sectPrChange`、`w:tblPrChange`、`w:trPrChange`、`w:tcPrChange`、`w:tblGridChange`、`w:numberingChange` | 删 |
| `w:customXmlIns/Del/MoveFrom/MoveTo` 及其 `RangeStart/End` | 删 |
| `w:author`、`w:date`、`w:userId`、`w:initials`、`w:ed`（Word 命名空间属性） | 全树遍历删除 |
| `w:commentRangeStart`、`w:commentRangeEnd`、`w:commentReference` | 删锚点，**并同步删除 `comments.xml` 部件、其关系与 `[Content_Types].xml` 中的 Override** |

**语义决策（必须先定，不能默认）**：处理修订等于「接受所有修订」，会改变文档可见内容（删除的段落真的消失）。因此：

- 必须是**独立选项**，不能并进 `--privacy` 或 `--all` 默认执行。
- GUI 勾选项与 CLI 选项的文案都要写明「会接受所有修订」。
- `w:ins` 解包而不是丢弃，否则整段新增内容会消失。

**风险**：

- 只删批注锚点而不删 `comments.xml`，Word 会报文档损坏——两者必须同批处理。
- 修订可能出现在页眉页脚、脚注、尾注里，扫描范围不能只看 `document.xml`。

**验收**：三节夹具中插入 `ins`/`del`/批注各一处，处理后 ①`ins` 的文字保留、`del` 的文字消失；②`comments.xml` 与关系消失；③`OpenXmlValidator` 错误数不增；④Word 打开无「此文档包含批注/修订」提示。

### P1 — `--report json` ⬜

**目标**：把处理结果变成可被脚本消费的结构，而不只是给人看的日志。

**要点**：

- JSON 走独立文件（`--report result.json`），stdout 日志格式**完全不变**，不影响既有脚本。
- 必含字段：`tool`、`version`、`timestamp`、`summary{total,success,failed,skipped}`、`files[{path,status,sizeBefore,sizeAfter,validatorErrors{before,after},actions[],error}]`。
- `validatorErrors` 是这份报告的核心价值——它把「没写坏」这件事从口头承诺变成可核对的数字。
- 与 `--dry-run` 组合时，报告里的 `actions` 为计划项、`sizeAfter` 与 `validatorErrors.after` 为 null。

**验收**：正常、跳过、异常三种文件各一，报告字段完整；`jq` 可解析；日志输出与不带 `--report` 时逐字节一致。

### P2 — 指纹消除三件套 ⬜（需裁剪，不能照搬）

mat2 的这一组里有两项**对只做 docx 的我们不适用**，照搬是浪费：

| 项目 | 对 docx 是否适用 | 结论 |
| --- | --- | --- |
| `w:nsid`（文档随机标识符，可用于判断两份文档同源） | ✅ 适用 | **做** |
| 属性排序规范化（按字典序重排属性，消除生产者指纹） | ✅ 适用 | **做**，但要注意它会重写整份 XML，必须靠校验基线兜底 |
| `p14:creationId`、`p:sldMasterId`、`tableStyles` GUID 随机化 | ❌ 只存在于 pptx | **不做**，等支持 pptx 时再议 |

**风险**：属性重排会改动大量节点，是唯一可能让校验结果变差的改动，必须走 `WordCleaner.ProcessFile` 的基线比对，不能绕开。

**验收**：处理前后 `OpenXmlValidator` 错误数一致；连续两次处理，第二次全部 `[跳过]`（保证幂等，顺序稳定）。

### P3 — 提交前拦截 ⬜

dmeta 提供 pre-commit hook，但它的生态前提是 Python 安装包。我们是 Windows exe，硬做 hook 属于生态错配。

**更合适的形态是 `--check` 模式**：

- 只检测不修改，发现残留元数据/修订/批注就返回退出码 `1`，干净返回 `0`。
- 直接可挂进任何 CI，也能被任意 pre-commit 框架调用（hook 里只需一行调 exe）。
- 与现有退出码语义（`0` 成功 / `1` 有失败 / `2` 参数错误）天然一致。

**验收**：干净文档返回 0；含作者名或未接受的修订返回 1 且不修改任何文件；`--check` 与清理选项同时给出时以参数错误（2）拒绝。

---

## 7. 不采纳的做法

**mat2 的激进白名单**：它按白名单删 zip 成员，连 `numbering.xml`、`glossary`、`theme`、`webSettings.xml` 都整个移除。删 `numbering.xml` 会让列表编号丢失，`theme` 会让配色漂移。代价是压不下来体积，换来的是文档不坏。

本项目的路线是「保守增量 + 校验兜底」，与它相反，并且这个取舍应当保留——**体积不是本工具的卖点，不写坏才是**。

---

## 8. 参考

- mat2：<https://0xacab.org/jvoisin/mat2>（GitHub 镜像 `jvoisin/mat2`），OOXML 逻辑在 `libmat2/office.py`
- dmeta：<https://github.com/openscilab/dmeta>
- docuclean：<https://github.com/xhu96/docuclean>
- OfficeMetaCleaner：<https://github.com/00DaEdRa00/OfficeMetaCleaner>
- standardize-docx-format：<https://github.com/zwt0204/standardize-docx-format>
