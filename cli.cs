using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace DocxToolbox;

/// <summary>
/// 命令行接口。同一个 exe：带参数走 CLI，无参数启动图形界面。
/// </summary>
internal static class Cli
{
    private const int ExitOk = 0;
    private const int ExitFailure = 1;
    private const int ExitUsage = 2;

    private const int AttachParentProcess = -1;

    private static bool _quiet;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AttachConsole(int dwProcessId);

    public static int Run(string[] args)
    {
        EnsureConsole();
        _quiet = false;

        try
        {
            return Execute(args);
        }
        catch (Exception ex)
        {
            WriteError($"[异常] 未预期的错误: {ex.Message}");
            return ExitFailure;
        }
    }

    private static int Execute(string[] args)
    {
        var o = Parse(args);
        if (o == null) return ExitUsage;

        _quiet = o.Quiet;

        if (o.Help) { PrintUsage(); return ExitOk; }
        if (o.Version) { WriteLineAlways($"{ExeName} {VersionText}"); return ExitOk; }

        // 只读模式：计算单个图片文件的参考哈希
        if (o.HashFile != null)
        {
            if (!File.Exists(o.HashFile))
            {
                WriteError($"[异常] 找不到文件: {o.HashFile}");
                return ExitFailure;
            }

            string hash = WordCleaner.ComputeFileHash(o.HashFile, o.Algo);
            WriteLineAlways($"{o.Algo}: {hash}  ({o.HashFile})");
            return ExitOk;
        }

        if (o.All)
        {
            o.Headers = true;
            o.LastSection = true;
            o.Props = true;
        }

        // 内容清理与重命名是两件独立的事，分开判定：只给 --rename 时不该去动文档内容
        bool doContent = o.Headers || o.LastSection || o.Props || o.Privacy || o.Orphans
                          || o.ImageHashes.Count > 0 || o.TrackChanges || o.WatermarkKeywords.Count > 0;
        bool doRename = o.RenamePattern != null || o.CleanName;

        if (!doContent && !doRename && !o.Analyze && !o.Validate)
        {
            WriteError("请至少指定一个操作: --headers / --last-section / --props / --privacy / --track-changes / --watermark / --orphans / --image-hash / --all / --rename / --clean-name / --analyze / --validate");
            PrintUsage();
            return ExitUsage;
        }

        if (o.Paths.Count == 0)
        {
            WriteError("请提供要处理的 .docx 文件或文件夹路径（可以给多个）。");
            PrintUsage();
            return ExitUsage;
        }

        if (o.WatermarkKeywords.Count == 0 && (o.WatermarkDropAll || o.OcrDir != null))
        {
            WriteError("--wm-drop-all-images / --ocr 需要配合 --watermark 使用（先给触发词，例如 --watermark=某公司）。");
            return ExitUsage;
        }

        if (o.ImageHashes.Count > 0)
        {
            int expected = o.Algo switch { "SHA1" => 40, "SHA256" => 64, _ => 32 };
            foreach (string h in o.ImageHashes)
            {
                if (h.Length != expected || !h.All(Uri.IsHexDigit))
                {
                    WriteError($"[异常] 图片哈希不是合法的 {o.Algo} 值（应为 {expected} 位十六进制字符）: {h}");
                    return ExitUsage;
                }
            }
        }

        var files = CollectFiles(o);
        if (files.Count == 0)
        {
            WriteError("[异常] 没有找到任何可处理的 .docx 文件。");
            return ExitFailure;
        }

        WriteLineAlways($"待处理文件: {files.Count} 个{(o.DryRun ? "（干跑模式，不会修改任何文件）" : "")}");

        int succeeded = 0;
        int failed = 0;

        var request = new CleanRequest
        {
            Headers = o.Headers,
            LastSection = o.LastSection,
            Props = o.Props,
            Privacy = o.Privacy,
            Orphans = o.Orphans,
            ForceLastSection = o.Force,
            Algo = o.Algo,
            TrackChanges = o.TrackChanges,
        };
        request.ImageHashes.AddRange(o.ImageHashes);
        request.WatermarkKeywords.AddRange(o.WatermarkKeywords);
        request.WatermarkDropAllImages = o.WatermarkDropAll;
        request.OcrDir = o.OcrDir;
        request.OcrMaxImages = o.OcrMaxImages;

        foreach (string file in files)
        {
            if (o.Validate)
            {
                int errors = WordCleaner.ValidateDocument(file, out string? reason);
                if (errors < 0)
                {
                    WriteError($"[异常] 无法校验: {file} - {reason}");
                    failed++;
                }
                else
                {
                    WriteLine($"[校验] {errors} 个错误: {file}");
                    // 有错误就记失败：这个模式的用处就是让脚本拿退出码判断文档有没有被写坏
                    if (errors > 0) failed++;
                    else succeeded++;
                }
            }

            if (o.Analyze)
            {
                try
                {
                    using var doc = WordprocessingDocument.Open(file, false);
                    WordCleaner.AnalyzeSections(doc, file, WriteLine);
                }
                catch (Exception ex)
                {
                    WriteError($"[异常] 分析失败: {file} - {ex.Message}");
                    failed++;
                    continue;
                }
            }

            if (!doContent)
            {
                // 只读模式（--validate / --analyze）：上面已经记过数，不再重复计数
                if (!o.Validate) succeeded++;
                continue;
            }

            if (o.DryRun)
            {
                WriteLine($"[干跑] {file} 将执行: {PlanText(o)}");
                if (o.WatermarkKeywords.Count > 0) ReportWatermarkScan(file, o);
                succeeded++;
                continue;
            }

            // 复制到临时文件 → 处理 → 校验 → 原子替换，失败时原文件不动
            if (WordCleaner.ProcessFile(file, request, WriteLine))
            {
                succeeded++;
            }
            else
            {
                failed++;
            }
        }

        // 重命名放在内容清理之后：清理失败的文件不会被改名，日志里的路径始终指向真实存在的文件。
        // 它必须整批一起算（跨文件判重），所以不放在上面的单文件循环里。
        if (doRename)
        {
            // --rename 显式给的规则优先；只给 --clean-name 时套用内置的「去 +/空格」规则
            string pattern = o.RenamePattern ?? NameCleanup.Pattern;
            string replacement = o.RenamePattern != null ? (o.RenameTo ?? "") : "";

            var rule = RenameRule.TryCreate(pattern, replacement, o.RenameIgnoreCase, out string renameErr);
            if (rule == null)
            {
                WriteError($"[异常] {renameErr}");
                return ExitUsage;
            }

            var plan = RenamePlanner.BuildPlan(files, rule);

            if (o.DryRun)
            {
                foreach (var it in plan)
                {
                    if (it.Status == RenameStatus.Ready)
                        WriteLine($"[干跑] 重命名: {it.OldName} → {it.NewName}");
                    else if (it.IsProblem)
                        WriteLine($"[干跑] 重命名跳过: {it.OldName} - {it.StatusText}");
                }
            }
            else
            {
                var (renamedOk, renamedFailed) = RenamePlanner.Apply(plan, WriteLine);
                failed += renamedFailed;
                WriteLineAlways($"------ 重命名: 成功 {renamedOk}，失败 {renamedFailed} ------");
            }
        }

        WriteLineAlways($"------ 完成: 共 {files.Count} 个文件，成功 {succeeded}，失败 {failed}"
                        + $"{(o.DryRun ? "（干跑模式，未修改任何文件）" : "")} ------");

        return failed > 0 ? ExitFailure : ExitOk;
    }

    #region 参数解析

    private sealed class Options
    {
        public readonly List<string> Paths = new();
        public bool Headers;
        public bool LastSection;
        public bool Props;
        public bool All;
        public bool Analyze;
        public bool Validate;
        public bool DryRun;
        public bool NoRecurse;
        public bool Quiet;
        public bool Help;
        public bool Version;
        public string? HashFile;
        public readonly List<string> ImageHashes = new();
        public bool Privacy;
        public bool Orphans;
        public bool Force;
        public bool TrackChanges;
        public string Algo = "MD5";
        public string? RenamePattern;
        public string? RenameTo;
        public bool RenameIgnoreCase;
        public bool CleanName;
        public readonly List<string> WatermarkKeywords = new();
        public bool WatermarkDropAll;
        public string? OcrDir;
        public int OcrMaxImages = 60;
    }

    private static Options? Parse(string[] args)
    {
        var o = new Options();

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];

            if (arg == "--")
            {
                for (i++; i < args.Length; i++) o.Paths.Add(args[i]);
                break;
            }

            if (arg.Length == 0) continue;

            if (arg[0] != '-')
            {
                o.Paths.Add(arg);
                continue;
            }

            string name = arg;
            string? value = null;
            int eq = arg.IndexOf('=');
            if (eq > 0)
            {
                name = arg.Substring(0, eq);
                value = arg.Substring(eq + 1);
            }

            // 长选项不区分大小写；短选项区分（-A 探测 / -a 全部）
            string key = name.StartsWith("--", StringComparison.Ordinal) ? name.ToLowerInvariant() : name;

            switch (key)
            {
                case "-h":
                case "-?":
                case "/?":
                case "--help":
                    o.Help = true;
                    break;

                case "--version":
                    o.Version = true;
                    break;

                case "-hf":
                case "--headers":
                    o.Headers = true;
                    break;

                case "-ls":
                case "--last-section":
                    o.LastSection = true;
                    break;

                case "-p":
                case "--props":
                    o.Props = true;
                    break;

                case "-a":
                case "--all":
                    o.All = true;
                    break;

                case "-A":
                case "--analyze":
                    o.Analyze = true;
                    break;

                case "-vd":
                case "--validate":
                    o.Validate = true;
                    break;

                case "-n":
                case "--dry-run":
                    o.DryRun = true;
                    break;

                case "-q":
                case "--quiet":
                    o.Quiet = true;
                    break;

                case "--no-recurse":
                    o.NoRecurse = true;
                    break;

                case "-al":
                case "--algo":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    string algo = value.Trim().ToUpperInvariant();
                    if (!WordCleaner.SupportedAlgorithms.Contains(algo))
                        return Usage($"不支持的哈希算法: {value}（可选 MD5 / SHA1 / SHA256）。");
                    o.Algo = algo;
                    break;

                case "-ih":
                case "--image-hash":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    // 可以重复给多次，也允许一个参数里用逗号/分号分隔多个
                    foreach (string h in SplitHashes(value))
                    {
                        o.ImageHashes.Add(h);
                    }
                    break;

                case "--image-hash-file":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    if (!File.Exists(value))
                        return Usage($"找不到哈希清单文件: {value}");
                    try
                    {
                        foreach (string line in File.ReadAllLines(value))
                        {
                            string text = line.Trim();
                            if (text.Length == 0 || text.StartsWith("#", StringComparison.Ordinal)) continue;
                            foreach (string h in SplitHashes(text))
                            {
                                o.ImageHashes.Add(h);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return Usage($"读取哈希清单失败: {value} - {ex.Message}");
                    }
                    break;

                case "-pr":
                case "--privacy":
                    o.Privacy = true;
                    break;

                case "-tc":
                case "--track-changes":
                    o.TrackChanges = true;
                    break;

                case "-oc":
                case "--orphans":
                    o.Orphans = true;
                    break;

                case "--force":
                    o.Force = true;
                    break;

                case "-rn":
                case "--rename":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    // 正则本身可能以 - 开头，那种情况请写成 --rename=<正则>
                    o.RenamePattern = value;
                    break;

                case "-rt":
                case "--rename-to":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    o.RenameTo = value;
                    break;

                case "--rename-ignore-case":
                    o.RenameIgnoreCase = true;
                    break;

                case "-cn":
                case "--clean-name":
                    o.CleanName = true;
                    break;

                case "-wm":
                case "--watermark":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    foreach (string k in SplitKeywords(value))
                    {
                        o.WatermarkKeywords.Add(k);
                    }
                    break;

                case "--wm-drop-all-images":
                    o.WatermarkDropAll = true;
                    break;

                case "--ocr":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    o.OcrDir = value.Trim().Trim('"');
                    break;

                case "--ocr-max":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    if (!int.TryParse(value, out int ocrMax) || ocrMax < 0)
                        return Usage("--ocr-max 需要一个非负整数（0 表示不限）。");
                    o.OcrMaxImages = ocrMax;
                    break;

                case "--hash-file":
                    value ??= TakeValue(args, ref i);
                    if (value == null) return Usage($"选项 {name} 缺少参数值。");
                    o.HashFile = value.Trim();
                    break;

                default:
                    return Usage($"未知选项: {name}");
            }
        }

        return o;
    }

    private static string? TakeValue(string[] args, ref int i)
        => (i + 1 < args.Length) ? args[++i] : null;

    /// <summary>把 "hash1,hash2;hash3" 或一行一个的写法拆成若干个哈希。</summary>
    private static IEnumerable<string> SplitHashes(string text)
        => text.Split(new[] { ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
               .Select(h => h.Trim().ToLowerInvariant())
               .Where(h => h.Length > 0);

    /// <summary>
    /// 把 "某公司,brand.example.com" 这类输入拆成多个触发词。只按中英文逗号/分号/顿号切，不按空格切——
    /// 触发词本身可能带空格（如 "某公司 水印"）。
    /// </summary>
    private static IEnumerable<string> SplitKeywords(string text)
        => text.Split(new[] { ',', ';', '，', '；', '、' }, StringSplitOptions.RemoveEmptyEntries)
               .Select(k => k.Trim())
               .Where(k => k.Length > 0);

    private static Options? Usage(string message)
    {
        WriteError(message);
        PrintUsage();
        return null;
    }

    #endregion

    #region 文件收集与输出

    private static List<string> CollectFiles(Options o)
    {
        var files = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var enumOptions = new EnumerationOptions
        {
            RecurseSubdirectories = !o.NoRecurse,
            IgnoreInaccessible = true,
            MatchCasing = MatchCasing.CaseInsensitive
        };

        foreach (string path in o.Paths)
        {
            if (Directory.Exists(path))
            {
                try
                {
                    var found = Directory.EnumerateFiles(path, "*.docx", enumOptions)
                                         .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                                         .OrderBy(f => f, StringComparer.OrdinalIgnoreCase);

                    foreach (string f in found)
                    {
                        if (seen.Add(f)) files.Add(f);
                    }
                }
                catch (Exception ex)
                {
                    WriteError($"[异常] 扫描文件夹失败: {path} - {ex.Message}");
                }
            }
            else if (File.Exists(path))
            {
                if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    WriteError($"[跳过] 只处理 .docx 文件: {path}");
                    continue;
                }
                if (seen.Add(path)) files.Add(path);
            }
            else
            {
                WriteError($"[异常] 路径不存在: {path}");
            }
        }

        return files;
    }

    private static string PlanText(Options o)
    {
        var parts = new List<string>();
        if (o.Headers) parts.Add("删除页眉页脚");
        if (o.LastSection) parts.Add(o.Force ? "删除最后一节（强制，含内容）" : "删除末尾空白节");
        if (o.Props) parts.Add("删除属性信息");
        if (o.Privacy) parts.Add("清除隐私残留（rsid/customXml/docVars/外部超链接）");
        if (o.TrackChanges) parts.Add("接受所有修订、清除批注（会改变文档可见内容）");
        if (o.WatermarkKeywords.Count > 0)
        {
            string mode = o.WatermarkDropAll ? "命中即删除全部图片" : "命中只删水印/埋点图";
            string ocr = o.OcrDir != null ? "，含图片 OCR" : "";
            parts.Add($"水印检测（{string.Join("、", o.WatermarkKeywords)}，{mode}{ocr}）");
        }
        if (o.ImageHashes.Count == 1) parts.Add($"删除哈希匹配的图片（{o.Algo}: {o.ImageHashes[0]}）");
        if (o.ImageHashes.Count > 1) parts.Add($"删除哈希匹配的图片（{o.Algo}，{o.ImageHashes.Count} 个哈希）");
        if (o.Orphans) parts.Add("清理未引用的媒体");
        if (o.RenamePattern != null) parts.Add($"按正则重命名文件名（{o.RenamePattern} → {o.RenameTo ?? "（空）"}）");
        else if (o.CleanName) parts.Add("清理文件名（去掉 + 和空格）");
        return parts.Count > 0 ? string.Join("、", parts) : "（无操作）";
    }

    /// <summary>
    /// 干跑时真扫一遍水印：命中什么、本来会删多少，但一个字节都不动。
    /// 和实跑共用同一套判定，所以这里的数字就是实跑会删的数字。
    /// </summary>
    private static void ReportWatermarkScan(string file, Options o)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(file, false);
            var wm = WordCleaner.ScanWatermarks(doc, o.WatermarkKeywords);

            if (!wm.Hit)
            {
                WriteLine($"[干跑] 未检出指定文字（{string.Join("、", o.WatermarkKeywords)}），图片不会动: {file}");
                return;
            }

            string why = wm.Evidence.Count > 0
                ? "命中：" + string.Join("；", wm.Evidence.Take(3))
                  + (wm.Evidence.Count > 3 ? $"…共 {wm.Evidence.Count} 类" : "")
                : "命中：图片文字";

            if (o.WatermarkDropAll)
            {
                WriteLine($"[干跑] 命中，将删除全部图片元素 {wm.ImageElements} 处；{why}: {file}");
            }
            else
            {
                WriteLine($"[干跑] 命中，将删除水印图形 {wm.ShapesRemoved} 处、命中文字 {wm.KeywordsRemoved} 处；{why}: {file}");
            }

            if (o.OcrDir != null)
            {
                WriteLine("[干跑] OCR 不会在干跑时执行，实跑才识别图片里的文字");
            }
        }
        catch (Exception ex)
        {
            WriteLine($"[干跑] 水印探测失败: {file} - {ex.Message}");
        }
    }

    private static void WriteLine(string message)
    {
        if (_quiet) return;
        try { Console.Out.WriteLine(message); } catch { /* 无可用输出时忽略 */ }
    }

    private static void WriteLineAlways(string message)
    {
        try { Console.Out.WriteLine(message); } catch { /* 无可用输出时忽略 */ }
    }

    private static void WriteError(string message)
    {
        try { Console.Error.WriteLine(message); } catch { /* 无可用输出时忽略 */ }
    }

    private static string ExeName
    {
        get
        {
            try { return Path.GetFileName(Environment.ProcessPath) ?? "DocxToolbox.exe"; }
            catch { return "DocxToolbox.exe"; }
        }
    }

    private static string VersionText
    {
        get
        {
            try { return typeof(Cli).Assembly.GetName().Version?.ToString() ?? "1.0.0.0"; }
            catch { return "1.0.0.0"; }
        }
    }

    /// <summary>
    /// 本程序是 WinExe：双击运行时不带控制台。从命令行启动时挂到父进程的控制台，
    /// 输出才可见；若 stdout/stderr 已被重定向（管道、文件），则沿用原有句柄。
    /// </summary>
    private static void EnsureConsole()
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                AttachConsole(AttachParentProcess);
            }

            var stdout = new StreamWriter(Console.OpenStandardOutput(), Console.OutputEncoding) { AutoFlush = true };
            Console.SetOut(stdout);

            var stderr = new StreamWriter(Console.OpenStandardError(), Console.OutputEncoding) { AutoFlush = true };
            Console.SetError(stderr);
        }
        catch
        {
            // 输出不可用时不影响主流程，WriteLine/WriteError 内部还会兜底
        }
    }

    private static void PrintUsage()
    {
        string e = ExeName;
        var sb = new System.Text.StringBuilder();

        sb.AppendLine($"{e} — Word 批处理工具箱");
        sb.AppendLine();
        sb.AppendLine("用法:");
        sb.AppendLine($"  {e}                            启动图形界面");
        sb.AppendLine($"  {e} <路径...> [选项]            命令行模式（路径可以是 .docx 文件或文件夹）");
        sb.AppendLine();
        sb.AppendLine("清理选项（可任意组合）:");
        sb.AppendLine("  -hf, --headers            删除页眉页脚");
        sb.AppendLine("  -ls, --last-section       删除末尾空白节（只删没有内容的节，有内容会跳过）");
        sb.AppendLine("       --force              配合 --last-section：不看空白判断，连内容一起删");
        sb.AppendLine("  -p,  --props              删除文档属性信息");
        sb.AppendLine("  -pr, --privacy            清除隐私残留：rsid、customXml、docVars、附加模板、外部超链接");
        sb.AppendLine("  -tc, --track-changes      接受所有修订并清除批注（会改变文档可见内容，请谨慎）");
        sb.AppendLine("  -ih, --image-hash <哈希>  删除哈希匹配的图片，可重复给出，也支持逗号分隔多个");
        sb.AppendLine("       --image-hash-file <文件>  从清单文件读哈希，一行一个，# 开头为注释");
        sb.AppendLine("  -al, --algo <算法>        图片哈希算法: MD5 | SHA1 | SHA256（默认 MD5）");
        sb.AppendLine("  -oc, --orphans            清理没人引用的媒体和包内失效条目（省体积）");
        sb.AppendLine("  -a,  --all                等于 --headers --last-section --props（不含隐私、修订/批注与媒体清理）");
        sb.AppendLine();
        sb.AppendLine("水印 / 来源检测（命中指定文字才动手，不命中一张图都不碰）:");
        sb.AppendLine("  -wm, --watermark <词>     触发词，可用逗号/分号/顿号分隔，可重复给");
        sb.AppendLine("                            触发词由用户自定，建议用完整品牌词而不是短片段");
        sb.AppendLine("       --wm-drop-all-images 命中就连正文插图一起删（激进，会毁掉试卷插图，慎用）");
        sb.AppendLine("       --ocr <目录>         对这些图片做文字识别（本机 ppocr.exe 所在目录）");
        sb.AppendLine("       --ocr-max <数量>     每个文档最多识别多少张图，0=不限（默认 60）");
        sb.AppendLine();
        sb.AppendLine("文件名重命名（只改文件名，不动文档内容）:");
        sb.AppendLine("  -rn, --rename <正则>      按正则重命名，只作用于主文件名，扩展名保留");
        sb.AppendLine("  -rt, --rename-to <替换>   替换模板，可用 $1 引用分组；省略表示删掉匹配的部分");
        sb.AppendLine("       --rename-ignore-case 匹配时忽略大小写");
        sb.AppendLine("  -cn, --clean-name         套用内置规则：去掉文件名里的 + 和各种空格（含全角）");
        sb.AppendLine("                            只删不改名；写了 --rename 时以 --rename 为准");
        sb.AppendLine("                            正则以 - 开头时请写成 --rename=<正则>；建议先加 -n 干跑");
        sb.AppendLine();
        sb.AppendLine("只读工具:");
        sb.AppendLine("  -A,  --analyze            逐节探测：块范围、段落/表格数、页眉页脚引用、是否空白");
        sb.AppendLine("  -vd, --validate           统计 OpenXML 校验错误数（0 = 结构合法，只读）");
        sb.AppendLine("       --hash-file <图片>    计算并打印一个图片文件的哈希值");
        sb.AppendLine();
        sb.AppendLine("通用选项:");
        sb.AppendLine("  -n,  --dry-run            只报告将要做什么，不修改文件");
        sb.AppendLine("       --no-recurse         扫描文件夹时不进入子文件夹（默认递归）");
        sb.AppendLine("  -q,  --quiet              只输出错误和最终汇总");
        sb.AppendLine("  -h,  --help               显示本帮助");
        sb.AppendLine("       --version            显示版本号");
        sb.AppendLine();
        sb.AppendLine("示例:");
        sb.AppendLine($"  {e} D:\\docs --analyze                      先逐节探测，看末尾是不是空白节");
        sb.AppendLine($"  {e} D:\\docs --all -n                       先干跑看一眼将做什么");
        sb.AppendLine($"  {e} D:\\docs --all                          批量删页眉页脚、空白末节、属性");
        sb.AppendLine($"  {e} D:\\docs --last-section --force         不管有没有内容，硬删最后一节");
        sb.AppendLine($"  {e} D:\\docs --privacy --orphans            再清一遍隐私残留和多余媒体");
        sb.AppendLine($"  {e} D:\\docs --privacy --track-changes      连修订痕迹和批注（作者、时间）一起清掉");
        sb.AppendLine($"  {e} D:\\docs --watermark=某公司,brand.example.com -n   先干跑：看哪些文档命中、会删几个水印对象");
        sb.AppendLine($"  {e} D:\\docs --watermark=某公司            命中就删水印艺术字和 0.05pt 埋点图");
        sb.AppendLine($"  {e} D:\\docs --watermark=某公司 --ocr D:\\software\\ppocr   连图片里的水印文字一起认");
        sb.AppendLine($"  {e} D:\\docs --image-hash <哈希1> --image-hash <哈希2> --algo MD5");
        sb.AppendLine($"  {e} D:\\docs --image-hash-file hashes.txt --algo SHA256");
        sb.AppendLine($"  {e} --hash-file logo.png --algo SHA256     计算参考哈希");
        sb.AppendLine($"  {e} D:\\docs --rename=\"\\s*\\((\\d+)\\)\" --rename-to=\"_v$1\" -n");
        sb.AppendLine($"  {e} D:\\docs --rename=\"\\s*\\((\\d+)\\)\" --rename-to=\"_v$1\"");
        sb.AppendLine($"  {e} D:\\docs --clean-name -n               先看文件名会被改成什么样");
        sb.AppendLine($"  {e} D:\\docs --clean-name                  去掉文件名里的 + 和空格");
        sb.AppendLine();
        sb.AppendLine("退出码: 0=全部成功  1=有文件处理失败  2=参数错误");

        try { Console.Out.Write(sb.ToString()); } catch { /* 无可用输出时忽略 */ }
    }

    #endregion
}
