using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocxToolbox;

/// <summary>一条重命名计划的处置结果。</summary>
public enum RenameStatus
{
    /// <summary>可以改名，等用户确认。</summary>
    Ready,

    /// <summary>已经改完。</summary>
    Done,

    /// <summary>正则没匹配上，保持原样。</summary>
    NoMatch,

    /// <summary>匹配上了但替换结果和原来一样，没必要动。</summary>
    Unchanged,

    /// <summary>替换完名字空了。</summary>
    EmptyName,

    /// <summary>新名字 Windows 不接受（非法字符、尾部空格/句点、保留设备名）。</summary>
    InvalidName,

    /// <summary>目标已存在，或和同批次其他文件撞名。</summary>
    TargetExists,

    /// <summary>源文件在出计划时就不见了。</summary>
    SourceMissing,

    /// <summary>执行阶段失败。</summary>
    Failed,
}

/// <summary>
/// 一条重命名计划：原路径 → 目标路径，以及为什么这么定。
/// 只有 Status 为 Ready 的条目才会被真正执行。
/// </summary>
public sealed class RenameItem
{
    public RenameItem(string sourcePath) => SourcePath = sourcePath;

    /// <summary>原始完整路径。</summary>
    public string SourcePath { get; }

    /// <summary>原始文件名（含扩展名）。</summary>
    public string OldName => Path.GetFileName(SourcePath);

    /// <summary>新文件名（含扩展名）；不需要改时与 OldName 相同。</summary>
    public string NewName { get; internal set; } = "";

    /// <summary>新的完整路径（同目录）。</summary>
    public string TargetPath { get; internal set; } = "";

    public RenameStatus Status { get; internal set; } = RenameStatus.NoMatch;

    /// <summary>状态补充说明，一般只在出问题时有内容。</summary>
    public string Note { get; internal set; } = "";

    /// <summary>预览对话框里用户有没有勾上。</summary>
    public bool Selected { get; set; }

    /// <summary>状态对应的界面文案。</summary>
    public string StatusText => Status switch
    {
        RenameStatus.Ready => "将重命名",
        RenameStatus.Done => "已重命名",
        RenameStatus.NoMatch => "未匹配",
        RenameStatus.Unchanged => "无变化",
        RenameStatus.EmptyName => "替换后为空",
        RenameStatus.SourceMissing => "源文件不存在",
        _ => Note,
    };

    /// <summary>出问题时是否用醒目的颜色显示。</summary>
    public bool IsProblem =>
        Status is RenameStatus.InvalidName or RenameStatus.TargetExists
                 or RenameStatus.SourceMissing or RenameStatus.Failed or RenameStatus.EmptyName;
}

/// <summary>
/// 一条重命名规则：匹配用的正则 + 替换模板。
/// 正则只在 TryCreate 里编译一次，编译失败会给出可直接显示给用户的错误文本。
/// </summary>
public sealed class RenameRule
{
    /// <summary>单次匹配的超时上限，防止灾难性回溯把界面卡死。</summary>
    private static readonly TimeSpan MatchTimeout = TimeSpan.FromSeconds(5);

    private RenameRule(string pattern, string replacement, bool ignoreCase, Regex regex)
    {
        Pattern = pattern;
        Replacement = replacement;
        IgnoreCase = ignoreCase;
        Regex = regex;
    }

    public string Pattern { get; }
    public string Replacement { get; }
    public bool IgnoreCase { get; }
    public Regex Regex { get; }

    /// <summary>编译规则。返回 null 时 error 是可直接展示给用户的中文说明。</summary>
    public static RenameRule? TryCreate(string? pattern, string? replacement, bool ignoreCase, out string error)
    {
        error = "";

        // 只拦「真的没填」。不能连纯空格一起拦：单个空格就是一条合法规则（删掉所有空格），
        // 之前用 IsNullOrWhiteSpace 会把它误判成未填写，用户会以为「空格匹配不到」。
        if (string.IsNullOrEmpty(pattern))
        {
            error = "请先填写匹配用的正则表达式。";
            return null;
        }

        var options = RegexOptions.CultureInvariant;
        if (ignoreCase) options |= RegexOptions.IgnoreCase;

        Regex regex;
        try
        {
            regex = new Regex(pattern, options, MatchTimeout);
        }
        catch (ArgumentException ex)
        {
            // .NET 的报错是英文的，补一句中文提示：最常见的坑是忘了转义元字符（比如直接写 +）
            error = "正则表达式无效: " + ex.Message
                    + "（提示：+ * ? . ( ) [ ] { } ^ $ | \\ 在正则里有特殊含义，"
                    + "要匹配这些字符本身，请在前面加反斜杠，例如 \\+ 表示一个加号；"
                    + "或者把它们放进方括号里，例如 [+] 也表示一个加号）";
            return null;
        }

        return new RenameRule(pattern, replacement ?? "", ignoreCase, regex);
    }
}

/// <summary>
/// 内置的文件名快捷清理规则。GUI 和 CLI 都从这里取，避免两边各写一份正则走样。
/// </summary>
public static class NameCleanup
{
    /// <summary>
    /// 「去掉文件名里的 + 和空格」：把连续的空格/加号一并删掉。
    /// \s 在 .NET 里已覆盖 Unicode 空白（含全角空格 \u3000、制表符），不必单列；
    /// 全角加号 ＋ 单列出来，从网上下载的文件名两种加号都会出现。
    /// </summary>
    public const string Pattern = @"[+\s＋]+";
}

/// <summary>
/// 按正则批量重命名文件。只改文件名、不打开文档包，因此不依赖 OpenXml。
/// 核心约定：作用范围是「主文件名」，扩展名原样保留，不会被正则碰。
/// </summary>
public static class RenamePlanner
{
    /// <summary>Windows 保留设备名，用作文件名会创建失败。</summary>
    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    };

    /// <summary>
    /// 给整批文件出重命名计划。必须整批一起算：重名判断要跨文件看，
    /// 否则 A→B、C→B 这种会在执行阶段才撞车。
    /// </summary>
    public static List<RenameItem> BuildPlan(IEnumerable<string> files, RenameRule rule)
    {
        var items = files.Select(f => new RenameItem(f)).ToList();

        // 本批次的源路径全集：用来判断「新名字撞上的是不是本批次里即将被改走的文件」
        var sources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // 已被本批次占用的目标路径
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var it in items) sources.Add(it.SourcePath);

        foreach (var it in items)
        {
            it.NewName = it.OldName;
            it.TargetPath = it.SourcePath;

            if (!File.Exists(it.SourcePath))
            {
                it.Status = RenameStatus.SourceMissing;
                continue;
            }

            // 正则只作用在主文件名上，扩展名原样保留
            string stem = Path.GetFileNameWithoutExtension(it.OldName);
            string ext = Path.GetExtension(it.OldName);

            string newStem;
            try
            {
                if (!rule.Regex.IsMatch(stem))
                {
                    it.Status = RenameStatus.NoMatch;
                    continue;
                }
                newStem = rule.Regex.Replace(stem, rule.Replacement);
            }
            catch (RegexMatchTimeoutException)
            {
                it.Status = RenameStatus.Failed;
                it.Note = "匹配超时，正则可能过于复杂";
                continue;
            }

            if (newStem.Length == 0 || newStem.Trim().Length == 0)
            {
                it.Status = RenameStatus.EmptyName;
                it.Note = "替换后新名为空";
                continue;
            }

            // Windows 会自动砍掉结尾的空格和句点，那样实际文件名和预期不一致，直接拦掉
            if (newStem != newStem.TrimEnd(' ', '.'))
            {
                it.Status = RenameStatus.InvalidName;
                it.Note = "新名不能以空格或句点结尾";
                continue;
            }

            if (newStem.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                it.Status = RenameStatus.InvalidName;
                it.Note = "新名含有 Windows 不允许的字符";
                continue;
            }

            if (ReservedNames.Contains(newStem))
            {
                it.Status = RenameStatus.InvalidName;
                it.Note = "新名是 Windows 保留名（如 CON、NUL）";
                continue;
            }

            string newName = newStem + ext;
            string dir = Path.GetDirectoryName(it.SourcePath) ?? "";
            string target = dir.Length == 0 ? newName : Path.Combine(dir, newName);

            if (string.Equals(newName, it.OldName, StringComparison.OrdinalIgnoreCase))
            {
                it.Status = RenameStatus.Unchanged;
                continue;
            }

            if (File.Exists(target) || Directory.Exists(target))
            {
                it.Status = RenameStatus.TargetExists;
                it.Note = "目标文件已存在";
                continue;
            }

            // 同批次内两个文件改成同一个名字
            if (taken.Contains(target))
            {
                it.Status = RenameStatus.TargetExists;
                it.Note = "与本批次其他文件重名";
                continue;
            }

            // 撞上本批次里另一个待改名的文件的旧名：链式改名（A→B、B→C）执行顺序敏感，
            // 保守起见一律判冲突，让用户分两批做，避免中间态撞车
            if (sources.Contains(target))
            {
                it.Status = RenameStatus.TargetExists;
                it.Note = "与本批次待改名的文件冲突";
                continue;
            }

            it.NewName = newName;
            it.TargetPath = target;
            it.Status = RenameStatus.Ready;
            it.Selected = true;
            taken.Add(target);
        }

        return items;
    }

    /// <summary>执行计划里被勾上的条目，逐条 File.Move（同目录即重命名）。</summary>
    public static (int Ok, int Failed) Apply(IEnumerable<RenameItem> items, Action<string>? log)
    {
        int ok = 0;
        int failed = 0;

        foreach (var it in items)
        {
            if (it.Status == RenameStatus.Ready && it.Selected)
            {
                try
                {
                    File.Move(it.SourcePath, it.TargetPath, false);
                    it.Status = RenameStatus.Done;
                    ok++;
                    log?.Invoke($"[成功] 重命名: {it.OldName} → {it.NewName}");
                }
                catch (Exception ex)
                {
                    it.Status = RenameStatus.Failed;
                    it.Note = ex.Message;
                    failed++;
                    log?.Invoke($"[异常] 重命名失败: {it.OldName} - {ex.Message}");
                }
                continue;
            }

            // 被拦下的条目也说一句，否则用户只看到改了几个、不知道哪些没动
            if (it.IsProblem)
            {
                log?.Invoke($"[跳过] 重命名: {it.OldName} - {it.StatusText}");
            }
        }

        return (ok, failed);
    }

    /// <summary>汇总：总数 / 将改名 / 跳过 / 其中因冲突或非法名被拦下。</summary>
    public static (int Total, int Ready, int Skipped, int Blocked) Summarize(IEnumerable<RenameItem> items)
    {
        int total = 0, ready = 0, blocked = 0;

        foreach (var it in items)
        {
            total++;
            if (it.Status == RenameStatus.Ready && it.Selected) ready++;
            else if (it.IsProblem) blocked++;
        }

        return (total, ready, total - ready, blocked);
    }

    /// <summary>把执行完的映射（旧路径 → 新路径）取出来，供文件列表刷新用。</summary>
    public static Dictionary<string, string> RenamedMap(IEnumerable<RenameItem> items)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var it in items)
        {
            if (it.Status == RenameStatus.Done) map[it.SourcePath] = it.TargetPath;
        }
        return map;
    }
}
