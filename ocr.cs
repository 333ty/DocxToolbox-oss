using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace DocxToolbox;

/// <summary>
/// 本机 ppocr.exe（PaddleOCR 编译版）的调用封装：给一个图片目录，返回「文件名 → 识别文字」。
/// 必须一次进程跑完整个目录——单张调用每次都要重新加载模型（约 5 秒），批量才是可用速度。
/// </summary>
public static class LocalOcr
{
    /// <summary>没有显式指定目录时，按顺序试这几个位置。</summary>
    private static readonly string[] CandidateHomes =
    {
        @"D:\software\ppocr",
        @"C:\software\ppocr",
        Path.Combine(AppContext.BaseDirectory, "ppocr"),
    };

    /// <summary>把用户给的路径（OCR 主目录或 ppocr.exe 本身）解析成主目录；目录不存在返回 null。</summary>
    public static string? ResolveHome(string? hint)
    {
        if (!string.IsNullOrWhiteSpace(hint))
        {
            string p = hint.Trim().Trim('"');
            if (File.Exists(p)) return Path.GetDirectoryName(p);
            return Directory.Exists(p) ? p : null;
        }

        foreach (string home in CandidateHomes)
        {
            if (File.Exists(Path.Combine(home, "ppocr.exe"))) return home;
        }
        return null;
    }

    /// <summary>
    /// 识别一个目录下的所有图片。返回「不含扩展名的文件名 → 该图全部文字（换行拼在一起）」。
    /// 失败或超时返回已解析到的部分结果，不抛异常——OCR 只是加分项，不该让整份文档处理失败。
    /// </summary>
    public static Dictionary<string, string> Recognize(string home, string imageDir, TimeSpan timeout,
                                                        Action<string>? log)
    {
        var results = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        string exe = Path.Combine(home, "ppocr.exe");
        if (!File.Exists(exe))
        {
            log?.Invoke($"找不到 OCR 程序: {exe}");
            return results;
        }

        string models = Path.Combine(home, "models");
        string dict = Path.Combine(home, "ppocr", "utils", "ppocr_keys_v1.txt");

        var psi = new ProcessStartInfo
        {
            FileName = exe,
            WorkingDirectory = home,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            // ppocr 自己按 UTF-8 打印中文；不设这一行，中文会变成乱码进不来
            StandardOutputEncoding = new UTF8Encoding(false),
            StandardErrorEncoding = new UTF8Encoding(false),
        };

        // 用 ArgumentList 逐个给，避免路径里有空格时手写引号出错
        psi.ArgumentList.Add("--det_model_dir=" + Path.Combine(models, "ch_PP-OCRv4_det_infer"));
        psi.ArgumentList.Add("--rec_model_dir=" + Path.Combine(models, "ch_PP-OCRv4_rec_infer"));
        psi.ArgumentList.Add("--cls_model_dir=" + Path.Combine(models, "ch_ppocr_mobile_v2.0_cls_infer"));
        if (File.Exists(dict)) psi.ArgumentList.Add("--rec_char_dict_path=" + dict);
        psi.ArgumentList.Add("--use_angle_cls=false");
        psi.ArgumentList.Add("--visualize=false");   // 不要可视化图，省时间也省磁盘
        psi.ArgumentList.Add("--image_dir=" + imageDir);

        var stdout = new StringBuilder();
        var stderr = new StringBuilder();

        try
        {
            using var proc = new Process { StartInfo = psi };
            proc.OutputDataReceived += (_, e) => { if (e.Data != null) lock (stdout) stdout.AppendLine(e.Data); };
            proc.ErrorDataReceived += (_, e) => { if (e.Data != null) lock (stderr) stderr.AppendLine(e.Data); };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();

            if (!proc.WaitForExit((int)timeout.TotalMilliseconds))
            {
                try { proc.Kill(entireProcessTree: true); } catch { /* 已经退出 */ }
                log?.Invoke($"OCR 超时（{timeout.TotalSeconds:0} 秒），已用已识别到的部分");
            }
            else
            {
                proc.WaitForExit();   // 等异步输出读完，否则末尾几行会丢
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"OCR 启动失败: {ex.Message}");
            return results;
        }

        Parse(stdout.ToString(), results);

        if (results.Count == 0 && stderr.Length > 0)
        {
            string first = stderr.ToString().Split('\n')[0].Trim();
            if (first.Length > 0) log?.Invoke("OCR 输出为空: " + first);
        }

        return results;
    }

    /// <summary>
    /// 解析 ppocr 的标准输出。格式：
    ///   predict img: C:\tmp\img0001.png
    ///   0  det boxes: [[...]] rec text: 某公司 rec score: 0.99
    /// </summary>
    private static void Parse(string output, Dictionary<string, string> results)
    {
        var imgLine = new Regex(@"predict img:\s*(?<p>.+?)\s*$", RegexOptions.Compiled);
        var textPart = new Regex(@"rec text:\s*(?<t>.*?)\s*rec score:", RegexOptions.Compiled);

        string? current = null;
        var builder = new StringBuilder();

        void Flush()
        {
            if (current == null) return;
            results[current] = builder.ToString().Trim();
            builder.Clear();
        }

        foreach (string raw in output.Split('\n'))
        {
            string line = raw.TrimEnd('\r');
            if (line.Length == 0) continue;

            var m = imgLine.Match(line);
            if (m.Success)
            {
                Flush();
                current = Path.GetFileNameWithoutExtension(m.Groups["p"].Value.Trim());
                continue;
            }

            if (current == null) continue;

            var t = textPart.Match(line);
            if (t.Success)
            {
                string text = t.Groups["t"].Value.Trim();
                if (text.Length > 0) builder.Append(text).Append('\n');
            }
        }

        Flush();
    }
}
