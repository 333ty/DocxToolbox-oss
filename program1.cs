using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocxToolbox;

// ===== 新增：用户设置持久化模型 =====
public record AppSettings
{
    public string ImageHash { get; init; } = "";
    public string HashAlgorithm { get; init; } = "MD5";

    // ===== 文件名正则重命名：记住上次填的规则 =====
    public string RenamePattern { get; init; } = "";
    public string RenameReplacement { get; init; } = "";
    public bool RenameIgnoreCase { get; init; }

    // ===== 水印检测：记住触发词与 OCR 目录 =====
    public bool WatermarkEnabled { get; init; }
    public string WatermarkKeywords { get; init; } = "";
    public bool WatermarkDropAllImages { get; init; }
    public string OcrDir { get; init; } = "";
}

// ===== 新增：设置存储/读取辅助类 =====
public static class SettingsStore
{
    private const string AppFolderName = "DocxToolbox";

    // 旧版把设置存在 "%AppData%\WordTool"。改名后仍会读一次旧文件，
    // 避免用户已记住的哈希与重命名规则丢失；下一次保存即写入新目录。
    private static readonly string Dir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        AppFolderName);

    private static readonly string LegacyFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WordTool",
        "settings.json");

    private static readonly string FilePath = Path.Combine(Dir, "settings.json");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true
    };

    public static AppSettings Load()
    {
        try
        {
            var settingsPath = File.Exists(FilePath) ? FilePath : LegacyFilePath;
            if (!File.Exists(settingsPath)) return new AppSettings();
            var json = File.ReadAllText(settingsPath);
            if (string.IsNullOrWhiteSpace(json)) return new AppSettings();
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            // 读取失败时返回默认值，不影响主功能
            return new AppSettings();
        }
    }

    public static void Save(AppSettings settings)
    {
        try
        {
            Directory.CreateDirectory(Dir);
            var json = JsonSerializer.Serialize(settings, JsonOpts);
            File.WriteAllText(FilePath, json);
        }
        catch
        {
            // 写入失败时静默忽略，不影响主功能
        }
    }
}

class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        // 带参数时进入命令行模式，无参数时启动图形界面
        if (args is { Length: > 0 })
        {
            return Cli.Run(args);
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
        return 0;
    }
}

public class MainForm : Form
{
    private ListBox fileListBox;
    private Button btnSelectFiles;
    private Button btnSelectFolder;
    private System.Windows.Forms.CheckBox chkRemoveHF;
    private System.Windows.Forms.CheckBox chkRemoveSec;
    private System.Windows.Forms.CheckBox chkRemoveProps;
    private System.Windows.Forms.CheckBox chkRemoveImageByHash;
    private System.Windows.Forms.CheckBox chkPrivacy;
    private System.Windows.Forms.CheckBox chkOrphans;
    private System.Windows.Forms.CheckBox chkTrackChanges;
    private System.Windows.Forms.CheckBox chkWatermark;
    private TextBox txtWatermarkKeywords;
    private System.Windows.Forms.CheckBox chkDropAllImages;
    private TextBox txtOcrDir;
    private ComboBox cmbHashAlgorithm;
    private TextBox txtImageHash;
    private Button btnCalcHash;
    private TextBox txtRenamePattern;
    private TextBox txtRenameReplacement;
    private System.Windows.Forms.CheckBox chkRenameIgnoreCase;
    private Button btnCleanName;
    private Button btnRename;
    private Button btnAnalyze;
    private Button btnExecute;
    private Button btnCancel;
    private ProgressBar progressBar;
    private Label lblStatus;
    private TextBox logBox;

    // ===== 新增：防止构造期内反向保存覆盖旧值 =====
    private bool _initialized = false;

    // ===== 新增：后台处理的取消源 =====
    private CancellationTokenSource? _cts;

    public MainForm()
    {
        this.Text = "Word 批处理工具箱";
        this.Width = 716;
        this.Height = 836;
        this.StartPosition = FormStartPosition.CenterScreen;

        btnSelectFiles = new Button { Text = "选择文件 (可多选)", Left = 20, Top = 20, Width = 290, Height = 30 };
        btnSelectFiles.Click += BtnSelectFiles_Click;

        btnSelectFolder = new Button { Text = "选择文件夹 (含子文件夹)", Left = 320, Top = 20, Width = 360, Height = 30 };
        btnSelectFolder.Click += BtnSelectFolder_Click;

        fileListBox = new ListBox { Left = 20, Top = 58, Width = 660, Height = 120, SelectionMode = SelectionMode.MultiExtended };

        chkRemoveHF       = new System.Windows.Forms.CheckBox { Text = "删除页眉页脚", Left = 20,  Top = 190, Width = 130 };
        chkRemoveSec      = new System.Windows.Forms.CheckBox { Text = "删除末尾空白节", Left = 155, Top = 190, Width = 145 };
        chkRemoveProps    = new System.Windows.Forms.CheckBox { Text = "删除属性信息", Left = 305, Top = 190, Width = 120 };

        btnAnalyze = new Button { Text = "探测分节结构", Left = 430, Top = 186, Width = 120, Height = 28 };
        btnAnalyze.Click += BtnAnalyze_Click;

        chkRemoveImageByHash = new System.Windows.Forms.CheckBox { Text = "删除匹配哈希图片", Left = 20,  Top = 220, Width = 150 };
        cmbHashAlgorithm = new ComboBox { Left = 175, Top = 218, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
        cmbHashAlgorithm.Items.AddRange(new object[] { "MD5", "SHA1", "SHA256" });
        cmbHashAlgorithm.SelectedIndex = 0;

        txtImageHash = new TextBox
        {
            Left = 260, Top = 218, Width = 300,
            PlaceholderText = "粘贴图片哈希，多个用空格/逗号/换行分隔"
        };

        btnCalcHash = new Button { Text = "计算参考哈希", Left = 565, Top = 216, Width = 115, Height = 25 };
        btnCalcHash.Click += BtnCalcHash_Click;

        chkPrivacy = new System.Windows.Forms.CheckBox { Text = "清除隐私残留 (rsid/自定义数据/外链)", Left = 20, Top = 252, Width = 320 };
        chkOrphans = new System.Windows.Forms.CheckBox { Text = "清理未引用媒体", Left = 350, Top = 252, Width = 160 };

        // ===== 接受修订并清除批注：会改变文档可见内容，独立成项、不进 --all =====
        chkTrackChanges = new System.Windows.Forms.CheckBox
        {
            Text = "接受修订、清除批注（会改变可见内容）",
            Left = 20, Top = 284, Width = 640
        };

        // ===== 水印 / 来源检测：命中指定文字才动手，默认只清水印对象 =====
        chkWatermark = new System.Windows.Forms.CheckBox { Text = "水印检测", Left = 20, Top = 316, Width = 92 };
        txtWatermarkKeywords = new TextBox
        {
            Left = 116, Top = 316, Width = 280,
            PlaceholderText = "触发词，逗号分隔，如：某公司,example.com"
        };
        chkDropAllImages = new System.Windows.Forms.CheckBox { Text = "命中删全部图", Left = 402, Top = 319, Width = 130 };
        txtOcrDir = new TextBox
        {
            Left = 536, Top = 316, Width = 144,
            PlaceholderText = "OCR 目录（可空）"
        };

        // ===== 文件名正则重命名：只改文件名，扩展名原样保留 =====
        txtRenamePattern = new TextBox
        {
            Left = 20, Top = 350, Width = 208,
            PlaceholderText = "文件名正则，如 \\s*\\((\\d+)\\)\\s*$"
        };

        txtRenameReplacement = new TextBox
        {
            Left = 236, Top = 350, Width = 150,
            PlaceholderText = "替换为（如 $1）"
        };

        chkRenameIgnoreCase = new System.Windows.Forms.CheckBox { Text = "忽略大小写", Left = 394, Top = 353, Width = 92 };

        // 一键套用内置规则（去 + 和空格）：省得为非技术场景手写正则
        btnCleanName = new Button { Text = "去 +/空格", Left = 490, Top = 348, Width = 92, Height = 25 };
        btnCleanName.Click += BtnCleanName_Click;

        btnRename = new Button { Text = "重命名…", Left = 588, Top = 348, Width = 92, Height = 25 };
        btnRename.Click += BtnRename_Click;

        btnExecute = new Button { Text = "开始处理", Left = 20, Top = 388, Width = 520, Height = 38 };
        btnExecute.Click += BtnExecute_Click;

        btnCancel = new Button { Text = "取消", Left = 550, Top = 388, Width = 130, Height = 38, Enabled = false };
        btnCancel.Click += BtnCancel_Click;

        progressBar = new ProgressBar { Left = 20, Top = 436, Width = 660, Height = 18, Minimum = 0, Maximum = 1 };
        lblStatus = new Label { Left = 20, Top = 460, Width = 660, Height = 18, Text = "就绪" };

        logBox = new TextBox
        {
            Left = 20, Top = 484, Width = 660, Height = 300,
            Multiline = true, ScrollBars = ScrollBars.Vertical,
            ReadOnly = true, BackColor = Color.White
        };

        this.Controls.Add(btnSelectFiles);
        this.Controls.Add(btnSelectFolder);
        this.Controls.Add(fileListBox);
        this.Controls.Add(chkRemoveHF);
        this.Controls.Add(chkRemoveSec);
        this.Controls.Add(chkRemoveProps);
        this.Controls.Add(btnAnalyze);
        this.Controls.Add(chkRemoveImageByHash);
        this.Controls.Add(cmbHashAlgorithm);
        this.Controls.Add(txtImageHash);
        this.Controls.Add(btnCalcHash);
        this.Controls.Add(chkPrivacy);
        this.Controls.Add(chkOrphans);
        this.Controls.Add(chkTrackChanges);
        this.Controls.Add(chkWatermark);
        this.Controls.Add(txtWatermarkKeywords);
        this.Controls.Add(chkDropAllImages);
        this.Controls.Add(txtOcrDir);
        this.Controls.Add(txtRenamePattern);
        this.Controls.Add(txtRenameReplacement);
        this.Controls.Add(chkRenameIgnoreCase);
        this.Controls.Add(btnCleanName);
        this.Controls.Add(btnRename);
        this.Controls.Add(btnExecute);
        this.Controls.Add(btnCancel);
        this.Controls.Add(progressBar);
        this.Controls.Add(lblStatus);
        this.Controls.Add(logBox);

        // ===== 新增：加载上一次保存的设置，并挂接自动保存事件 =====
        LoadSettings();
        _initialized = true;
        txtImageHash.TextChanged += TxtImageHash_TextChanged;
        cmbHashAlgorithm.SelectedIndexChanged += CmbHashAlgorithm_SelectedIndexChanged;
        txtRenamePattern.TextChanged += TxtRename_Changed;
        txtRenameReplacement.TextChanged += TxtRename_Changed;
        chkRenameIgnoreCase.CheckedChanged += TxtRename_Changed;

        chkWatermark.CheckedChanged += Watermark_Changed;
        txtWatermarkKeywords.TextChanged += Watermark_Changed;
        chkDropAllImages.CheckedChanged += Watermark_Changed;
        txtOcrDir.TextChanged += Watermark_Changed;

        // 处理中不允许直接关窗，避免临时文件留在磁盘上
        this.FormClosing += MainForm_FormClosing;
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_cts == null) return;

        e.Cancel = true;
        MessageBox.Show("正在处理文件，请先点「取消」等当前文件结束。", "提示",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ===== 新增：从磁盘加载上次设置并回填 UI =====
    private void LoadSettings()
    {
        var s = SettingsStore.Load();

        // 回填哈希算法下拉框
        var algo = string.IsNullOrEmpty(s.HashAlgorithm) ? "MD5" : s.HashAlgorithm;
        int idx = cmbHashAlgorithm.Items.IndexOf(algo);
        cmbHashAlgorithm.SelectedIndex = idx >= 0 ? idx : 0;

        // 回填哈希值文本框
        txtImageHash.Text = s.ImageHash ?? "";

        // 回填上次的重命名规则
        txtRenamePattern.Text = s.RenamePattern ?? "";
        txtRenameReplacement.Text = s.RenameReplacement ?? "";
        chkRenameIgnoreCase.Checked = s.RenameIgnoreCase;

        // 回填水印检测：触发词以用户上次填写的为准，程序不预置默认词
        chkWatermark.Checked = s.WatermarkEnabled;
        txtWatermarkKeywords.Text = s.WatermarkKeywords ?? "";
        chkDropAllImages.Checked = s.WatermarkDropAllImages;

        // OCR 目录：上次填过就用上次的；没填过则探测一次本机安装位置（探不到就留空＝不启用）
        txtOcrDir.Text = string.IsNullOrWhiteSpace(s.OcrDir)
            ? (LocalOcr.ResolveHome(null) ?? "")
            : s.OcrDir;
    }

    // ===== 新增：把当前 UI 状态写入磁盘 =====
    private void SaveSettings()
    {
        if (!_initialized) return;

        var s = new AppSettings
        {
            ImageHash = txtImageHash.Text ?? "",
            HashAlgorithm = cmbHashAlgorithm.SelectedItem?.ToString() ?? "MD5",
            RenamePattern = txtRenamePattern.Text ?? "",
            RenameReplacement = txtRenameReplacement.Text ?? "",
            RenameIgnoreCase = chkRenameIgnoreCase.Checked,
            WatermarkEnabled = chkWatermark.Checked,
            WatermarkKeywords = txtWatermarkKeywords.Text ?? "",
            WatermarkDropAllImages = chkDropAllImages.Checked,
            OcrDir = txtOcrDir.Text ?? ""
        };
        SettingsStore.Save(s);
    }

    private void TxtImageHash_TextChanged(object? sender, EventArgs e) => SaveSettings();

    private void CmbHashAlgorithm_SelectedIndexChanged(object? sender, EventArgs e) => SaveSettings();

    private void TxtRename_Changed(object? sender, EventArgs e) => SaveSettings();

    private void Watermark_Changed(object? sender, EventArgs e) => SaveSettings();

    /// <summary>处理期间的界面状态切换。</summary>
    private void SetRunning(bool running)
    {
        btnExecute.Enabled = !running;
        btnCancel.Enabled = running;
        btnSelectFiles.Enabled = !running;
        btnSelectFolder.Enabled = !running;
        btnAnalyze.Enabled = !running;
        btnCalcHash.Enabled = !running;
        fileListBox.Enabled = !running;
        chkRemoveHF.Enabled = !running;
        chkRemoveSec.Enabled = !running;
        chkRemoveProps.Enabled = !running;
        chkRemoveImageByHash.Enabled = !running;
        chkPrivacy.Enabled = !running;
        chkOrphans.Enabled = !running;
        chkTrackChanges.Enabled = !running;
        chkWatermark.Enabled = !running;
        txtWatermarkKeywords.Enabled = !running;
        chkDropAllImages.Enabled = !running;
        txtOcrDir.Enabled = !running;
        cmbHashAlgorithm.Enabled = !running;
        txtImageHash.Enabled = !running;
        txtRenamePattern.Enabled = !running;
        txtRenameReplacement.Enabled = !running;
        chkRenameIgnoreCase.Enabled = !running;
        btnCleanName.Enabled = !running;
        btnRename.Enabled = !running;
        UseWaitCursor = running;
    }

    private void BtnCancel_Click(object? sender, EventArgs e)
    {
        if (_cts == null) return;
        _cts.Cancel();
        btnCancel.Enabled = false;
        Log("[取消] 已请求取消，当前文件处理完就停。");
    }

    private void BtnSelectFiles_Click(object? sender, EventArgs e)
    {
        using (OpenFileDialog ofd = new OpenFileDialog())
        {
            ofd.Multiselect = true;
            ofd.Filter = "Word 文档 (*.docx)|*.docx";
            ofd.Title = "请选择需要处理的 Word 文件（按住 Ctrl 多选）";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fileListBox.Items.Clear();
                foreach (string file in ofd.FileNames)
                {
                    fileListBox.Items.Add(file);
                }
                Log($"已选择 {ofd.FileNames.Length} 个文件。");
            }
        }
    }

    private void BtnSelectFolder_Click(object? sender, EventArgs e)
    {
        using (FolderBrowserDialog fbd = new FolderBrowserDialog())
        {
            fbd.Description = "请选择文件夹（将递归处理其中所有子文件夹内的 docx）";
            fbd.ShowNewFolderButton = false;

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 递归扫描所有子文件夹；IgnoreInaccessible 避免无权限目录导致整体失败
                    var opts = new EnumerationOptions
                    {
                        RecurseSubdirectories = true,
                        IgnoreInaccessible = true,
                        MatchCasing = MatchCasing.CaseInsensitive
                    };

                    string[] files = Directory.EnumerateFiles(fbd.SelectedPath, "*.docx", opts)
                                              .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                                              .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                                              .ToArray();

                    if (files.Length == 0)
                    {
                        MessageBox.Show("所选文件夹（含子文件夹）中没有找到 .docx 文件！", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 统计涉及了几个不同的文件夹，方便确认递归范围
                    int folderCount = files.Select(f => Path.GetDirectoryName(f)!.ToLowerInvariant())
                                           .Distinct()
                                           .Count();

                    fileListBox.Items.Clear();
                    foreach (string file in files)
                    {
                        fileListBox.Items.Add(file);
                    }
                    Log($"已递归加载 {files.Length} 个 docx 文件，分布在 {folderCount} 个文件夹中：{fbd.SelectedPath}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("访问文件夹失败: " + ex.Message, "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private void BtnCalcHash_Click(object? sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog();
        ofd.Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.emf;*.wmf)|"
                     + "*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.emf;*.wmf|所有文件|*.*";
        ofd.Title = "选择一个图片文件以计算其哈希值";
        if (ofd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                string algo = cmbHashAlgorithm.SelectedItem?.ToString() ?? "MD5";
                string hash = WordCleaner.ComputeFileHash(ofd.FileName, algo);
                txtImageHash.Text = hash;   // 赋值会自动触发 TextChanged 完成保存
                Log($"已计算 {Path.GetFileName(ofd.FileName)} 的 {algo} 哈希: {hash}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("计算哈希失败: " + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    /// <summary>一键套用内置的「去 + 和空格」规则，然后走同一条预览确认流程。</summary>
    private void BtnCleanName_Click(object? sender, EventArgs e)
    {
        // 只把规则填进输入框，仍然要靠预览对话框确认，绝不偷偷改文件
        txtRenamePattern.Text = NameCleanup.Pattern;
        txtRenameReplacement.Text = "";
        BtnRename_Click(sender, e);
    }

    /// <summary>按正则重命名文件名：先出计划、弹预览确认，确认后才动文件。</summary>
    private void BtnRename_Click(object? sender, EventArgs e)
    {
        var files = SelectedFiles();
        if (files.Length == 0)
        {
            MessageBox.Show("请先选择文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var rule = RenameRule.TryCreate(txtRenamePattern.Text, txtRenameReplacement.Text,
                                        chkRenameIgnoreCase.Checked, out string err);
        if (rule == null)
        {
            MessageBox.Show(err, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 整批一起出计划：重名、非法名这些判断要跨文件看
        var plan = RenamePlanner.BuildPlan(files, rule);

        using var dlg = new RenamePreviewForm(plan, rule);
        if (dlg.ShowDialog(this) != DialogResult.OK)
        {
            Log("[取消] 已取消重命名，没有改动任何文件名。");
            return;
        }

        logBox.Clear();
        var (ok, failed) = RenamePlanner.Apply(plan, Log);
        Log($"------ 重命名完成: 成功 {ok}，失败 {failed} ------");

        // 列表里的路径换成新名字，否则后续清理会指向已经不存在的旧路径
        var renamed = RenamePlanner.RenamedMap(plan);
        if (renamed.Count > 0)
        {
            for (int i = 0; i < fileListBox.Items.Count; i++)
            {
                string cur = fileListBox.Items[i]?.ToString() ?? "";
                if (cur.Length > 0 && renamed.TryGetValue(cur, out string? fresh))
                {
                    fileListBox.Items[i] = fresh;
                }
            }
            Log($"已更新文件列表中的 {renamed.Count} 条路径。");
        }
    }

    /// <summary>把列表里的文件路径取出来（UI 线程调用）。</summary>
    private string[] SelectedFiles()
        => fileListBox.Items.Cast<object>().Select(x => x.ToString() ?? "").Where(x => x.Length > 0).ToArray();

    private async void BtnAnalyze_Click(object? sender, EventArgs e)
    {
        var files = SelectedFiles();
        if (files.Length == 0)
        {
            MessageBox.Show("请先选择文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SetRunning(true);
        logBox.Clear();
        progressBar.Maximum = files.Length;
        progressBar.Value = 0;

        // Progress<T> 在 UI 线程创建，Report 的回调自动回到 UI 线程
        IProgress<string> log = new Progress<string>(Log);
        IProgress<int> tick = new Progress<int>(n =>
        {
            progressBar.Value = Math.Min(n, progressBar.Maximum);
            lblStatus.Text = $"探测中 {n} / {files.Length}";
        });

        try
        {
            await Task.Run(() =>
            {
                for (int i = 0; i < files.Length; i++)
                {
                    try
                    {
                        using (WordprocessingDocument doc = WordprocessingDocument.Open(files[i], false))
                        {
                            WordCleaner.AnalyzeSections(doc, files[i], m => log.Report(m));
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Report($"[异常] 分析失败: {Path.GetFileName(files[i])} - {ex.Message}");
                    }
                    tick.Report(i + 1);
                }
            });
        }
        catch (Exception ex)
        {
            Log($"[异常] 探测中断: {ex.Message}");
        }

        Log("------ 探测完成 ------");
        lblStatus.Text = $"探测完成，共 {files.Length} 个文件";
        SetRunning(false);
    }

    private async void BtnExecute_Click(object? sender, EventArgs e)
    {
        var files = SelectedFiles();
        if (files.Length == 0)
        {
            MessageBox.Show("请先选择文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!chkRemoveHF.Checked && !chkRemoveSec.Checked && !chkRemoveProps.Checked &&
            !chkRemoveImageByHash.Checked && !chkPrivacy.Checked && !chkOrphans.Checked &&
            !chkTrackChanges.Checked && !chkWatermark.Checked)
        {
            MessageBox.Show("请至少勾选一个要执行的功能！", "提示",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string algo = cmbHashAlgorithm.SelectedItem?.ToString() ?? "MD5";
        var imageHashes = new List<string>();

        if (chkRemoveImageByHash.Checked)
        {
            // 支持一行一个，或用逗号/分号/空格分隔多个哈希
            imageHashes.AddRange(txtImageHash.Text
                .Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(h => h.Trim().ToLowerInvariant()));

            if (imageHashes.Count == 0)
            {
                MessageBox.Show("请填写要删除的图片哈希值！", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int expected = algo switch { "SHA1" => 40, "SHA256" => 64, _ => 32 };
            string? bad = imageHashes.FirstOrDefault(h => h.Length != expected || !h.All(Uri.IsHexDigit));
            if (bad != null)
            {
                MessageBox.Show($"哈希值「{bad}」不是合法的 {algo} 值（应为 {expected} 位十六进制字符）。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        // 界面状态在这里一次性快照，运行期间改勾选也不会影响本次任务
        var request = new CleanRequest
        {
            Headers = chkRemoveHF.Checked,
            LastSection = chkRemoveSec.Checked,
            Props = chkRemoveProps.Checked,
            Privacy = chkPrivacy.Checked,
            Orphans = chkOrphans.Checked,
            TrackChanges = chkTrackChanges.Checked,
            Algo = algo,
        };
        request.ImageHashes.AddRange(imageHashes);

        // 水印检测：勾了才启用，触发词必须由用户填写，没有内置默认；
        // OCR 目录留空表示不对图片做识别，只按文字和图形属性判断
        if (chkWatermark.Checked)
        {
            var words = txtWatermarkKeywords.Text
                .Split(new[] { ',', ';', '，', '；', '、' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .Where(k => k.Length > 0)
                .ToList();

            if (words.Count == 0)
            {
                MessageBox.Show("已勾选「水印检测」，请填写触发词（多个词用逗号分隔）！", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            request.WatermarkKeywords.AddRange(words);
            request.WatermarkDropAllImages = chkDropAllImages.Checked;
            request.OcrDir = txtOcrDir.Text.Trim();
        }

        SetRunning(true);
        logBox.Clear();
        progressBar.Maximum = files.Length;
        progressBar.Value = 0;
        lblStatus.Text = $"0 / {files.Length}";

        IProgress<string> log = new Progress<string>(Log);
        IProgress<int> tick = new Progress<int>(n =>
        {
            progressBar.Value = Math.Min(n, progressBar.Maximum);
            lblStatus.Text = $"{n} / {files.Length}";
        });

        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        int succeeded = 0;
        int failed = 0;

        try
        {
            await Task.Run(() =>
            {
                for (int i = 0; i < files.Length; i++)
                {
                    if (token.IsCancellationRequested) break;

                    log.Report($"—— [{i + 1}/{files.Length}] {Path.GetFileName(files[i])} ——");
                    bool ok;
                    try
                    {
                        // 复制到临时文件 → 处理 → 校验 → 原子替换，失败时原文件不动
                        ok = WordCleaner.ProcessFile(files[i], request, m => log.Report(m));
                    }
                    catch (Exception ex)
                    {
                        log.Report($"[异常] 处理失败: {files[i]} - {ex.Message}");
                        ok = false;
                    }

                    if (ok) succeeded++; else failed++;
                    tick.Report(i + 1);
                }
            });
        }
        catch (Exception ex)
        {
            Log($"[异常] 处理中断: {ex.Message}");
        }
        finally
        {
            bool canceled = token.IsCancellationRequested;
            _cts.Dispose();
            _cts = null;
            SetRunning(false);

            Log($"------ 完成: 共 {files.Length} 个文件，成功 {succeeded}，失败 {failed}"
                + $"{(canceled ? "，其余已取消" : "")} ------");
            lblStatus.Text = canceled
                ? $"已取消，处理了 {succeeded + failed} / {files.Length}"
                : $"完成，共 {files.Length} 个（失败 {failed}）";
        }
    }

    private void Log(string message)
    {
        logBox.AppendText(message + Environment.NewLine);
        logBox.ScrollToCaret();
    }

}
