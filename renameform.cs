using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace DocxToolbox;

/// <summary>
/// 重命名预览对话框：列出「原文件名 → 新文件名」和每条的处置状态，
/// 用户逐条勾选、点「应用重命名」后才会真正动文件。规则本身在这里只读。
/// </summary>
public sealed class RenamePreviewForm : Form
{
    private const int PadX = 12;
    private const int WidthInner = 780;

    // 只用两种强调色：问题行红色，可改名行保持常规色，避免整屏花掉
    private static readonly Color ProblemColor = Color.FromArgb(0xA3, 0x2D, 0x2D);
    private static readonly Color MutedColor = Color.FromArgb(0x5F, 0x5E, 0x5A);

    private readonly List<RenameItem> _items;
    private readonly ListView _list;
    private readonly Label _lblSummary;
    private readonly Button _btnApply;

    public RenamePreviewForm(List<RenameItem> items, RenameRule rule)
    {
        _items = items;

        Text = "重命名预览";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(WidthInner + PadX * 2, 520);

        var lblRule = new Label
        {
            Left = PadX, Top = 12, Width = WidthInner, Height = 32,
            Text = $"匹配: {rule.Pattern}　　替换为: {(rule.Replacement.Length == 0 ? "（空，即删掉匹配部分）" : rule.Replacement)}"
                   + $"　　忽略大小写: {(rule.IgnoreCase ? "是" : "否")}",
            AutoEllipsis = true,
        };

        _list = new ListView
        {
            Left = PadX, Top = 50, Width = WidthInner, Height = 388,
            View = View.Details,
            FullRowSelect = true,
            CheckBoxes = true,
            MultiSelect = true,
            HideSelection = false,
            ShowItemToolTips = true,
            GridLines = true,
            BackColor = Color.White,
        };
        _list.Columns.Add("原文件名", 300);
        _list.Columns.Add("新文件名", 300);
        _list.Columns.Add("状态", 160);

        foreach (var it in items)
        {
            string newText = it.Status == RenameStatus.Ready || it.IsProblem
                ? it.NewName
                : "（不变）";

            var row = new ListViewItem(it.OldName)
            {
                Tag = it,
                Checked = it.Status == RenameStatus.Ready,
                ToolTipText = it.Status == RenameStatus.Ready
                    ? $"{it.SourcePath}\n→ {it.TargetPath}"
                    : it.SourcePath,
            };
            row.SubItems.Add(newText);
            row.SubItems.Add(it.StatusText);

            if (it.IsProblem) row.ForeColor = ProblemColor;
            else if (it.Status is RenameStatus.NoMatch or RenameStatus.Unchanged) row.ForeColor = MutedColor;

            _list.Items.Add(row);
        }

        // 只有 Ready 的条目允许勾选；勾选状态写回 RenameItem.Selected
        _list.ItemCheck += (_, e) =>
        {
            var it = (RenameItem)_list.Items[e.Index].Tag!;
            if (it.Status != RenameStatus.Ready)
            {
                e.NewValue = e.CurrentValue;
                return;
            }
            it.Selected = e.NewValue == CheckState.Checked;
            // ItemCheck 触发时 Checked 还没落定，推迟一拍再刷新统计
            BeginInvoke(new Action(UpdateSummary));
        };

        _lblSummary = new Label { Left = PadX, Top = 448, Width = 560, Height = 20 };

        _btnApply = new Button
        {
            Text = "应用重命名", Left = 700, Top = 442, Width = 92, Height = 28,
            DialogResult = DialogResult.OK,
        };

        var btnCancel = new Button
        {
            Text = "取消", Left = 604, Top = 442, Width = 84, Height = 28,
            DialogResult = DialogResult.Cancel,
        };

        Controls.Add(lblRule);
        Controls.Add(_list);
        Controls.Add(_lblSummary);
        Controls.Add(btnCancel);
        Controls.Add(_btnApply);

        AcceptButton = _btnApply;
        CancelButton = btnCancel;

        UpdateSummary();
    }

    private void UpdateSummary()
    {
        var (total, ready, skipped, blocked) = RenamePlanner.Summarize(_items);

        _lblSummary.Text = $"共 {total} 个，将改 {ready} 个，跳过 {skipped} 个"
                           + (blocked > 0 ? $"（其中 {blocked} 个因冲突或非法名被拦下）" : "");
        _btnApply.Enabled = ready > 0;
    }
}
