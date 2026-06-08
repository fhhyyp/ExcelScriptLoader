namespace ExcelScriptLoader.Dialogs;

/// <summary>
/// 独立输出窗口 — 脚本 print() 输出显示区域
/// 用法: OutputWindow.Instance.Show() / .Append(text) / .Clear()
/// </summary>
public class OutputWindow : Form
{
    private static OutputWindow? _instance;
    private TextBox _txtOutput = null!;
    private Button _btnClear = null!;

    public static OutputWindow Instance => _instance ??= new OutputWindow();

    private OutputWindow()
    {
        Text = "脚本输出";
        Size = new Size(620, 360);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = true;
        MaximizeBox = true;
        ShowInTaskbar = false;

        var y = 8;

        _btnClear = new Button
        {
            Text = "清空",
            Location = new Point(12, y),
            Size = new Size(60, 26),
        };
        _btnClear.Click += (_, _) => Clear();
        y += 32;

        _txtOutput = new TextBox
        {
            Location = new Point(12, y),
            Size = new Size(580, 270),
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 10F),
            BackColor = System.Drawing.Color.FromArgb(0x1E, 0x1E, 0x1E),
            ForeColor = System.Drawing.Color.FromArgb(0xD4, 0xD4, 0xD4),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom
                   | AnchorStyles.Left | AnchorStyles.Right,
        };

        Controls.AddRange([_btnClear, _txtOutput]);

        FormClosing += (_, e) =>
        {
            e.Cancel = true;
            Hide();
        };
    }

    public void Append(string text)
    {
        if (string.IsNullOrEmpty(text)) return;
        if (_txtOutput.InvokeRequired)
        {
            _txtOutput.Invoke(() => Append(text));
            return;
        }
        _txtOutput.AppendText(text);
        if (!text.EndsWith("\n")) _txtOutput.AppendText("\r\n");
        _txtOutput.ScrollToCaret();
    }

    public void Clear()
    {
        if (_txtOutput.InvokeRequired)
        {
            _txtOutput.Invoke(Clear);
            return;
        }
        _txtOutput.Clear();
    }

    public static void Log(string text) => Instance.Append(text);
    public static void LogLine(string text) => Instance.Append(text + "\r\n");

    /// <summary>线程安全追加</summary>
    public static void AppendSafe(string text)
    {
        var w = Instance;
        if (w._txtOutput.InvokeRequired)
            w._txtOutput.Invoke(() => w.Append(text));
        else
            w.Append(text);
    }
}
