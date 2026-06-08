namespace ExcelScriptLoader.Dialogs;

/// <summary>
/// 宏编辑器对话框 — 模态
/// </summary>
public class MacroEditorDialog : Form
{
    private readonly ExcelMacro? _existingMacro;
    private readonly bool _isEditMode;
    private readonly bool _hasVsCode;

    private TextBox _txtName = null!;
    private TextBox _txtDescription = null!;
    private TextBox _txtCode = null!;
    private Button _btnVsCode = null!;
    private Button _btnRun = null!;
    private Button _btnSave = null!;
    private Button _btnClose = null!;
    private Button _btnOutput = null!;
    private Panel _bottomPanel = null!;

    public MacroEditorDialog()
    {
        _isEditMode = false; _existingMacro = null;
        _hasVsCode = ExternalEditor.IsVsCodeAvailable;
        InitializeUI();
        _txtCode.Text = "import { excel } from \"excel\"\r\n\r\n";
        _txtCode.Select(_txtCode.Text.Length, 0);
    }

    public MacroEditorDialog(ExcelMacro existing)
    {
        _isEditMode = true; _existingMacro = existing;
        _hasVsCode = ExternalEditor.IsVsCodeAvailable;
        InitializeUI();
        LoadExistingData();
    }

    // ---- 布局 ----

    private void InitializeUI()
    {
        Text = (_isEditMode ? "编辑宏" : "新建宏") + (_hasVsCode ? " — VS Code 可用" : "");
        Size = new Size(720, 520);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = true; MaximizeBox = true;
        ShowInTaskbar = false;

        var y = 8;

        // 名称
        var lblName = new Label { Text = "宏名称：", Location = new Point(12, y + 4), Size = new Size(60, 23) };
        _txtName = new TextBox { Location = new Point(78, y), Size = new Size(220, 23), MaxLength = 100 };
        y += 30;

        // 描述
        var lblDesc = new Label { Text = "描述：", Location = new Point(12, y + 4), Size = new Size(60, 23) };
        _txtDescription = new TextBox { Location = new Point(78, y), Size = new Size(520, 23), MaxLength = 500 };
        y += 32;

        // 代码标签 + VS Code 按钮
        var lblCode = new Label { Text = "脚本代码：", Location = new Point(12, y + 4), Size = new Size(70, 23) };
        if (_hasVsCode)
        {
            _btnVsCode = new Button
            {
                Text = "VS Code",
                Location = new Point(550, y), Size = new Size(80, 28),
                BackColor = System.Drawing.Color.FromArgb(0x00, 0x7A, 0xCC),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
            };
            _btnVsCode.FlatAppearance.BorderSize = 0;
            _btnVsCode.Click += OnOpenVsCode;
            var lblHint = new Label
            {
                Text = "💡 支持语法高亮与自动补全",
                Location = new Point(78, y + 30), Size = new Size(280, 18),
                ForeColor = System.Drawing.Color.Gray, Font = new Font(Font.FontFamily, 8F),
            };
            Controls.AddRange([lblHint, _btnVsCode]);
        }
        y += _hasVsCode ? 50 : 24;

        // 代码编辑区 — 填充剩余空间
        _txtCode = new TextBox
        {
            Location = new Point(12, y),
            Size = new Size(680, 320),
            Multiline = true, ScrollBars = ScrollBars.Both,
            Font = new Font("Consolas", 10F),
            AcceptsTab = true, WordWrap = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
        };

        // 底部按钮面板 — 固定在底部
        _bottomPanel = new Panel
        {
            Height = 42,
            Dock = DockStyle.Bottom,
            Padding = new Padding(12, 6, 12, 4),
        };

        _btnRun     = MkBtn("运行", 0);
        _btnSave    = MkBtn("保存", 88);
        _btnOutput  = MkBtn("查看输出", 176);
        _btnClose   = MkBtn("关闭", 292);

        _btnRun.Click += OnRun;
        _btnSave.Click += OnSave;
        _btnOutput.Click += (_, _) => { OutputWindow.Instance.Show(); OutputWindow.Instance.Activate(); };
        _btnClose.Click += (_, _) => Close();

        _bottomPanel.Controls.AddRange([_btnRun, _btnSave, _btnOutput, _btnClose]);

        var controls = new List<Control> { lblName, _txtName, lblDesc, _txtDescription, lblCode, _txtCode, _bottomPanel };
        Controls.AddRange(controls.ToArray());

        AcceptButton = _btnSave;
        CancelButton = _btnClose;
    }

    private static Button MkBtn(string text, int x) => new()
    {
        Text = text, Location = new Point(x, 6), Size = new Size(80, 28),
    };

    private void LoadExistingData()
    {
        if (_existingMacro == null) return;
        _txtName.Text = _existingMacro.Name;
        _txtDescription.Text = _existingMacro.Description;
        _txtCode.Text = _existingMacro.ScriptCode;
    }

    // ---- VS Code ----
    private void OnOpenVsCode(object? s, EventArgs e)
    {
        var name = string.IsNullOrWhiteSpace(_txtName.Text) ? "untitled" : _txtName.Text.Trim();
        var result = ExternalEditor.EditInVsCode(_txtCode.Text, name);
        if (result != null) _txtCode.Text = result;
    }

    // ---- 事件 ----
    private void OnSave(object? s, EventArgs e) { SaveMacro(); }  // 保存后保持打开
    private void OnRun(object? s, EventArgs e)
    {
        if (!SaveMacro()) return;
        RunCurrent();
    }

    // ---- 逻辑 ----
    private bool SaveMacro()
    {
        if (!ValidateInput()) return false;
        var macro = _existingMacro ?? new ExcelMacro();
        macro.Name = _txtName.Text.Trim();
        macro.Description = _txtDescription.Text.Trim();
        macro.ScriptCode = _txtCode.Text;
        macro.ModifiedAt = DateTime.UtcNow;

        if (!_isEditMode) AddIn.CurrentMacros.Add(macro);
        AddIn.SaveCurrentMacros();
        return true;
    }

    private bool ValidateInput()
    {
        var name = _txtName.Text.Trim();
        var err = ExcelMacro.ValidateName(name);
        if (err != null) { MessageBox.Show(err, "输入错误"); _txtName.Focus(); return false; }
        var wb = AddIn.CurrentWorkbook;
        if (wb != null && MacroStorage.NameExists(wb, name, _isEditMode ? _existingMacro?.Id : null))
        { MessageBox.Show($"名称 \"{name}\" 已存在", "名称冲突"); _txtName.Focus(); return false; }
        if (string.IsNullOrWhiteSpace(_txtCode.Text))
        { MessageBox.Show("代码不能为空", "输入错误"); _txtCode.Focus(); return false; }
        return true;
    }

    public static void RunToOutput(ExcelMacro macro)
    {
        var result = AddIn.RunMacro(macro);
        AppendOutput(macro.Name, result);
    }

    private void RunCurrent()
    {
        var macro = _existingMacro ?? new ExcelMacro { Name = _txtName.Text.Trim(), ScriptCode = _txtCode.Text };
        RunToOutput(macro);
    }

    private static void AppendOutput(string macroName, ScriptResult result)
    {
        var win = OutputWindow.Instance;
        if (!win.Visible) win.Show();
        OutputWindow.AppendSafe($"=== {macroName} {DateTime.Now:HH:mm:ss} ===\r\n");

        if (!string.IsNullOrEmpty(result.Output))
            OutputWindow.AppendSafe(result.Output);

        if (result.Success)
        {
            var ret = string.IsNullOrEmpty(result.ReturnValue) ? "" : $"  返回值: {result.ReturnValue}";
            OutputWindow.AppendSafe($"-- 完成{ret} ({result.ExecutionTimeMs}ms) --\r\n");
        }
        else
        {
            OutputWindow.AppendSafe($"-- 错误: {result.ErrorMessage} --\r\n");
        }
    }
}
