namespace ExcelScriptLoader.Dialogs;

/// <summary>
/// 宏选择器操作模式
/// </summary>
public enum SelectorMode
{
    View,   // 仅查看
    Run,    // 选择并运行
    Edit,   // 选择并编辑
    Delete, // 选择并删除
}

/// <summary>
/// 宏选择器对话框 — 列出所有宏，支持运行/编辑/删除
/// </summary>
public class MacroSelectorDialog : Form
{
    private readonly SelectorMode _mode;

    // 控件
    private ListView _listView = null!;
    private Button _btnPrimary = null!;
    private Button _btnEdit = null!;
    private Button _btnDelete = null!;
    private Button _btnClose = null!;
    private Label _lblTitle = null!;

    public MacroSelectorDialog(SelectorMode mode)
    {
        _mode = mode;
        InitializeUI();
        RefreshList();
    }

    private void InitializeUI()
    {
        var title = _mode switch
        {
            SelectorMode.Run => "运行宏",
            SelectorMode.Edit => "编辑宏",
            SelectorMode.Delete => "删除宏",
            _ => "宏列表",
        };

        Text = $"{title} — ({AddIn.CurrentWorkbook?.Name ?? "无工作簿"})";
        Size = new Size(600, 420);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;

        // 标题
        _lblTitle = new Label
        {
            Text = $"当前工作簿的宏（共 {AddIn.CurrentMacros.Count} 个）：",
            Location = new Point(12, 10),
            Size = new Size(560, 23),
            Font = new Font(Font.FontFamily, 9F, FontStyle.Bold),
        };

        // ListView
        _listView = new ListView
        {
            Location = new Point(12, 36),
            Size = new Size(560, 280),
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false,
            HeaderStyle = ColumnHeaderStyle.Nonclickable,
        };
        _listView.Columns.Add("宏名称", 150);
        _listView.Columns.Add("描述", 280);
        _listView.Columns.Add("修改时间", 126);
        _listView.DoubleClick += OnDoubleClick;

        // 按钮面板
        var btnY = 326;
        var btnX = 12;

        // 主操作按钮
        var primaryText = _mode switch
        {
            SelectorMode.Run => "运行",
            SelectorMode.Edit => "编辑",
            SelectorMode.Delete => "删除",
            _ => "关闭",
        };
        _btnPrimary = new Button
        {
            Text = primaryText,
            Location = new Point(btnX, btnY),
            Size = new Size(80, 30),
            Enabled = _mode != SelectorMode.View,
        };
        _btnPrimary.Click += OnPrimaryAction;
        btnX += 90;

        // 编辑按钮（除 Edit 模式外显示）
        if (_mode != SelectorMode.Edit)
        {
            _btnEdit = new Button
            {
                Text = "编辑",
                Location = new Point(btnX, btnY),
                Size = new Size(80, 30),
            };
            _btnEdit.Click += (_, _) => EditSelected();
            btnX += 90;
        }

        // 删除按钮（除 Delete 模式外显示）
        if (_mode != SelectorMode.Delete)
        {
            _btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(btnX, btnY),
                Size = new Size(80, 30),
            };
            _btnDelete.Click += (_, _) => DeleteSelected();
            btnX += 90;
        }

        // 关闭按钮
        _btnClose = new Button
        {
            Text = "关闭",
            Location = new Point(btnX + 20, btnY),
            Size = new Size(80, 30),
        };
        _btnClose.Click += (_, _) => Close();

        var controls = new List<Control> { _lblTitle, _listView, _btnPrimary, _btnClose };
        if (_btnEdit != null) controls.Add(_btnEdit);
        if (_btnDelete != null) controls.Add(_btnDelete);
        Controls.AddRange(controls.ToArray());

        AcceptButton = _btnPrimary;
        CancelButton = _btnClose;

        // 如果没有宏，禁用操作按钮
        if (AddIn.CurrentMacros.Count == 0)
        {
            _btnPrimary.Enabled = false;
            if (_btnEdit != null) _btnEdit.Enabled = false;
            if (_btnDelete != null) _btnDelete.Enabled = false;
        }
    }

    // ==================== 列表操作 ====================

    private void RefreshList()
    {
        _listView.Items.Clear();

        foreach (var macro in AddIn.CurrentMacros)
        {
            var item = new ListViewItem(macro.Name);
            item.SubItems.Add(macro.Description);
            item.SubItems.Add(macro.ModifiedAt.ToString("yyyy-MM-dd HH:mm"));
            item.Tag = macro;
            _listView.Items.Add(item);
        }

        _lblTitle.Text = $"当前工作簿的宏（共 {AddIn.CurrentMacros.Count} 个）：";
    }

    private ExcelMacro? GetSelectedMacro()
    {
        if (_listView.SelectedItems.Count == 0)
        {
            MessageBox.Show("请先选择一个宏。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return null;
        }
        return _listView.SelectedItems[0].Tag as ExcelMacro;
    }

    // ==================== 事件处理 ====================

    private void OnPrimaryAction(object? sender, EventArgs e)
    {
        var macro = GetSelectedMacro();
        if (macro == null) return;

        switch (_mode)
        {
            case SelectorMode.Run:
                RunSelected(macro);
                break;
            case SelectorMode.Edit:
                EditMacro(macro);
                break;
            case SelectorMode.Delete:
                DeleteMacro(macro);
                break;
        }
    }

    private void OnDoubleClick(object? sender, EventArgs e)
    {
        var macro = GetSelectedMacro();
        if (macro == null) return;

        switch (_mode)
        {
            case SelectorMode.Run:    RunSelected(macro); break;
            case SelectorMode.Delete: DeleteMacro(macro); break;
            default:                  EditMacro(macro); break;
        }
    }

    // ==================== 业务逻辑 ====================

    private void RunSelected(ExcelMacro macro)
    {
        MacroEditorDialog.RunToOutput(macro);
    }

    private void EditMacro(ExcelMacro macro)
    {
        var editor = new MacroEditorDialog(macro);
        editor.ShowDialog();
        RefreshList();
    }

    private void EditSelected()
    {
        var macro = GetSelectedMacro();
        if (macro != null)
        {
            EditMacro(macro);
        }
    }

    private void DeleteMacro(ExcelMacro macro)
    {
        var result = MessageBox.Show(
            $"确定要删除宏 \"{macro.Name}\" 吗？\n\n此操作不可撤销。",
            "确认删除",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

        if (result != DialogResult.Yes) return;

        // 从列表中移除
        AddIn.CurrentMacros.RemoveAll(m => m.Id == macro.Id);

        // 持久化
        AddIn.SaveCurrentMacros();

        RefreshList();

        if (_mode == SelectorMode.Delete)
        {
            Close();
        }
    }

    private void DeleteSelected()
    {
        var macro = GetSelectedMacro();
        if (macro != null)
        {
            DeleteMacro(macro);
        }
    }
}
