using ExcelDna.Integration.CustomUI;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelScriptLoader;

/// <summary>
/// 自定义 Ribbon 标签页控制器
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    // ==================== Ribbon XML ====================

    public override string GetCustomUI(string ribbonId)
    {
        Debug.WriteLine("[RibbonController] GetCustomUI called, ribbonId=" + ribbonId);

        return """
        <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
          <ribbon>
            <tabs>
              <tab id="tabScriptMacro" label="脚本宏">
                <group id="grpMacro" label="宏管理">
                  <button id="btnNewMacro"
                          label="新建宏"
                          imageMso="CreateReport"
                          size="large"
                          onAction="OnNewMacro"
                          screentip="创建一个新的脚本宏"
                          supertip="打开宏编辑器，输入名称和 C# 脚本代码并保存"/>
                  <button id="btnEditMacro"
                          label="编辑宏"
                          imageMso="EditDocument"
                          size="large"
                          onAction="OnEditMacro"
                          screentip="编辑已有的脚本宏"/>
                  <button id="btnRunMacro"
                          label="运行宏"
                          imageMso="PlayMacro"
                          size="large"
                          onAction="OnRunMacro"
                          screentip="选择并运行一个宏"/>
                  <button id="btnDeleteMacro"
                          label="删除宏"
                          imageMso="Delete"
                          size="large"
                          onAction="OnDeleteMacro"
                          screentip="删除选中的宏"/>
                  <button id="btnListMacros"
                          label="宏列表"
                          imageMso="ViewList"
                          size="large"
                          onAction="OnListMacros"
                          screentip="查看当前工作簿中的所有宏"/>
                  <button id="btnOutput"
                          label="输出"
                          imageMso="ZoomPrintPreviewExcelPage"
                          size="large"
                          onAction="OnShowOutput"
                          screentip="显示脚本输出窗口"/>
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """;
    }

    // ==================== 按钮回调（全部包裹 try-catch，防止 Excel 静默吞异常） ====================

    public void OnNewMacro(IRibbonControl control)
    {
        SafeExecute("新建宏", () =>
        {
            if (!EnsureReady()) return;
            var dialog = new Dialogs.MacroEditorDialog();
            dialog.ShowDialog();
        });
    }

    public void OnEditMacro(IRibbonControl control)
    {
        SafeExecute("编辑宏", () =>
        {
            if (!EnsureReady()) return;
            if (AddIn.CurrentMacros.Count == 0)
            {
                MessageBox.Show("当前工作簿中没有宏。请先新建一个宏。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var dialog = new Dialogs.MacroSelectorDialog(Dialogs.SelectorMode.Edit);
            dialog.ShowDialog();
        });
    }

    public void OnRunMacro(IRibbonControl control)
    {
        SafeExecute("运行宏", () =>
        {
            if (!EnsureReady()) return;
            if (AddIn.CurrentMacros.Count == 0)
            {
                MessageBox.Show("当前工作簿中没有宏。请先新建一个宏。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var dialog = new Dialogs.MacroSelectorDialog(Dialogs.SelectorMode.Run);
            dialog.ShowDialog();
        });
    }

    public void OnDeleteMacro(IRibbonControl control)
    {
        SafeExecute("删除宏", () =>
        {
            if (!EnsureReady()) return;
            if (AddIn.CurrentMacros.Count == 0)
            {
                MessageBox.Show("当前工作簿中没有宏。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var dialog = new Dialogs.MacroSelectorDialog(Dialogs.SelectorMode.Delete);
            dialog.ShowDialog();
        });
    }

    public void OnListMacros(IRibbonControl control)
    {
        SafeExecute("宏列表", () =>
        {
            if (!EnsureReady()) return;
            var dialog = new Dialogs.MacroSelectorDialog(Dialogs.SelectorMode.View);
            dialog.ShowDialog();
        });
    }

    public void OnShowOutput(IRibbonControl control)
    {
        SafeExecute("输出窗口", () =>
        {
            Dialogs.OutputWindow.Instance.Show();
            Dialogs.OutputWindow.Instance.Activate();
        });
    }

    // ==================== 辅助方法 ====================

    /// <summary>
    /// 安全检查：插件是否已就绪，工作簿是否已打开
    /// 返回 false 表示无法继续操作
    /// </summary>
    private static bool EnsureReady()
    {
        // 检查 AddIn 是否已初始化
        if (AddIn.ExcelApp == null)
        {
            Debug.WriteLine("[RibbonController] AddIn.ExcelApp 为 null，尝试延迟初始化...");

            // 尝试延迟初始化
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)
                    ExcelDna.Integration.ExcelDnaUtil.Application;
                AddIn.LazyInitialize(app);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[RibbonController] 延迟初始化失败: {ex.Message}");
                MessageBox.Show(
                    "插件未能正确初始化。请尝试重新加载插件。\n\n" +
                    $"错误: {ex.Message}",
                    "插件错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        // 检查是否有活动工作簿
        if (AddIn.CurrentWorkbook == null)
        {
            MessageBox.Show(
                "请先打开或新建一个工作簿。",
                "Excel Script",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return false;
        }

        return true;
    }

    /// <summary>
    /// 安全执行回调，捕获所有异常并用 MessageBox 显示
    /// （Excel Ribbon 会静默吞掉异常，所以必须自己 catch）
    /// </summary>
    private static void SafeExecute(string actionName, Action action)
    {
        try
        {
            Debug.WriteLine($"[RibbonController] 执行操作: {actionName}");
            action();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[RibbonController] {actionName} 失败: {ex}");
            MessageBox.Show(
                $"操作 [{actionName}] 失败:\n\n{ex.Message}\n\n" +
                $"详细信息:\n{ex}",
                "操作失败",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
