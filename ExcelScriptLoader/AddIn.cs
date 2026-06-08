using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using WinFormsApp = System.Windows.Forms.Application;

namespace ExcelScriptLoader;

/// <summary>
/// Excel-DNA 插件入口，管理插件生命周期
/// </summary>
public class AddIn : IExcelAddIn
{
    // ==================== 静态构造：程序集解析钩子 ====================

    static AddIn()
    {
        // .NET 10 + Excel-DNA 自定义 AssemblyLoadContext 下，COM interop
        // 程序集 office.dll 无法自动解析。AppDomain.AssemblyResolve 是全局钩子，
        // 对默认 ALC 和 Excel-DNA 自定义 ALC 均生效。
        AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
        {
            var aname = new AssemblyName(args.Name);
            if (aname.Name != "office") return null;

            string path;
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "office.dll"),
                Path.Combine(Path.GetDirectoryName(typeof(AddIn).Assembly.Location)!, "office.dll"),
            };

            foreach (var p in candidates)
            {
                if (File.Exists(p))
                {
                    System.Diagnostics.Debug.WriteLine($"[AddIn] 加载 office.dll: {p}");
                    return Assembly.LoadFrom(p);
                }
            }
            return null;
        };
    }
    // ==================== 静态访问点（供 RibbonController 使用） ====================

    /// <summary>当前脚本引擎适配器实例</summary>
    public static ScriptEngineAdapter? Adapter { get; private set; }

    /// <summary>Excel Application 对象</summary>
    public static ExcelApplication? ExcelApp { get; private set; }

    /// <summary>当前工作簿的宏列表缓存</summary>
    public static List<ExcelMacro> CurrentMacros { get; private set; } = [];

    /// <summary>当前活动工作簿</summary>
    public static Workbook? CurrentWorkbook
    {
        get
        {
            try { return ExcelApp?.ActiveWorkbook; }
            catch { return null; }
        }
    }

    // ==================== IExcelAddIn 实现 ====================

    public void AutoOpen()
    {
        try
        {
            // 获取 Excel Application 引用
            ExcelApp = (ExcelApplication)ExcelDnaUtil.Application;

            // 初始化脚本引擎适配器
            Adapter = new ScriptEngineAdapter();
            Adapter.Initialize(ExcelApp);

            // 加载当前工作簿的宏（如果已有打开的工作簿）
            RefreshMacrosFromCurrentWorkbook();

            // 监听工作簿事件
            ExcelApp.WorkbookOpen += OnWorkbookOpen;
            ExcelApp.WorkbookActivate += OnWorkbookActivate;
            ExcelApp.WorkbookBeforeClose += OnWorkbookBeforeClose;

            System.Diagnostics.Debug.WriteLine("[AddIn] ExcelScriptLoader 已启动");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AddIn] AutoOpen 失败: {ex.Message}");
            System.Windows.Forms.MessageBox.Show(
                $"ExcelScriptLoader 启动失败:\n{ex.Message}",
                "插件错误",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    public void AutoClose()
    {
        try
        {
            if (ExcelApp != null)
            {
                ExcelApp.WorkbookOpen -= OnWorkbookOpen;
                ExcelApp.WorkbookActivate -= OnWorkbookActivate;
                ExcelApp.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            }

            // 清理引擎（释放脚本中的 COM 引用）
            Adapter?.Dispose();
            Adapter = null;
            CurrentMacros.Clear();

            // 释放 Excel Application COM 引用
            if (ExcelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                ExcelApp = null;
            }

            // 强制 GC 回收所有 RCW，确保 Excel 进程能退出
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            System.Diagnostics.Debug.WriteLine("[AddIn] ExcelScriptLoader 已关闭");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[AddIn] AutoClose 失败: {ex.Message}");
        }
    }

    // ==================== 事件处理 ====================

    /// <summary>
    /// 打开工作簿时加载宏（Excel 加载插件时已有工作簿打开，Activate 事件不触发）
    /// </summary>
    private void OnWorkbookOpen(Workbook wb)
    {
        try
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AddIn] WorkbookOpen: {wb.Name}, 尝试加载宏...");
            RefreshMacrosFromWorkbook(wb);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AddIn] WorkbookOpen 处理失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 工作簿切换时：先保存当前工作簿宏，再加载新工作簿宏
    /// </summary>
    private void OnWorkbookActivate(Workbook wb)
    {
        try
        {
            // 离开旧工作簿前保存宏（防止未保存修改丢失）
            var oldWb = CurrentWorkbook;
            if (oldWb != null && oldWb != wb && CurrentMacros.Count > 0)
            {
                MacroStorage.SaveMacros(oldWb, CurrentMacros);
            }

            // 加载新工作簿的宏
            RefreshMacrosFromWorkbook(wb);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AddIn] WorkbookActivate 处理失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 工作簿关闭前清理宏缓存
    /// </summary>
    private void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
    {
        try
        {
            // 如果关闭的是当前活动工作簿，清空宏列表
            if (CurrentWorkbook == wb || ExcelApp?.Workbooks.Count <= 1)
            {
                CurrentMacros.Clear();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AddIn] WorkbookBeforeClose 处理失败: {ex.Message}");
        }
    }

    // ==================== 公共方法 ====================

    /// <summary>
    /// 从当前活动工作簿刷新宏列表
    /// </summary>
    public static void RefreshMacrosFromCurrentWorkbook()
    {
        var wb = CurrentWorkbook;
        if (wb != null)
        {
            RefreshMacrosFromWorkbook(wb);
        }
        else
        {
            CurrentMacros = [];
        }
    }

    /// <summary>
    /// 从指定工作簿加载宏列表
    /// </summary>
    private static void RefreshMacrosFromWorkbook(Workbook wb)
    {
        try
        {
            CurrentMacros = MacroStorage.LoadMacros(wb);
            System.Diagnostics.Debug.WriteLine(
                $"[AddIn] 从工作簿 \"{wb.Name}\" 加载了 {CurrentMacros.Count} 个宏");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AddIn] 加载宏失败: {ex.Message}");
            CurrentMacros = [];
        }
    }

    /// <summary>
    /// 运行指定的宏
    /// </summary>
    public static ScriptResult RunMacro(ExcelMacro macro)
    {
        if (Adapter == null)
            return ScriptResult.Fail("脚本引擎未初始化");

        if (string.IsNullOrWhiteSpace(macro.ScriptCode))
            return ScriptResult.Fail("宏代码为空");

        try
        {
            return Adapter.Execute(macro.ScriptCode, macro.Name);
        }
        catch (Exception ex)
        {
            return ScriptResult.Fail($"执行异常: {ex.Message}");
        }
    }

    /// <summary>
    /// 保存当前宏列表到工作簿
    /// </summary>
    public static void SaveCurrentMacros()
    {
        var wb = CurrentWorkbook;
        if (wb == null)
            throw new InvalidOperationException("没有活动的工作簿");

        MacroStorage.SaveMacros(wb, CurrentMacros);
        System.Diagnostics.Debug.WriteLine(
            $"[AddIn] 保存了 {CurrentMacros.Count} 个宏到工作簿 \"{wb.Name}\"");
    }

    /// <summary>
    /// 延迟初始化（RibbonController 在 AutoOpen 失败时的回退方案）
    /// </summary>
    public static void LazyInitialize(ExcelApplication excelApp)
    {
        if (Adapter != null && ExcelApp != null)
            return; // 已初始化

        System.Diagnostics.Debug.WriteLine("[AddIn] 执行延迟初始化...");

        ExcelApp = excelApp;
        Adapter = new ScriptEngineAdapter();
        Adapter.Initialize(excelApp);

        // 刷新宏列表
        RefreshMacrosFromCurrentWorkbook();

        // 注册事件
        excelApp.WorkbookActivate += (Workbook wb) =>
        {
            try { RefreshMacrosFromWorkbook(wb); }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[AddIn] WorkbookActivate 失败: {ex.Message}");
            }
        };

        System.Diagnostics.Debug.WriteLine("[AddIn] 延迟初始化完成");
    }
}
