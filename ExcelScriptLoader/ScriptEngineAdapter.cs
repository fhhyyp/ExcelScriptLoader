using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using ScriptLang;
using ScriptLang.Runtime;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;

namespace ExcelScriptLoader;

/// <summary>
/// 脚本执行结果
/// </summary>
public class ScriptResult
{
    public bool Success { get; init; }
    public string? ReturnValue { get; init; }
    public string? ErrorMessage { get; init; }
    public string Output { get; init; } = "";
    public long ExecutionTimeMs { get; init; }

    public static ScriptResult Ok(string? returnValue, string output, long elapsedMs) => new()
    {
        Success = true,
        ReturnValue = returnValue,
        Output = output,
        ExecutionTimeMs = elapsedMs,
    };

    public static ScriptResult Fail(string error, string output = "", long elapsedMs = 0) => new()
    {
        Success = false,
        ErrorMessage = error,
        Output = output,
        ExecutionTimeMs = elapsedMs,
    };
}

/// <summary>
/// SereinScript 引擎适配层
/// 管理引擎生命周期、Excel 对象注入、脚本执行
/// </summary>
public class ScriptEngineAdapter : IDisposable
{
    private ScriptEngine? _engine;
    private Scope? _excelScope;
    private ExcelApplication? _excelApp;
    private bool _disposed;

    /// <summary>
    /// 跟踪所有通过 DefineClrObject 注入到作用域的 COM 对象，
    /// 确保在 Dispose 时显式释放，防止 Excel 进程残留。
    /// </summary>
    private readonly List<object> _trackedComObjects = [];

    /// <summary>引擎实例（用于 Ribbon 等访问）</summary>
    public ScriptEngine? Engine => _engine;

    /// <summary>
    /// 初始化引擎并注入 Excel 对象上下文
    /// </summary>
    public void Initialize(ExcelApplication excelApp)
    {
        _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
        _engine = new ScriptEngine();

        // 注册包装类原型（必需！否则脚本无法调用 wrapper 方法）
        RegisterWrapperPrototypes();

        // 构建 Excel 对象作用域
        RefreshExcelScope();
    }

    /// <summary>
    /// 刷新 Excel 对象引用（每次执行前调用，确保 workbook/sheet/cell 是最新的）
    /// </summary>
    public void RefreshExcelContext()
    {
        if (_engine == null || _excelApp == null)
            throw new InvalidOperationException("引擎未初始化，请先调用 Initialize()");

        RefreshExcelScope();
    }

    /// <summary>
    /// 执行脚本代码
    /// </summary>
    /// <param name="scriptCode">脚本代码字符串</param>
    /// <param name="macroName">宏名称（用于错误报告）</param>
    /// <returns>执行结果</returns>
    public async Task<ScriptResult> ExecuteAsync(string scriptCode, string macroName)
    {
        if (_engine == null)
            return ScriptResult.Fail("脚本引擎未初始化");

        var sw = Stopwatch.StartNew();
        var output = new StringWriter();

        try
        {
            // 重定向 Console.Out → 捕获 print() 输出
            var originalOut = Console.Out;
            Console.SetOut(output);

            try
            {
                // 刷新 Excel 上下文
                RefreshExcelContext();

                // 注册 Excel 虚拟模块
                RegisterExcelModule();

                // 清空编译缓存
                _engine.ClearCache();

                // 创建任务并执行
                var scope = _excelScope ?? new Scope(_engine.GlobalScope);
                var task = _engine.CreateTaskFromSource(scriptCode, macroName, scope);

                var result = await task.RunAsync();

                sw.Stop();
                Console.SetOut(originalOut);

                var returnValue = result.IsNull ? null : result.ToString();
                return ScriptResult.Ok(returnValue, output.ToString(), sw.ElapsedMilliseconds);
            }
            catch
            {
                Console.SetOut(originalOut);
                throw;
            }
        }
        catch (Exception ex)
        {
            sw.Stop();
            var partialOutput = output.ToString();
            return ScriptResult.Fail(FormatException(ex, macroName), partialOutput, sw.ElapsedMilliseconds);
        }
    }

    /// <summary>
    /// 同步执行脚本代码（用于 UI 线程调用）
    /// 直接在当前线程同步执行，避免 Task.Run 导致的跨线程 COM 访问问题。
    /// ExecuteAsync 内部无真正异步操作（RunAsync 返回已完成 Task），
    /// 因此 GetAwaiter().GetResult() 不会造成死锁。
    /// </summary>
    public ScriptResult Execute(string scriptCode, string macroName)
    {
        try
        {
            return ExecuteAsync(scriptCode, macroName)
                .GetAwaiter().GetResult();
        }
        catch (AggregateException ae)
        {
            return ScriptResult.Fail(ae.InnerException?.Message ?? ae.Message);
        }
        catch (Exception ex)
        {
            return ScriptResult.Fail(ex.Message);
        }
    }

    // ==================== 内部方法 ====================

    /// <summary>
    /// 重建 Excel 对象作用域。
    /// 在创建新作用域之前，先显式释放旧作用域中的 COM 对象引用，
    /// 防止 RCW 累积导致 Excel 进程无法退出。
    /// </summary>
    private void RefreshExcelScope()
    {
        if (_engine == null || _excelApp == null) return;

        try
        {
            // 在创建新作用域前，显式释放旧作用域中跟踪的 COM 对象
            ReleaseTrackedComObjects();

            _excelScope = new Scope(_engine.GlobalScope);

            // 注入 Excel 对象（使用 ClrObjectValue 包装）
            // 同时跟踪这些 COM 对象，以便后续显式释放
            _excelScope.DefineClrObject("app", _excelApp);
            // _excelApp 已经在 AddIn 层单独管理，此处不重复跟踪


            try
            {
                var wb = _excelApp.ActiveWorkbook;
                _excelScope.DefineClrObject("workbook", wb);
                TrackComObject(wb);
            }
            catch { /* 无活动工作簿 */ }

            try
            {
                var sheet = _excelApp.ActiveSheet;
                _excelScope.DefineClrObject("sheet", sheet);
                TrackComObject(sheet);
            }
            catch { /* 无活动工作表 */ }

            try
            {
                var cell = _excelApp.ActiveCell;
                _excelScope.DefineClrObject("cell", cell);
                TrackComObject(cell);
            }
            catch { /* 无活动单元格 */ }

            try
            {
                var sel = _excelApp.Selection;
                _excelScope.DefineClrObject("selection", sel);
                TrackComObject(sel);
            }
            catch { /* 无选中区域 */ }


           

        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ScriptEngineAdapter] 刷新 Excel 作用域失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 跟踪 COM 对象，确保在 Dispose 时显式释放
    /// </summary>
    private void TrackComObject(object? comObj)
    {
        if (comObj != null && Marshal.IsComObject(comObj))
        {
            _trackedComObjects.Add(comObj);
        }
    }

    /// <summary>
    /// 显式释放所有跟踪的 COM 对象引用。
    /// 使用 FinalReleaseComObject 一次性释放 RCW 持有的所有 COM 引用计数，
    /// 不依赖 GC 回收与 RCW 终结器。
    /// </summary>
    private void ReleaseTrackedComObjects()
    {
        // 逆序释放（遵循 COM 释放惯例：先创建的后释放）
        for (int i = _trackedComObjects.Count - 1; i >= 0; i--)
        {
            try
            {
                var obj = _trackedComObjects[i];
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.FinalReleaseComObject(obj);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ScriptEngineAdapter] 释放 COM 对象失败: {ex.Message}");
            }
        }
        _trackedComObjects.Clear();
    }

    /// <summary>
    /// 注册所有 Excel 包装类的原型到 PrototypeManager
    /// 没有注册则脚本无法调用 wrapper 方法（如 sheet.table()）
    /// </summary>
    private void RegisterWrapperPrototypes()
    {
        // 仅注册顶层 ExcelModule 原型（其他对象用 ObjectValue+FunctionValue 模式）
        var excelProto = (IPrototype)new ExcelModule(_excelApp!);
        _engine?.PrototypeManager.Register(excelProto);
    }

    /// <summary>
    /// 注册 "excel" 虚拟模块
    /// 用法: import { excel } from "excel"
    /// </summary>
    private void RegisterExcelModule()
    {
        if (_engine == null || _excelApp == null) return;

        try
        {
            var excel = new ExcelModule(_excelApp);

            var properties = new Dictionary<string, Value>
            {
                { "excel", new ClrObjectValue(excel) },
                { "msgbox", new FunctionValue("msgbox", args =>
                {
                    if (args.Count > 0)
                        MessageBox.Show(args[0].AsString(), "Excel Script",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                }) },
                { "inputbox", new FunctionValue("inputbox", args =>
                {
                    var prompt = args.FirstOrDefault()?.AsString() ?? "";
                    return StringValue.Create(Interaction.InputBox(prompt, "Excel Script", ""));
                }) },
            };

            var module = new ObjectValue(properties);
            _engine.ImportResolver.RegisterBuiltinModule("excel", module);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ScriptEngineAdapter] 注册 excel 模块失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 格式化异常信息为可读文本
    /// </summary>
    private static string FormatException(Exception ex, string macroName)
    {
        // 解析错误比运行时错误提供更详细的位置信息
        var type = ex is ScriptLang.Parser.ParseException ? "语法错误" : "运行时错误";

        var message = $"[{type}] 宏 \"{macroName}\" 执行失败:\n\n{ex.Message}";

        // 如果有内部异常，追加信息
        if (ex.InnerException != null)
        {
            message += $"\n\n内部错误: {ex.InnerException.Message}";
        }

        return message;
    }

    /// <summary>
    /// 释放所有资源。
    /// 显式释放跟踪的 COM 对象，清空脚本引擎与作用域引用。
    /// 注意：_excelApp 不由此类释放（由 AddIn.AutoClose 统一管理）。
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // 1. 清空编译缓存
        _engine?.ClearCache();

        // 2. 显式释放所有跟踪的 COM 对象
        ReleaseTrackedComObjects();

        // 3. 清空作用域与引擎引用
        _excelScope = null;
        _engine = null;

        // 4. 不释放 _excelApp（由 AddIn.AutoClose 负责）
        _excelApp = null;
    }
}
