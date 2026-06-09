// ============================================================================
// ExcelModule — "excel" 模块顶层入口
// 通过 ScriptEngineAdapter 注册为内置模块: import { excel } from "excel"
//
// 架构:
//   本文件仅包含 [PrototypeExtension] 顶层 API（用户直接调用 excel.xxx）。
//   子对象工厂和工具方法位于 ExcelFactory.cs，通过 using static 导入。
//
// 模式:
//   - 顶层 excel.*  → PrototypeExtension（本文件）
//   - 子对象 API    → ObjectValue + FunctionValue（ExcelFactory.cs）
//   - 零反射、零原型依赖，直接委托调用 COM
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang;
using ScriptLang.Runtime;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using static ExcelScriptLoader.ExcelFactory;

namespace ExcelScriptLoader
{

    [PrototypeExtension(PushThis = true)]
    public partial class ExcelModule
    {
        private readonly ExcelApplication _app;

        public ExcelModule(ExcelApplication app) => _app = app;

        /// <summary>
        /// 类型匹配检查 — 脚本引擎用于判断 Value 是否属于此原型。
        /// </summary>
        public partial bool IsTarget(Value value) =>
            value is ClrObjectValue clr && clr.Value is ExcelModule;

        // ========================================================================
        // 属性
        // ========================================================================

        /// <summary>excel.active → 当前宿主工作簿</summary>
        [PrototypeProperty(Name = "active")]
        public static Value active(ExcelModule self)
        {
            try { return WorkbookToObject(self._app.ActiveWorkbook); }
            catch { return Value.Null; }
        }

        // ========================================================================
        // 基本访问器
        // ========================================================================

        /// <summary>excel.sheet(name?) → 获取工作表（无参=活动工作表，有参=按名称）</summary>
        [PrototypeFunction(Name = "sheet")]
        public static Value sheet(ExcelModule self, Value? name = null)
        {
            try
            {
                Worksheet ws;
                if (name is StringValue sv)
                    ws = (Worksheet)self._app.ActiveWorkbook?.Sheets[sv.Value];
                else
                    ws = (Worksheet)self._app.ActiveSheet;
                return WorksheetToObject(ws);
            }
            catch { return Value.Null; }
        }

        /// <summary>excel.cell() → 当前活动单元格</summary>
        [PrototypeFunction(Name = "cell")]
        public static Value cell(ExcelModule self)
        {
            try { return CellToObject(self._app.ActiveCell); }
            catch { return Value.Null; }
        }

        /// <summary>excel.selection() → 当前选中区域</summary>
        [PrototypeFunction(Name = "selection")]
        public static Value selection(ExcelModule self)
        {
            try
            {
                var rng = self._app.Selection as ExcelRange;
                return rng != null ? RangeToObject(rng) : Value.Null;
            }
            catch { return Value.Null; }
        }

        // ========================================================================
        // 文件操作
        // ========================================================================

        /// <summary>excel.open(path) → 打开外部工作簿</summary>
        [PrototypeFunction(Name = "open")]
        public static Value open(ExcelModule self, StringValue path)
        {
            try
            {
                var wb = self._app.Workbooks.Open(path.Value);
                return WorkbookToObject(wb);
            }
            catch { return Value.Null; }
        }

        /// <summary>excel.create() → 新建空白工作簿</summary>
        [PrototypeFunction(Name = "create")]
        public static Value create(ExcelModule self)
        {
            try { return WorkbookToObject(self._app.Workbooks.Add()); }
            catch { return Value.Null; }
        }

        /// <summary>
        /// excel.read(path, sheetName?) → 快捷读取。
        /// 打开→读取→关闭。不指定工作表时默认读取 Sheets[1]。
        /// </summary>
        [PrototypeFunction(Name = "read")]
        public static Value read(ExcelModule self, StringValue path, Value? sheetName = null)
        {
            try
            {
                var wb = self._app.Workbooks.Open(path.Value);
                try
                {
                    Worksheet ws;
                    if (sheetName is StringValue sn)
                        ws = (Worksheet)wb.Sheets[sn.Value];
                    else
                        ws = (Worksheet)wb.Sheets[1];
                    return ReadRange(ws.UsedRange);
                }
                finally { wb.Close(false); }
            }
            catch { return Value.Null; }
        }

        /// <summary>
        /// excel.write(path, data) → 快捷写入。
        /// 打开文件（不存在则创建）→写入 Sheets[1]→保存→关闭。
        /// </summary>
        [PrototypeFunction(Name = "write")]
        public static Value write(ExcelModule self, StringValue path, Value data)
        {
            try
            {
                var wb = self._app.Workbooks.Open(path.Value);
                if (wb == null) wb = self._app.Workbooks.Add();
                try
                {
                    var ws = (Worksheet)wb.Sheets[1];
                    WriteRange(ws.Range["A1"], data);
                    wb.Save();
                }
                finally { wb.Close(false); }
            }
            catch { }
            return Value.Null;
        }

        // ========================================================================
        // Application 控制 — 批量操作性能与行为开关
        // ========================================================================

        /// <summary>excel.screenUpdating() / (bool) → 获取/设置屏幕刷新</summary>
        [PrototypeFunction(Name = "screenUpdating")]
        public static Value screenUpdating(ExcelModule self, Value? value = null)
        {
            if (value == null || value is NullValue)
            {
                try { return BoolValue.Create(self._app.ScreenUpdating); }
                catch { return Value.Null; }
            }
            if (value is BoolValue bv)
            {
                try { self._app.ScreenUpdating = bv.Value; }
                catch { }
            }
            return Value.Null;
        }

        /// <summary>excel.displayAlerts() / (bool) → 获取/设置弹窗警告</summary>
        [PrototypeFunction(Name = "displayAlerts")]
        public static Value displayAlerts(ExcelModule self, Value? value = null)
        {
            if (value == null || value is NullValue)
            {
                try { return BoolValue.Create(self._app.DisplayAlerts); }
                catch { return Value.Null; }
            }
            if (value is BoolValue bv)
            {
                try { self._app.DisplayAlerts = bv.Value; }
                catch { }
            }
            return Value.Null;
        }

        /// <summary>excel.enableEvents() / (bool) → 获取/设置事件触发</summary>
        [PrototypeFunction(Name = "enableEvents")]
        public static Value enableEvents(ExcelModule self, Value? value = null)
        {
            if (value == null || value is NullValue)
            {
                try { return BoolValue.Create(self._app.EnableEvents); }
                catch { return Value.Null; }
            }
            if (value is BoolValue bv)
            {
                try { self._app.EnableEvents = bv.Value; }
                catch { }
            }
            return Value.Null;
        }

        /// <summary>
        /// excel.calculation() / (mode) → 获取/设置计算模式。
        /// mode: "auto" | "manual" | "semiauto"
        /// </summary>
        [PrototypeFunction(Name = "calculation")]
        public static Value calculation(ExcelModule self, Value? mode = null)
        {
            if (mode == null || mode is NullValue)
            {
                try
                {
                    return self._app.Calculation switch
                    {
                        XlCalculation.xlCalculationAutomatic => StringValue.Create("auto"),
                        XlCalculation.xlCalculationManual => StringValue.Create("manual"),
                        XlCalculation.xlCalculationSemiautomatic => StringValue.Create("semiauto"),
                        _ => Value.Null,
                    };
                }
                catch { return Value.Null; }
            }
            if (mode is StringValue sv)
            {
                try
                {
                    self._app.Calculation = sv.Value switch
                    {
                        "manual" => XlCalculation.xlCalculationManual,
                        "semiauto" => XlCalculation.xlCalculationSemiautomatic,
                        _ => XlCalculation.xlCalculationAutomatic,
                    };
                }
                catch { }
            }
            return Value.Null;
        }

        /// <summary>excel.calculate() → 强制重新计算所有打开的工作簿</summary>
        [PrototypeFunction(Name = "calculate")]
        public static Value calculate(ExcelModule self)
        {
            try { self._app.Calculate(); }
            catch { }
            return Value.Null;
        }

        // ========================================================================
        // 剪贴板工具
        // ========================================================================

        /// <summary>excel.copyText(text) → 复制文本到系统剪贴板</summary>
        [PrototypeFunction(Name = "copyText")]
        public static Value copyText(ExcelModule self, StringValue text)
        {
            try { System.Windows.Forms.Clipboard.SetText(text.Value); }
            catch { }
            return Value.Null;
        }

        /// <summary>excel.pasteText() → 从系统剪贴板读取文本</summary>
        [PrototypeFunction(Name = "pasteText")]
        public static Value pasteText(ExcelModule self)
        {
            try { return StringValue.Create(System.Windows.Forms.Clipboard.GetText()); }
            catch { return Value.Null; }
        }

    }

}
