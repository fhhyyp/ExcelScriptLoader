// ============================================================================
// ExcelModule.Cell — Cell API 工厂
// ObjectValue + FunctionValue 模式，将 Excel COM Range（单个单元格）包装为脚本对象
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang.Runtime;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{
    public partial class ExcelModule
    {

    // ============================================================================
    // CellToObject — 将单个 ExcelRange (Cell) 包装为脚本可用的 ObjectValue
    //
    // 暴露的脚本 API:
    //   .value() / .value(v)    → any/void   读/写单元格值
    //   .formula() / .formula(f)→ str/void   读/写公式
    //   .format() / .format(fmt)→ obj/void   读/写格式
    //   .address()              → str        单元格地址 (如 "$A$1")
    //   .row()                  → num        行号
    //   .column()               → num        列号
    //   .hasFormula()           → bool       是否包含公式
    //
    // format 支持字段: bold, color, bgColor, size, numberFormat
    // ============================================================================

    internal static ObjectValue CellToObject(ExcelRange cell)
    {
        var c = cell;
        return new ObjectValue(new()
        {
            // -- 值读写 (getter/setter 同名方法) --
            ["value"] = F("value", args =>
            {
                if (args.Count == 0) // get
                {
                    try { return WrapCell(c.Value2); }
                    catch { return Value.Null; }
                }
                else // set
                {
                    try { c.Value2 = UnwrapV(args[0]); } catch { }
                    return Value.Null;
                }
            }),

            // -- 公式读写 (getter/setter 同名方法) --
            ["formula"] = F("formula", args =>
            {
                if (args.Count == 0) // get
                {
                    try { return StringValue.Create(c.Formula as string ?? ""); }
                    catch { return Value.Null; }
                }
                else if (args[0] is StringValue f) // set
                {
                    try { c.Formula = f.Value; } catch { }
                }
                return Value.Null;
            }),

            // -- 格式读写 (getter/setter 同名方法) --
            ["format"] = F("format", args =>
            {
                if (args.Count == 0) // get — 返回当前格式摘要
                {
                    var info = new Dictionary<string, Value>
                    {
                        ["bold"] = BoolValue.Create((bool)(c.Font.Bold ?? false)),
                    };
                    try { info["size"] = NumberValueFactory.Create((double)(c.Font.Size ?? 0)); } catch { }
                    return new ObjectValue(info);
                }
                else if (args[0] is ObjectValue fmt) // set — 应用格式对象
                {
                    try { ApplyFormat(c, fmt); } catch { }
                }
                return Value.Null;
            }),

            // -- 查询方法 --
            ["address"] = F("address", _ =>
            {
                try { return StringValue.Create(c.Address as string ?? ""); }
                catch { return Value.Null; }
            }),
            ["row"] = F("row", _ =>
            {
                try { return NumberValueFactory.Create(c.Row); }
                catch { return Value.Null; }
            }),
            ["column"] = F("column", _ =>
            {
                try { return NumberValueFactory.Create(c.Column); }
                catch { return Value.Null; }
            }),
            ["hasFormula"] = F("hasFormula", _ =>
            {
                try { return BoolValue.Create((bool)c.HasFormula); }
                catch { return Value.Null; }
            }),
        });
    }

    }
}
