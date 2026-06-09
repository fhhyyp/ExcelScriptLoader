// ============================================================================
// ExcelModule.Worksheet — Worksheet API 工厂
// ObjectValue + FunctionValue 模式，将 Excel COM Worksheet 包装为脚本对象
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang.Runtime;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{
    public partial class ExcelModule
    {

        // ============================================================================
        // WorksheetToObject — 将 Worksheet COM 对象包装为脚本可用的 ObjectValue
        //
        // 暴露的脚本 API:
        //   .name                → String   工作表名（只读属性）
        //   .activate()          → void     激活/切换到当前工作表
        //   .cell(addr)          → Cell     按地址获取单元格
        //   .range(addr)         → Range    按地址获取区域
        //   .table(name)         → Table    按名称获取 Excel 表格
        //   .usedRange()         → Range    获取已用区域
        //   .read(addr?)         → Array    读取数据（无参=全部，有参=指定区域）
        //   .write(addr, data)   → void     写入二维数组
        //   .writeObjects(addr, data) → void  写入对象数组（自动生成表头）
        //   .insertRow(index)    → void     在指定行前插入一行
        //   .removeRow(index)    → void     删除指定行
        //   .insertColumn(index) → void     在指定列前插入一列
        //   .removeColumn(index) → void     删除指定列
        //   .hideRow(index)      → void     隐藏指定行
        //   .hideColumn(index)   → void     隐藏指定列
        //   .rowHeight()         → Num      读标准行高
        //   .rowHeight(row)      → Num      读指定行行高
        //   .rowHeight(row, h)   → void     设置指定行行高
        //   .columnWidth()       → Num      读标准列宽
        //   .columnWidth(col)    → Num      读指定列列宽
        //   .columnWidth(col, w) → void     设置指定列列宽
        //   .visible()           → bool     读可见性
        //   .visible(mode)       → void     写可见性 (true/false/"hidden"/"veryHidden")
        //   .showAllData()       → void     清除工作表中所有筛选
        // ============================================================================

        internal static ObjectValue WorksheetToObject(Worksheet ws)
        {
            var w = ws;
            return new ObjectValue(new()
            {
                // -- 属性 & 基本操作 --
                ["name"] = StringValue.Create(w.Name),
                ["activate"] = F("activate", _ => { try { w.Activate(); } catch { } }),

                // -- 行列操作 --
                ["insertRow"] = F("insertRow", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                    {
                        try { ((ExcelRange)w.Rows[(int)idx.Value]).Insert(XlInsertShiftDirection.xlShiftDown); }
                        catch { }
                    }
                }),
                ["removeRow"] = F("removeRow", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                    {
                        try { ((ExcelRange)w.Rows[(int)idx.Value]).Delete(XlDeleteShiftDirection.xlShiftUp); }
                        catch { }
                    }
                }),
                ["insertColumn"] = F("insertColumn", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                    {
                        try { ((ExcelRange)w.Columns[(int)idx.Value]).Insert(XlInsertShiftDirection.xlShiftToRight); }
                        catch { }
                    }
                }),
                ["removeColumn"] = F("removeColumn", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                    {
                        try { ((ExcelRange)w.Columns[(int)idx.Value]).Delete(XlDeleteShiftDirection.xlShiftToLeft); }
                        catch { }
                    }
                }),
                ["hideRow"] = F("hideRow", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                    {
                        try { ((ExcelRange)w.Rows[(int)idx.Value]).Hidden = true; }
                        catch { }
                    }
                }),
                ["hideColumn"] = F("hideColumn", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                    {
                        try { ((ExcelRange)w.Columns[(int)idx.Value]).Hidden = true; }
                        catch { }
                    }
                }),

                // -- 行高 / 列宽 (getter/setter 同名方法) --
                ["rowHeight"] = F("rowHeight", args =>
                {
                    if (args.Count == 0)
                    {
                        try { return NumberValueFactory.Create((double)w.StandardHeight); }
                        catch { return Value.Null; }
                    }
                    if (args[0] is NumberValue<double> row)
                    {
                        if (args.Count == 1) // get
                        {
                            try { return NumberValueFactory.Create((double)((ExcelRange)w.Rows[(int)row.Value]).RowHeight); }
                            catch { return Value.Null; }
                        }
                        if (args.Count >= 2 && args[1] is NumberValue<double> h) // set
                        {
                            try { ((ExcelRange)w.Rows[(int)row.Value]).RowHeight = h.Value; }
                            catch { }
                        }
                    }
                    return Value.Null;
                }),
                ["columnWidth"] = F("columnWidth", args =>
                {
                    if (args.Count == 0)
                    {
                        try { return NumberValueFactory.Create((double)w.StandardWidth); }
                        catch { return Value.Null; }
                    }
                    if (args[0] is NumberValue<double> col)
                    {
                        if (args.Count == 1) // get
                        {
                            try { return NumberValueFactory.Create((double)((ExcelRange)w.Columns[(int)col.Value]).ColumnWidth); }
                            catch { return Value.Null; }
                        }
                        if (args.Count >= 2 && args[1] is NumberValue<double> wd) // set
                        {
                            try { ((ExcelRange)w.Columns[(int)col.Value]).ColumnWidth = wd.Value; }
                            catch { }
                        }
                    }
                    return Value.Null;
                }),

                // -- 可见性 (getter/setter 同名方法) --
                ["visible"] = F("visible", args =>
                {
                    if (args.Count == 0) // get
                    {
                        try { return BoolValue.Create(w.Visible == XlSheetVisibility.xlSheetVisible); }
                        catch { return Value.Null; }
                    }
                    try // set
                    {
                        if (args[0] is BoolValue bv)
                            w.Visible = bv.Value ? XlSheetVisibility.xlSheetVisible
                                                 : XlSheetVisibility.xlSheetHidden;
                        else if (args[0] is StringValue sv)
                            w.Visible = sv.Value switch
                            {
                                "veryHidden" => XlSheetVisibility.xlSheetVeryHidden,
                                "hidden" => XlSheetVisibility.xlSheetHidden,
                                _ => XlSheetVisibility.xlSheetVisible,
                            };
                    }
                    catch { }
                    return Value.Null;
                }),

                // -- 子对象访问器 --
                ["cell"] = F("cell", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue a)
                        try { return CellToObject(w.Range[a.Value]); } catch { }
                    return Value.Null;
                }),
                ["range"] = F("range", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue a)
                        try { return RangeToObject(w.Range[a.Value]); } catch { }
                    return Value.Null;
                }),
                ["table"] = F("table", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue tn)
                        try { return TableToObject(w.ListObjects[tn.Value]); } catch { }
                    return Value.Null;
                }),
                ["usedRange"] = F("usedRange", _ =>
                {
                    try { return RangeToObject(w.UsedRange); }
                    catch { return Value.Null; }
                }),

                // -- 数据读写 --
                ["read"] = F("read", args =>
                {
                    try
                    {
                        ExcelRange rng = args.Count > 0 && args[0] is StringValue a
                            ? w.Range[a.Value] : w.UsedRange;
                        return ReadRange(rng);
                    }
                    catch { return Value.Null; }
                }),
                ["write"] = F("write", args =>
                {
                    if (args.Count >= 2 && args[0] is StringValue a)
                        try { WriteRange(w.Range[a.Value], args[1]); } catch { }
                }),
                ["writeObjects"] = F("writeObjects", args =>
                {
                    if (args.Count >= 2 && args[0] is StringValue a)
                        try { WriteObjects(w.Range[a.Value], args[1]); } catch { }
                }),

                // -- 筛选控制 --
                ["showAllData"] = F("showAllData", _ =>
                {
                    try { w.ShowAllData(); } catch { }
                }),
            });
        }

    }
}
