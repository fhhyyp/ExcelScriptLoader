// ============================================================================
// ExcelModule.Range — Range API 工厂
// ObjectValue + FunctionValue 模式，将 Excel COM Range 包装为脚本对象
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang.Runtime;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{
    public partial class ExcelModule
    {

        // ============================================================================
        // RangeToObject — 将 ExcelRange COM 对象包装为脚本可用的 ObjectValue
        //
        // 暴露的脚本 API:
        //   .values() / .values(arr)     → Array/void  读/写全部值
        //   .formulas() / .formulas(arr) → Array/void  读/写全部公式
        //   .rows()                      → Array       所有行（二维数组的数组）
        //   .rowCount()                  → Num         行数
        //   .colCount()                  → Num         列数
        //   .clear()                     → void        清空内容和格式
        //   .border() / .border(opts)    → obj/void    读/写边框
        //        opts: { style, edges, color? }
        //        style: "thin"|"dash"|"dot"|"double"|"none"
        //        edges: "all"|"top"|"bottom"|"left"|"right" (可组合)
        //   .wrapText() / .wrapText(b)   → bool/void   读/写自动换行
        //   .alignment() / .alignment(o) → obj/void    读/写对齐
        //        o: { horizontal?, vertical? }
        //        horizontal: "center"|"left"|"right"|"justify"|"general"
        //        vertical:   "middle"|"top"|"bottom"|"justify"
        //   .offset(rows, cols)          → Range       偏移区域
        //   .find(what, opts?)           → Range|null  查找单元格
        //        opts: { lookAt?, matchCase? }
        //   .findNext()                  → Range|null  继续查找
        //   .autoFilter(field?, crit?)   → void        自动筛选
        //   .sort(key, opts?)            → void        排序
        //        opts: { order?, header? }
        //   .cut() / .cut(target)        → void        剪切到剪贴板 / 剪切到目标
        //   .insert(direction?)          → void        插入单元格 ("down"|"right")
        //   .delete(direction?)          → void        删除单元格 ("up"|"left")
        //   .entireRow()                 → Range       整行范围
        //   .entireColumn()              → Range       整列范围
        //   .merge()                     → void        合并单元格
        //   .unmerge()                   → void        取消合并
        //   .autoFit()                   → void        自动调整行高列宽
        //   .copy(target)                → void        复制到目标地址
        // ============================================================================

        internal static ObjectValue RangeToObject(ExcelRange rng)
        {
            var r = rng;
            return new ObjectValue(new()
            {
                // -- 数据读写 --
                ["values"] = F("values", args =>
                {
                    if (args.Count == 0) // get
                    {
                        try { return ReadRange(r); }
                        catch { return Value.Null; }
                    }
                    else // set
                    {
                        try { WriteRange(r, args[0]); } catch { }
                        return Value.Null;
                    }
                }),
                ["formulas"] = F("formulas", args =>
                {
                    if (args.Count == 0) // get — 读取公式二维数组
                    {
                        try
                        {
                            var data = (object[,])r.Formula;
                            int rows = data.GetLength(0), cols = data.GetLength(1);
                            var result = new List<Value>();
                            for (int rr = 1; rr <= rows; rr++)
                            {
                                var row = new List<Value>();
                                for (int cc = 1; cc <= cols; cc++)
                                    row.Add(StringValue.Create(data[rr, cc]?.ToString() ?? ""));
                                result.Add(new ArrayValue(row));
                            }
                            return new ArrayValue(result);
                        }
                        catch { return Value.Null; }
                    }
                    else // set — 写入公式二维数组
                    {
                        if (args[0] is ArrayValue av && av.Elements.Count > 0
                            && av.Elements[0] is ArrayValue firstRow)
                        {
                            try
                            {
                                int rows = av.Elements.Count;
                                int cols = firstRow.Elements.Count;
                                var arr = new object[rows, cols];
                                for (int rr = 0; rr < rows; rr++)
                                {
                                    if (av.Elements[rr] is not ArrayValue row) continue;
                                    for (int cc = 0; cc < cols && cc < row.Elements.Count; cc++)
                                        arr[rr, cc] = row.Elements[cc] is StringValue sv ? sv.Value : "";
                                }
                                r.Resize[rows, cols].Formula = arr;
                            }
                            catch { }
                        }
                    }
                    return Value.Null;
                }),

                // -- 查询 --
                ["rows"] = F("rows", _ =>
                {
                    try
                    {
                        var data = (object[,])r.Value2;
                        var result = new List<Value>();
                        int rows = data.GetLength(0), cols = data.GetLength(1);
                        for (int rr = 1; rr <= rows; rr++)
                        {
                            var row = new List<Value>();
                            for (int cc = 1; cc <= cols; cc++)
                                row.Add(WrapCell(data[rr, cc]));
                            result.Add(new ArrayValue(row));
                        }
                        return new ArrayValue(result);
                    }
                    catch { return Value.Null; }
                }),
                ["rowCount"] = F("rowCount", _ => NumberValueFactory.Create(r.Rows.Count)),
                ["colCount"] = F("colCount", _ => NumberValueFactory.Create(r.Columns.Count)),
                ["clear"] = F("clear", _ => { try { r.Clear(); } catch { } }),

                // -- 格式化 --
                ["border"] = F("border", args =>
                {
                    if (args.Count == 0) // get
                    {
                        try
                        {
                            var b = r.Borders[XlBordersIndex.xlEdgeBottom];
                            return new ObjectValue(new()
                            {
                                ["style"] = StringValue.Create(
                                    (XlLineStyle)b.LineStyle == XlLineStyle.xlContinuous ? "thin" : "none"),
                            });
                        }
                        catch { return Value.Null; }
                    }
                    else if (args[0] is ObjectValue opts) // set
                    {
                        try
                        {
                            XlLineStyle lineStyle = XlLineStyle.xlContinuous;
                            if (opts.Properties.TryGetValue("style", out var sty) && sty is StringValue stys)
                                lineStyle = stys.Value switch
                                {
                                    "thin" => XlLineStyle.xlContinuous,
                                    "dash" => XlLineStyle.xlDash,
                                    "dot" => XlLineStyle.xlDot,
                                    "double" => XlLineStyle.xlDouble,
                                    "none" => XlLineStyle.xlLineStyleNone,
                                    _ => XlLineStyle.xlContinuous,
                                };
                            int colorVal = 0;
                            if (opts.Properties.TryGetValue("color", out var bc))
                                colorVal = (int)ParseColor(bc);
                            string edges = "all";
                            if (opts.Properties.TryGetValue("edges", out var ed) && ed is StringValue eds)
                                edges = eds.Value;
                            // 按 edges 字符串匹配四个边
                            foreach (XlBordersIndex idx in new[] {
                            XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom,
                            XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight })
                            {
                                bool match = edges == "all"
                                    || (idx == XlBordersIndex.xlEdgeTop && edges.Contains("top"))
                                    || (idx == XlBordersIndex.xlEdgeBottom && edges.Contains("bottom"))
                                    || (idx == XlBordersIndex.xlEdgeLeft && edges.Contains("left"))
                                    || (idx == XlBordersIndex.xlEdgeRight && edges.Contains("right"));
                                if (match)
                                {
                                    r.Borders[idx].LineStyle = lineStyle;
                                    if (colorVal != 0) r.Borders[idx].Color = colorVal;
                                }
                            }
                        }
                        catch { }
                    }
                    return Value.Null;
                }),
                ["wrapText"] = F("wrapText", args =>
                {
                    if (args.Count == 0) // get
                    {
                        try { return BoolValue.Create((bool)r.WrapText); }
                        catch { return Value.Null; }
                    }
                    else if (args[0] is BoolValue bv) // set
                    {
                        try { r.WrapText = bv.Value; } catch { }
                    }
                    return Value.Null;
                }),
                ["alignment"] = F("alignment", args =>
                {
                    if (args.Count == 0) // get
                    {
                        try
                        {
                            return new ObjectValue(new()
                            {
                                ["horizontal"] = StringValue.Create(
                                    r.HorizontalAlignment?.ToString()?.ToLowerInvariant() ?? "general"),
                                ["vertical"] = StringValue.Create(
                                    r.VerticalAlignment?.ToString()?.ToLowerInvariant() ?? "bottom"),
                            });
                        }
                        catch { return Value.Null; }
                    }
                    else if (args[0] is ObjectValue opts) // set
                    {
                        try
                        {
                            if (opts.Properties.TryGetValue("horizontal", out var hv) && hv is StringValue hvs)
                                r.HorizontalAlignment = hvs.Value switch
                                {
                                    "center" => XlHAlign.xlHAlignCenter,
                                    "left" => XlHAlign.xlHAlignLeft,
                                    "right" => XlHAlign.xlHAlignRight,
                                    "justify" => XlHAlign.xlHAlignJustify,
                                    _ => XlHAlign.xlHAlignGeneral,
                                };
                            if (opts.Properties.TryGetValue("vertical", out var vv) && vv is StringValue vvs)
                                r.VerticalAlignment = vvs.Value switch
                                {
                                    "middle" => XlVAlign.xlVAlignCenter,
                                    "top" => XlVAlign.xlVAlignTop,
                                    "bottom" => XlVAlign.xlVAlignBottom,
                                    "justify" => XlVAlign.xlVAlignJustify,
                                    _ => XlVAlign.xlVAlignBottom,
                                };
                        }
                        catch { }
                    }
                    return Value.Null;
                }),

                // -- 导航 --
                ["offset"] = F("offset", args =>
                {
                    if (args.Count >= 2
                        && args[0] is NumberValue<double> ro
                        && args[1] is NumberValue<double> co)
                    {
                        try { return RangeToObject(r.Offset[(int)ro.Value, (int)co.Value]); }
                        catch { }
                    }
                    return Value.Null;
                }),

                // -- 查找 --
                ["find"] = F("find", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue what)
                    {
                        try
                        {
                            XlLookAt lookAt = XlLookAt.xlPart;
                            bool matchCase = false;
                            if (args.Count > 1 && args[1] is ObjectValue opts)
                            {
                                if (opts.Properties.TryGetValue("lookAt", out var la) && la is StringValue las)
                                    lookAt = las.Value == "whole" ? XlLookAt.xlWhole : XlLookAt.xlPart;
                                if (opts.Properties.TryGetValue("matchCase", out var mc) && mc is BoolValue mcb)
                                    matchCase = mcb.Value;
                            }
                            var found = r.Find(what.Value, Type.Missing,
                                XlFindLookIn.xlValues, lookAt,
                                XlSearchOrder.xlByRows, XlSearchDirection.xlNext, matchCase);
                            return found != null ? RangeToObject((ExcelRange)found) : Value.Null;
                        }
                        catch { }
                    }
                    return Value.Null;
                }),
                ["findNext"] = F("findNext", _ =>
                {
                    try
                    {
                        var found = r.FindNext(Type.Missing);
                        return found != null ? RangeToObject((ExcelRange)found) : Value.Null;
                    }
                    catch { return Value.Null; }
                }),

                // -- 筛选排序 --
                ["autoFilter"] = F("autoFilter", args =>
                {
                    try
                    {
                        if (args.Count == 0)
                            r.AutoFilter(); // 切换筛选开关
                        else if (args.Count >= 2
                            && args[0] is NumberValue<double> field
                            && args[1] is StringValue crit)
                            r.AutoFilter((int)field.Value, crit.Value); // 按条件筛选列
                    }
                    catch { }
                }),
                ["sort"] = F("sort", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue key)
                    {
                        try
                        {
                            XlSortOrder order = XlSortOrder.xlAscending;
                            bool hasHeader = true;
                            if (args.Count > 1 && args[1] is ObjectValue opts)
                            {
                                if (opts.Properties.TryGetValue("order", out var o) && o is StringValue os)
                                    order = os.Value == "desc" ? XlSortOrder.xlDescending : XlSortOrder.xlAscending;
                                if (opts.Properties.TryGetValue("header", out var h) && h is BoolValue hb)
                                    hasHeader = hb.Value;
                            }
                            // Range.Sort(Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, ...)
                            r.Sort(r.Columns[key.Value], order,
                                Type.Missing, Type.Missing, XlSortOrder.xlAscending,
                                Type.Missing, XlSortOrder.xlAscending,
                                hasHeader ? XlYesNoGuess.xlYes : XlYesNoGuess.xlNo);
                        }
                        catch { }
                    }
                }),

                // -- 剪贴板与行列操作 --
                ["cut"] = F("cut", args =>
                {
                    if (args.Count == 0) // 剪切到剪贴板
                    {
                        try { r.Cut(); } catch { }
                    }
                    else if (args[0] is StringValue t) // 剪切到目标地址
                    {
                        try { r.Cut(r.Worksheet.Range[t.Value]); } catch { }
                    }
                }),
                ["insert"] = F("insert", args =>
                {
                    try
                    {
                        var dir = XlInsertShiftDirection.xlShiftDown;
                        if (args.Count > 0 && args[0] is StringValue sd)
                            dir = sd.Value == "right" ? XlInsertShiftDirection.xlShiftToRight
                                                       : XlInsertShiftDirection.xlShiftDown;
                        r.Insert(dir);
                    }
                    catch { }
                }),
                ["delete"] = F("delete", args =>
                {
                    try
                    {
                        var dir = XlDeleteShiftDirection.xlShiftUp;
                        if (args.Count > 0 && args[0] is StringValue sd)
                            dir = sd.Value == "left" ? XlDeleteShiftDirection.xlShiftToLeft
                                                       : XlDeleteShiftDirection.xlShiftUp;
                        r.Delete(dir);
                    }
                    catch { }
                }),
                ["entireRow"] = F("entireRow", _ =>
                {
                    try { return RangeToObject(r.EntireRow); }
                    catch { return Value.Null; }
                }),
                ["entireColumn"] = F("entireColumn", _ =>
                {
                    try { return RangeToObject(r.EntireColumn); }
                    catch { return Value.Null; }
                }),

                // -- 合并与自适应 --
                ["merge"] = F("merge", _ => { try { r.Merge(); } catch { } }),
                ["unmerge"] = F("unmerge", _ => { try { r.UnMerge(); } catch { } }),
                ["autoFit"] = F("autoFit", _ => { try { r.AutoFit(); } catch { } }),

                // -- 复制 --
                ["copy"] = F("copy", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue t)
                        try { r.Copy(r.Worksheet.Range[t.Value]); } catch { }
                }),
            });
        }

    }
}
