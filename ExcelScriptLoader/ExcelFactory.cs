// ============================================================================
// ExcelFactory — Excel COM 对象包装工厂 + 内部工具方法
// 将 Workbook/Worksheet/Cell/Range/Table COM 对象包装为脚本 ObjectValue。
// 通过 using static ExcelScriptLoader.ExcelFactory 导入到 ExcelModule。
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang;
using ScriptLang.Runtime;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{
    internal static class ExcelFactory
    {

        // ========================================================================
        // 工具: F() — FunctionValue 工厂
        // ========================================================================

        public static FunctionValue F(string name, Func<List<Value>, Value> func)
            => new(name, func);

        public static FunctionValue F(string name, Action<List<Value>> action)
            => new(name, args => { action(args); return Value.Null; });

        // ========================================================================
        // 工具: WrapCell / UnwrapV — ScriptLang ↔ COM 值转换
        // ========================================================================

        /// <summary>从单元格值转换为 Value </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        public static Value WrapCell(object? v) => v switch
        {
            null => Value.Null,
            string s => StringValue.Create(s),
            int i => NumberValueFactory.Create(i),
            long l => NumberValueFactory.Create(l),
            double d => NumberValueFactory.Create(d),
            float f => NumberValueFactory.Create(f),
            bool b => BoolValue.Create(b),
            _ => StringValue.Create(v.ToString() ?? ""),
        };

        /// <summary>转换为写入到单元格的值</summary>
        /// <param name="v"></param>
        /// <returns></returns>
        public static object UnwrapV(Value v) => v switch
        {
            StringValue sv => sv.Value,
            NumberValue<int> ni => ni.Value,
            NumberValue<long> nl => nl.Value,
            NumberValue<double> nd => nd.Value,
            NumberValue<float> nf => nf.Value,
            BoolValue bv => bv.Value,
            NullValue => "",
            _ => v.ToString() ?? "",
        };


        /// <summary>读取区域值</summary>
        /// <param name="rng"></param>
        /// <returns></returns>
        public static Value ReadRange(ExcelRange rng)
        {
            var data = (object[,])rng.Value2;
            int rows = data.GetLength(0), cols = data.GetLength(1);
            var result = new List<Value>();
            for (int r = 1; r <= rows; r++)
            {
                var row = new List<Value>();
                for (int c = 1; c <= cols; c++)
                    row.Add(WrapCell(data[r, c]));
                result.Add(new ArrayValue(row));
            }
            return new ArrayValue(result);
        }


        /// <summary>写入到区域单元格</summary>
        /// <param name="rng"></param>
        /// <param name="data"></param>
        public static void WriteRange(ExcelRange rng, Value data)
        {
            if (data is not ArrayValue av || av.Elements.Count == 0) return;
            if (av.Elements[0] is not ArrayValue firstRow) return;
            int rows = av.Elements.Count;
            int cols = firstRow.Elements.Count;
            var arr = new object[rows, cols];
            for (int r = 0; r < rows; r++)
            {
                if (av.Elements[r] is not ArrayValue row) continue;
                for (int c = 0; c < cols && c < row.Elements.Count; c++)
                    arr[r, c] = UnwrapV(row.Elements[c]);
            }
            rng.Resize[rows, cols].Value2 = arr;
        }

        /// <summary>写入到对象</summary>
        /// <param name="rng"></param>
        /// <param name="data"></param>
        public static void WriteObjects(ExcelRange rng, Value data)
        {
            if (data is not ArrayValue av || av.Elements.Count == 0) return;
            if (av.Elements[0] is not ObjectValue first) return;
            var keys = first.Properties.Keys.ToArray();
            int rows = av.Elements.Count + 1;
            int cols = keys.Length;
            var arr = new object[rows, cols];
            for (int c = 0; c < cols; c++) arr[0, c] = keys[c];
            for (int r = 0; r < av.Elements.Count; r++)
            {
                if (av.Elements[r] is not ObjectValue obj) continue;
                for (int c = 0; c < cols; c++)
                    arr[r + 1, c] = obj.Properties.TryGetValue(keys[c], out var v)
                        ? UnwrapV(v) : "";
            }
            rng.Resize[rows, cols].Value2 = arr;
        }

        /// <summary> 对区域追加样式设置 </summary>
        /// <param name="cell"></param>
        /// <param name="fmt"></param>
        public static void ApplyFormat(ExcelRange cell, ObjectValue fmt)
        {
            if (fmt.Properties.TryGetValue("bold", out var bv) && bv is BoolValue b)
                cell.Font.Bold = b.Value;
            if (fmt.Properties.TryGetValue("color", out var cv))
                cell.Font.Color = ParseColor(cv);
            if (fmt.Properties.TryGetValue("bgColor", out var bg))
                cell.Interior.Color = ParseColor(bg);
            if (fmt.Properties.TryGetValue("size", out var sz))
            {
                if (sz is NumberValue<double> nd) cell.Font.Size = nd.Value;
                else if (sz is NumberValue<int> ni) cell.Font.Size = ni.Value;
            }
            if (fmt.Properties.TryGetValue("numberFormat", out var nf) && nf is StringValue nfs)
                cell.NumberFormat = nfs.Value;
        }

        /// <summary> 转换为颜色值 </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static object ParseColor(Value c) => c switch
        {
            StringValue sv => sv.Value.ToLowerInvariant() switch
            {
                "red" => 0xFF0000,
                "green" => 0x00FF00,
                "blue" => 0x0070C0,
                "yellow" => 0xFFFF00,
                "white" => 0xFFFFFF,
                "black" => 0x000000,
                "gray" or "grey" => 0x808080,
                "orange" => 0xFF8C00,
                _ => 0,
            },
            NumberValue<int> ni => ni.Value,
            NumberValue<double> nd => (int)nd.Value,
            _ => 0,
        };

        // ============================================================================
        // WorkbookToObject — 将 Workbook COM 对象包装为脚本可用的 ObjectValue
        //
        // 暴露的脚本 API:
        //   .name              → String   文件名（只读属性）
        //   .path()            → String   文件所在目录路径
        //   .fullName()        → String   完整路径（含文件名）
        //   .save()            → void     保存
        //   .saveAs(path)      → void     另存为
        //   .close()           → void     关闭（不保存）
        //   .sheet(name)       → Worksheet  按名称获取工作表
        //   .addSheet(name)    → Worksheet  新建工作表
        //   .removeSheet(name) → void       删除工作表
        //   .sheets()          → Array      所有工作表数组
        // ============================================================================

        public static ObjectValue WorkbookToObject(Workbook wb)
        {
            var w = wb;
            return new ObjectValue(new()
            {
                ["name"] = StringValue.Create(w.Name),
                ["path"] = StringValue.Create(w.Path ?? ""),
                ["fullName"] = StringValue.Create(w.FullName ?? ""),
                ["save"] = F("save", _ => { try { w.Save(); } catch { } }),
                ["saveAs"] = F("saveAs", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue p)
                        try { w.SaveAs(p.Value); } catch { }
                }),
                ["close"] = F("close", _ => { try { w.Close(false); } catch { } }),
                ["sheet"] = F("sheet", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue sn)
                        try { return WorksheetToObject((Worksheet)w.Sheets[sn.Value]); } catch { }
                    return Value.Null;
                }),
                ["addSheet"] = F("addSheet", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue sn)
                    {
                        try
                        {
                            var ws = (Worksheet)w.Sheets.Add();
                            ws.Name = sn.Value;
                            return WorksheetToObject(ws);
                        }
                        catch { }
                    }
                    return Value.Null;
                }),
                ["removeSheet"] = F("removeSheet", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue sn)
                        try { ((Worksheet)w.Sheets[sn.Value]).Delete(); } catch { }
                }),
                ["sheets"] = F("sheets", _ =>
                {
                    var list = new List<Value>();
                    try { foreach (Worksheet ws in w.Sheets) list.Add(WorksheetToObject(ws)); } catch { }
                    return new ArrayValue(list);
                }),
            });
        }

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
        public static ObjectValue WorksheetToObject(Worksheet ws)
        {
            var w = ws;
            return new ObjectValue(new()
            {
                ["name"] = StringValue.Create(w.Name),
                ["activate"] = F("activate", _ => { try { w.Activate(); } catch { } }),

                // 行列操作
                ["insertRow"] = F("insertRow", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                        try { ((ExcelRange)w.Rows[(int)idx.Value]).Insert(XlInsertShiftDirection.xlShiftDown); } catch { }
                }),
                ["removeRow"] = F("removeRow", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                        try { ((ExcelRange)w.Rows[(int)idx.Value]).Delete(XlDeleteShiftDirection.xlShiftUp); } catch { }
                }),
                ["insertColumn"] = F("insertColumn", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                        try { ((ExcelRange)w.Columns[(int)idx.Value]).Insert(XlInsertShiftDirection.xlShiftToRight); } catch { }
                }),
                ["removeColumn"] = F("removeColumn", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                        try { ((ExcelRange)w.Columns[(int)idx.Value]).Delete(XlDeleteShiftDirection.xlShiftToLeft); } catch { }
                }),
                ["hideRow"] = F("hideRow", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                        try { ((ExcelRange)w.Rows[(int)idx.Value]).Hidden = true; } catch { }
                }),
                ["hideColumn"] = F("hideColumn", args =>
                {
                    if (args.Count > 0 && args[0] is NumberValue<double> idx)
                        try { ((ExcelRange)w.Columns[(int)idx.Value]).Hidden = true; } catch { }
                }),

                // 行高/列宽
                ["rowHeight"] = F("rowHeight", args =>
                {
                    if (args.Count == 0)
                    {
                        try { return NumberValueFactory.Create((double)w.StandardHeight); }
                        catch { return Value.Null; }
                    }
                    if (args[0] is NumberValue<double> row)
                    {
                        if (args.Count == 1)
                        {
                            try { return NumberValueFactory.Create((double)((ExcelRange)w.Rows[(int)row.Value]).RowHeight); }
                            catch { return Value.Null; }
                        }
                        if (args.Count >= 2 && args[1] is NumberValue<double> h)
                            try { ((ExcelRange)w.Rows[(int)row.Value]).RowHeight = h.Value; } catch { }
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
                        if (args.Count == 1)
                        {
                            try { return NumberValueFactory.Create((double)((ExcelRange)w.Columns[(int)col.Value]).ColumnWidth); }
                            catch { return Value.Null; }
                        }
                        if (args.Count >= 2 && args[1] is NumberValue<double> wd)
                            try { ((ExcelRange)w.Columns[(int)col.Value]).ColumnWidth = wd.Value; } catch { }
                    }
                    return Value.Null;
                }),

                // 可见性
                ["visible"] = F("visible", args =>
                {
                    if (args.Count == 0)
                    {
                        try { return BoolValue.Create(w.Visible == XlSheetVisibility.xlSheetVisible); }
                        catch { return Value.Null; }
                    }
                    try
                    {
                        if (args[0] is BoolValue bv)
                            w.Visible = bv.Value ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
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

                // 子对象访问器
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
                    try { return RangeToObject(w.UsedRange); } catch { return Value.Null; }
                }),

                // 数据读写
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

                // 筛选控制
                ["showAllData"] = F("showAllData", _ => { try { w.ShowAllData(); } catch { } }),
            });
        }

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

        public static ObjectValue CellToObject(ExcelRange cell)
        {
            var c = cell;
            return new ObjectValue(new()
            {
                ["value"] = F("value", args =>
                {
                    if (args.Count == 0) { try { return WrapCell(c.Value2); } catch { return Value.Null; } }
                    else { try { c.Value2 = UnwrapV(args[0]); } catch { } return Value.Null; }
                }),
                ["formula"] = F("formula", args =>
                {
                    if (args.Count == 0)
                    {
                        try { return StringValue.Create(c.Formula as string ?? ""); }
                        catch { return Value.Null; }
                    }
                    else if (args[0] is StringValue f) { try { c.Formula = f.Value; } catch { } }
                    return Value.Null;
                }),
                ["format"] = F("format", args =>
                {
                    if (args.Count == 0)
                    {
                        var info = new Dictionary<string, Value>
                        {
                            ["bold"] = BoolValue.Create((bool)(c.Font.Bold ?? false)),
                        };
                        try { info["size"] = NumberValueFactory.Create((double)(c.Font.Size ?? 0)); } catch { }
                        return new ObjectValue(info);
                    }
                    else if (args[0] is ObjectValue fmt) { try { ApplyFormat(c, fmt); } catch { } }
                    return Value.Null;
                }),
                ["address"] = F("address", _ => { try { return StringValue.Create(c.Address as string ?? ""); } catch { return Value.Null; } }),
                ["row"] = F("row", _ => { try { return NumberValueFactory.Create(c.Row); } catch { return Value.Null; } }),
                ["column"] = F("column", _ => { try { return NumberValueFactory.Create(c.Column); } catch { return Value.Null; } }),
                ["hasFormula"] = F("hasFormula", _ => { try { return BoolValue.Create((bool)c.HasFormula); } catch { return Value.Null; } }),
            });
        }

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
        public static ObjectValue RangeToObject(ExcelRange rng)
        {
            var r = rng;
            return new ObjectValue(new()
            {
                ["values"] = F("values", args =>
                {
                    if (args.Count == 0) { try { return ReadRange(r); } catch { return Value.Null; } }
                    else { try { WriteRange(r, args[0]); } catch { } return Value.Null; }
                }),
                ["formulas"] = F("formulas", args =>
                {
                    if (args.Count == 0)
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
                    else if (args[0] is ArrayValue av && av.Elements.Count > 0 && av.Elements[0] is ArrayValue firstRow)
                    {
                        try
                        {
                            int rows = av.Elements.Count, cols = firstRow.Elements.Count;
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
                    return Value.Null;
                }),
                ["rows"] = F("rows", _ => { try { var d = (object[,])r.Value2; var res = new List<Value>(); int rws = d.GetLength(0), cls = d.GetLength(1); for (int rr = 1; rr <= rws; rr++) { var row = new List<Value>(); for (int cc = 1; cc <= cls; cc++) row.Add(WrapCell(d[rr, cc])); res.Add(new ArrayValue(row)); } return new ArrayValue(res); } catch { return Value.Null; } }),
                ["rowCount"] = F("rowCount", _ => NumberValueFactory.Create(r.Rows.Count)),
                ["colCount"] = F("colCount", _ => NumberValueFactory.Create(r.Columns.Count)),
                ["clear"] = F("clear", _ => { try { r.Clear(); } catch { } }),

                // 格式化
                ["border"] = F("border", args =>
                {
                    if (args.Count == 0)
                    {
                        try { var b = r.Borders[XlBordersIndex.xlEdgeBottom]; return new ObjectValue(new() { ["style"] = StringValue.Create((XlLineStyle)b.LineStyle == XlLineStyle.xlContinuous ? "thin" : "none") }); }
                        catch { return Value.Null; }
                    }
                    else if (args[0] is ObjectValue opts)
                    {
                        try
                        {
                            XlLineStyle ls = XlLineStyle.xlContinuous;
                            if (opts.Properties.TryGetValue("style", out var st) && st is StringValue sts) ls = sts.Value switch { "thin" => XlLineStyle.xlContinuous, "dash" => XlLineStyle.xlDash, "dot" => XlLineStyle.xlDot, "double" => XlLineStyle.xlDouble, "none" => XlLineStyle.xlLineStyleNone, _ => XlLineStyle.xlContinuous };
                            int cv = 0; if (opts.Properties.TryGetValue("color", out var bc)) cv = (int)ParseColor(bc);
                            string ed = "all"; if (opts.Properties.TryGetValue("edges", out var e) && e is StringValue eds) ed = eds.Value;
                            foreach (XlBordersIndex idx in new[] { XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight })
                                if (ed == "all" || (idx == XlBordersIndex.xlEdgeTop && ed.Contains("top")) || (idx == XlBordersIndex.xlEdgeBottom && ed.Contains("bottom")) || (idx == XlBordersIndex.xlEdgeLeft && ed.Contains("left")) || (idx == XlBordersIndex.xlEdgeRight && ed.Contains("right")))
                                { r.Borders[idx].LineStyle = ls; if (cv != 0) r.Borders[idx].Color = cv; }
                        }
                        catch { }
                    }
                    return Value.Null;
                }),
                ["wrapText"] = F("wrapText", args =>
                {
                    if (args.Count == 0) { try { return BoolValue.Create((bool)r.WrapText); } catch { return Value.Null; } }
                    else if (args[0] is BoolValue bv) { try { r.WrapText = bv.Value; } catch { } }
                    return Value.Null;
                }),
                ["alignment"] = F("alignment", args =>
                {
                    if (args.Count == 0)
                    {
                        try { return new ObjectValue(new() { ["horizontal"] = StringValue.Create(r.HorizontalAlignment?.ToString()?.ToLowerInvariant() ?? "general"), ["vertical"] = StringValue.Create(r.VerticalAlignment?.ToString()?.ToLowerInvariant() ?? "bottom") }); }
                        catch { return Value.Null; }
                    }
                    else if (args[0] is ObjectValue opts)
                    {
                        try
                        {
                            if (opts.Properties.TryGetValue("horizontal", out var hv) && hv is StringValue hvs) r.HorizontalAlignment = hvs.Value switch { "center" => XlHAlign.xlHAlignCenter, "left" => XlHAlign.xlHAlignLeft, "right" => XlHAlign.xlHAlignRight, "justify" => XlHAlign.xlHAlignJustify, _ => XlHAlign.xlHAlignGeneral };
                            if (opts.Properties.TryGetValue("vertical", out var vv) && vv is StringValue vvs) r.VerticalAlignment = vvs.Value switch { "middle" => XlVAlign.xlVAlignCenter, "top" => XlVAlign.xlVAlignTop, "bottom" => XlVAlign.xlVAlignBottom, "justify" => XlVAlign.xlVAlignJustify, _ => XlVAlign.xlVAlignBottom };
                        }
                        catch { }
                    }
                    return Value.Null;
                }),

                // 导航/查找
                ["offset"] = F("offset", args =>
                {
                    if (args.Count >= 2 && args[0] is NumberValue<double> ro && args[1] is NumberValue<double> co)
                    { try { return RangeToObject(r.Offset[(int)ro.Value, (int)co.Value]); } catch { } }
                    return Value.Null;
                }),
                ["find"] = F("find", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue what)
                    {
                        try
                        {
                            XlLookAt la = XlLookAt.xlPart; bool mc = false;
                            if (args.Count > 1 && args[1] is ObjectValue opts)
                            {
                                if (opts.Properties.TryGetValue("lookAt", out var l) && l is StringValue las) la = las.Value == "whole" ? XlLookAt.xlWhole : XlLookAt.xlPart;
                                if (opts.Properties.TryGetValue("matchCase", out var m) && m is BoolValue mb) mc = mb.Value;
                            }
                            var found = r.Find(what.Value, Type.Missing, XlFindLookIn.xlValues, la, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, mc);
                            return found != null ? RangeToObject((ExcelRange)found) : Value.Null;
                        }
                        catch { }
                    }
                    return Value.Null;
                }),
                ["findNext"] = F("findNext", _ =>
                {
                    try { var f = r.FindNext(Type.Missing); return f != null ? RangeToObject((ExcelRange)f) : Value.Null; }
                    catch { return Value.Null; }
                }),

                // 筛选排序
                ["autoFilter"] = F("autoFilter", args =>
                {
                    try
                    {
                        if (args.Count == 0) r.AutoFilter();
                        else if (args.Count >= 2 && args[0] is NumberValue<double> f && args[1] is StringValue c) r.AutoFilter((int)f.Value, c.Value);
                    }
                    catch { }
                }),
                ["sort"] = F("sort", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue key)
                    {
                        try
                        {
                            XlSortOrder o = XlSortOrder.xlAscending; bool h = true;
                            if (args.Count > 1 && args[1] is ObjectValue opts)
                            {
                                if (opts.Properties.TryGetValue("order", out var ov) && ov is StringValue os) o = os.Value == "desc" ? XlSortOrder.xlDescending : XlSortOrder.xlAscending;
                                if (opts.Properties.TryGetValue("header", out var hv) && hv is BoolValue hb) h = hb.Value;
                            }
                            r.Sort(r.Columns[key.Value], o, Type.Missing, Type.Missing, XlSortOrder.xlAscending, Type.Missing, XlSortOrder.xlAscending, h ? XlYesNoGuess.xlYes : XlYesNoGuess.xlNo);
                        }
                        catch { }
                    }
                }),

                // 剪贴板/行列
                ["cut"] = F("cut", args =>
                {
                    if (args.Count == 0) { try { r.Cut(); } catch { } }
                    else if (args[0] is StringValue t) { try { r.Cut(r.Worksheet.Range[t.Value]); } catch { } }
                }),
                ["insert"] = F("insert", args =>
                {
                    try { var d = XlInsertShiftDirection.xlShiftDown; if (args.Count > 0 && args[0] is StringValue sd) d = sd.Value == "right" ? XlInsertShiftDirection.xlShiftToRight : XlInsertShiftDirection.xlShiftDown; r.Insert(d); } catch { }
                }),
                ["delete"] = F("delete", args =>
                {
                    try { var d = XlDeleteShiftDirection.xlShiftUp; if (args.Count > 0 && args[0] is StringValue sd) d = sd.Value == "left" ? XlDeleteShiftDirection.xlShiftToLeft : XlDeleteShiftDirection.xlShiftUp; r.Delete(d); } catch { }
                }),
                ["entireRow"] = F("entireRow", _ => { try { return RangeToObject(r.EntireRow); } catch { return Value.Null; } }),
                ["entireColumn"] = F("entireColumn", _ => { try { return RangeToObject(r.EntireColumn); } catch { return Value.Null; } }),

                // 合并/自适应/复制
                ["merge"] = F("merge", _ => { try { r.Merge(); } catch { } }),
                ["unmerge"] = F("unmerge", _ => { try { r.UnMerge(); } catch { } }),
                ["autoFit"] = F("autoFit", _ => { try { r.AutoFit(); } catch { } }),
                ["copy"] = F("copy", args => { if (args.Count > 0 && args[0] is StringValue t) try { r.Copy(r.Worksheet.Range[t.Value]); } catch { } }),
            });
        }

        // ============================================================================
        // TableToObject — 将 ListObject (Excel 表格) 包装为脚本可用的 ObjectValue
        //
        // 暴露的脚本 API:
        //   .name          → String   表名（只读属性）
        //   .rows()        → Array    所有数据行（对象数组，key=列名）
        //   .add(row)      → void     新增一行（对象格式）
        //   .addAll(rows)  → void     批量新增（对象数组）
        //   .clear()       → void     清空数据（保留表头和结构）
        // ============================================================================
        public static ObjectValue TableToObject(ListObject table)
        {
            var t = table;
            return new ObjectValue(new()
            {
                ["name"] = StringValue.Create(t.Name),
                ["rows"] = F("rows", _ =>
                {
                    var result = new List<Value>();
                    try
                    {
                        var rng = (ExcelRange?)(t.DataBodyRange ?? t.Range);
                        if (rng == null) return new ArrayValue(result);
                        var data = (object[,])rng.Value2;
                        int rows = data.GetLength(0), cols = data.GetLength(1);
                        var hdrRng = t.HeaderRowRange as ExcelRange;
                        var hdrData = (object[,]?)hdrRng?.Value2;
                        var headers = new string[cols];
                        for (int c = 1; c <= cols; c++) headers[c - 1] = hdrData?[1, c]?.ToString() ?? $"Col{c}";
                        for (int rr = 1; rr <= rows; rr++)
                        {
                            var row = new Dictionary<string, Value>();
                            for (int cc = 1; cc <= cols; cc++) row[headers[cc - 1]] = WrapCell(data[rr, cc]);
                            result.Add(new ObjectValue(row));
                        }
                    }
                    catch { }
                    return new ArrayValue(result);
                }),
                ["add"] = F("add", args =>
                {
                    if (args.Count > 0 && args[0] is ObjectValue obj)
                    {
                        try { var lr = t.ListRows.Add(); int col = 1; foreach (var kv in obj.Properties) ((ExcelRange)lr.Range[1, col++]).Value2 = UnwrapV(kv.Value); }
                        catch { }
                    }
                }),
                ["addAll"] = F("addAll", args =>
                {
                    if (args.Count > 0 && args[0] is ArrayValue arr)
                        foreach (var row in arr.Elements)
                            if (row is ObjectValue obj)
                            {
                                try { var lr = t.ListRows.Add(); int col = 1; foreach (var kv in obj.Properties) ((ExcelRange)lr.Range[1, col++]).Value2 = UnwrapV(kv.Value); }
                                catch { }
                            }
                }),
                ["clear"] = F("clear", _ => { try { t.DataBodyRange?.Delete(); } catch { } }),
            });
        }

    }
}
