// ============================================================================
// Excel 标准库 — ObjectValue + FunctionValue 模式（参考 TimerModule）
// 零反射、零原型依赖，直接委托调用
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang;
using ScriptLang.Runtime;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{

// ============================================================================
// 1. ExcelModule — 顶层入口，通过 ClrObjectValue 注入 "excel" 模块
// ============================================================================

[PrototypeExtension(PushThis = true)]
public partial class ExcelModule
{
    private readonly ExcelApplication _app;
    public ExcelModule(ExcelApplication app) => _app = app;

    public partial bool IsTarget(Value value) =>
        value is ClrObjectValue clr && clr.Value is ExcelModule;

    // ---- 属性 ----

    [PrototypeProperty(Name = "active")]
    public static Value active(ExcelModule self)
    {
        try { return WorkbookToObject(self._app.ActiveWorkbook); }
        catch { return Value.Null; }
    }

    // ---- 方法 ----

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

    [PrototypeFunction(Name = "cell")]
    public static Value cell(ExcelModule self)
    {
        try { return CellToObject(self._app.ActiveCell); }
        catch { return Value.Null; }
    }

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

    [PrototypeFunction(Name = "create")]
    public static Value create(ExcelModule self)
    {
        try { return WorkbookToObject(self._app.Workbooks.Add()); }
        catch { return Value.Null; }
    }

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

    // ---- 内部工厂 ----

    internal static ObjectValue WorkbookToObject(Workbook wb)
    {
        var w = wb; // capture
        return new ObjectValue(new()
        {
            ["name"] = StringValue.Create(w.Name),
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

    internal static ObjectValue WorksheetToObject(Worksheet ws)
    {
        var w = ws;
        return new ObjectValue(new()
        {
            ["name"] = StringValue.Create(w.Name),
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
        });
    }

    internal static ObjectValue CellToObject(ExcelRange cell)
    {
        var c = cell;
        return new ObjectValue(new()
        {
            ["value"] = F("value", args =>
            {
                if (args.Count == 0)
                {
                    try { return WrapCell(c.Value2); } catch { return Value.Null; }
                }
                else
                {
                    try { c.Value2 = UnwrapV(args[0]); } catch { }
                    return Value.Null;
                }
            }),
            ["formula"] = F("formula", args =>
            {
                if (args.Count == 0)
                {
                    try { return StringValue.Create(c.Formula as string ?? ""); } catch { return Value.Null; }
                }
                else if (args[0] is StringValue f)
                {
                    try { c.Formula = f.Value; } catch { }
                }
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
                else if (args[0] is ObjectValue fmt)
                {
                    try { ApplyFormat(c, fmt); } catch { }
                }
                return Value.Null;
            }),
        });
    }

    internal static ObjectValue RangeToObject(ExcelRange rng)
    {
        var r = rng;
        return new ObjectValue(new()
        {
            ["values"] = F("values", args =>
            {
                if (args.Count == 0)
                {
                    try { return ReadRange(r); } catch { return Value.Null; }
                }
                else
                {
                    try { WriteRange(r, args[0]); } catch { }
                    return Value.Null;
                }
            }),
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
            ["copy"] = F("copy", args =>
            {
                if (args.Count > 0 && args[0] is StringValue t)
                    try { r.Copy(r.Worksheet.Range[t.Value]); } catch { }
            }),
        });
    }

    internal static ObjectValue TableToObject(ListObject table)
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

                    var headerRng = t.HeaderRowRange as ExcelRange;
                    var headerData = (object[,]?)headerRng?.Value2;
                    var headers = new string[cols];
                    for (int c = 1; c <= cols; c++)
                        headers[c - 1] = headerData?[1, c]?.ToString() ?? $"Col{c}";

                    for (int rr = 1; rr <= rows; rr++)
                    {
                        var row = new Dictionary<string, Value>();
                        for (int cc = 1; cc <= cols; cc++)
                            row[headers[cc - 1]] = WrapCell(data[rr, cc]);
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
                    try
                    {
                        var lr = t.ListRows.Add();
                        int col = 1;
                        foreach (var kv in obj.Properties)
                            ((ExcelRange)lr.Range[1, col++]).Value2 = UnwrapV(kv.Value);
                    }
                    catch { }
                }
            }),
            ["addAll"] = F("addAll", args =>
            {
                if (args.Count > 0 && args[0] is ArrayValue arr)
                {
                    foreach (var row in arr.Elements)
                    {
                        if (row is ObjectValue obj)
                        {
                            try
                            {
                                var lr = t.ListRows.Add();
                                int col = 1;
                                foreach (var kv in obj.Properties)
                                    ((ExcelRange)lr.Range[1, col++]).Value2 = UnwrapV(kv.Value);
                            }
                            catch { }
                        }
                    }
                }
            }),
            ["clear"] = F("clear", _ => { try { t.DataBodyRange?.Delete(); } catch { } }),
        });
    }

    // ---- 内部工具 ----

    internal static FunctionValue F(string name, Func<List<Value>, Value> func)
        => new(name, func);

    internal static FunctionValue F(string name, Action<List<Value>> action)
        => new(name, args => { action(args); return Value.Null; });

    internal static Value WrapCell(object? v) => v switch
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

    internal static object UnwrapV(Value v) => v switch
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

    internal static Value ReadRange(ExcelRange rng)
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

    internal static void WriteRange(ExcelRange rng, Value data)
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

    internal static void WriteObjects(ExcelRange rng, Value data)
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

    internal static void ApplyFormat(ExcelRange cell, ObjectValue fmt)
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

    internal static object ParseColor(Value c) => c switch
    {
        StringValue sv => sv.Value.ToLowerInvariant() switch
        {
            "red" => 0xFF0000, "green" => 0x00FF00, "blue" => 0x0070C0,
            "yellow" => 0xFFFF00, "white" => 0xFFFFFF, "black" => 0x000000,
            "gray" or "grey" => 0x808080, "orange" => 0xFF8C00, _ => 0
        },
        NumberValue<int> ni => ni.Value,
        NumberValue<double> nd => (int)nd.Value,
        _ => 0,
    };
}
}
