// ============================================================================
// ExcelModule.Table — Table API 工厂
// ObjectValue + FunctionValue 模式，将 Excel COM ListObject 包装为脚本对象
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang.Runtime;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{
    public partial class ExcelModule
    {

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

        internal static ObjectValue TableToObject(ListObject table)
        {
            var t = table;
            return new ObjectValue(new()
            {
                // -- 属性 --
                ["name"] = StringValue.Create(t.Name),

                // -- 读取所有行（对象数组，key=列名） --
                ["rows"] = F("rows", _ =>
                {
                    var result = new List<Value>();
                    try
                    {
                        // 优先使用 DataBodyRange（数据区域），回退到整个 Range
                        var rng = (ExcelRange?)(t.DataBodyRange ?? t.Range);
                        if (rng == null) return new ArrayValue(result);

                        var data = (object[,])rng.Value2;
                        int rows = data.GetLength(0), cols = data.GetLength(1);

                        // 读取表头
                        var headerRng = t.HeaderRowRange as ExcelRange;
                        var headerData = (object[,]?)headerRng?.Value2;
                        var headers = new string[cols];
                        for (int c = 1; c <= cols; c++)
                            headers[c - 1] = headerData?[1, c]?.ToString() ?? $"Col{c}";

                        // 将每行映射为 {列名: 值} 对象
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

                // -- 新增一行 --
                ["add"] = F("add", args =>
                {
                    if (args.Count > 0 && args[0] is ObjectValue obj)
                    {
                        try
                        {
                            var lr = t.ListRows.Add();
                            int col = 1;
                            // 按对象属性的插入顺序写入列（需与表头顺序一致）
                            foreach (var kv in obj.Properties)
                                ((ExcelRange)lr.Range[1, col++]).Value2 = UnwrapV(kv.Value);
                        }
                        catch { }
                    }
                }),

                // -- 批量新增 --
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

                // -- 清空数据 --
                ["clear"] = F("clear", _ =>
                {
                    try { t.DataBodyRange?.Delete(); } catch { }
                }),
            });
        }

    }
}
