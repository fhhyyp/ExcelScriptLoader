// ============================================================================
// ExcelModule.Utilities — 内部工具方法
// 值转换、区域读写、格式化等底层操作，被各工厂方法共用
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang;
using ScriptLang.Runtime;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace ExcelScriptLoader
{
    public partial class ExcelModule
    {

        // ============================================================================
        // 工具: F() — FunctionValue 工厂
        // ============================================================================

        /// <summary>
        /// 创建有返回值的 FunctionValue
        /// </summary>
        internal static FunctionValue F(string name, Func<List<Value>, Value> func)
            => new(name, func);

        /// <summary>
        /// 创建无返回值的 FunctionValue（void action）
        /// </summary>
        internal static FunctionValue F(string name, Action<List<Value>> action)
            => new(name, args => { action(args); return Value.Null; });

        // ============================================================================
        // 工具: WrapCell / UnwrapV — ScriptLang ↔ COM 值转换
        // ============================================================================

        /// <summary>
        /// 将 COM 值（object）转换为 ScriptLang Value。
        /// 支持: null → Null, string → StringValue,
        ///       int/long/double/float → NumberValue, bool → BoolValue,
        ///       其他 → StringValue(ToString)
        /// </summary>
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

        /// <summary>
        /// 将 ScriptLang Value 转换为 COM 可用的 object。
        /// NullValue → ""（空字符串），不可识别的类型 → ToString()
        /// </summary>
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

        // ============================================================================
        // 工具: ReadRange / WriteRange / WriteObjects — 区域数据读写
        // ============================================================================

        /// <summary>
        /// 将 ExcelRange 的 Value2 读取为二维 ArrayValue。
        /// 内部使用 object[,] 中转以提升批量读取性能。
        /// </summary>
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

        /// <summary>
        /// 将二维 ArrayValue 写入 ExcelRange。
        /// 自动调用 Resize 以匹配数据尺寸，然后通过 Value2 批量赋值。
        /// </summary>
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

        /// <summary>
        /// 将对象数组写入 ExcelRange，自动生成表头行。
        /// 第一行 = 对象 key 集合，后续行 = 对象 value。
        /// </summary>
        internal static void WriteObjects(ExcelRange rng, Value data)
        {
            if (data is not ArrayValue av || av.Elements.Count == 0) return;
            if (av.Elements[0] is not ObjectValue first) return;

            var keys = first.Properties.Keys.ToArray();
            int rows = av.Elements.Count + 1;
            int cols = keys.Length;

            var arr = new object[rows, cols];
            // 表头行
            for (int c = 0; c < cols; c++) arr[0, c] = keys[c];
            // 数据行
            for (int r = 0; r < av.Elements.Count; r++)
            {
                if (av.Elements[r] is not ObjectValue obj) continue;
                for (int c = 0; c < cols; c++)
                    arr[r + 1, c] = obj.Properties.TryGetValue(keys[c], out var v)
                        ? UnwrapV(v) : "";
            }
            rng.Resize[rows, cols].Value2 = arr;
        }

        // ============================================================================
        // 工具: ApplyFormat / ParseColor — 格式应用
        // ============================================================================

        /// <summary>
        /// 将脚本侧格式对象应用到 ExcelRange。
        ///
        /// 支持字段:
        ///   bold         (bool)         加粗
        ///   color        (str|int)      字体色 — 颜色名或 RGB 整数
        ///   bgColor      (str|int)      背景色
        ///   size         (number)       字号
        ///   numberFormat (str)          数字格式 (如 "0.00%")
        /// </summary>
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

        /// <summary>
        /// 将脚本侧颜色值转换为 COM 颜色整数。
        ///
        /// 支持: 颜色名字符串 ("red", "green", "blue", "yellow", "white",
        ///        "black", "gray", "grey", "orange") 或 RGB 整数值。
        /// </summary>
        internal static object ParseColor(Value c) => c switch
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

    }
}
