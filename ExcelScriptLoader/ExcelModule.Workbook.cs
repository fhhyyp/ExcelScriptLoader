// ============================================================================
// ExcelModule.Workbook — Workbook API 工厂
// ObjectValue + FunctionValue 模式，将 Excel COM Workbook 包装为脚本对象
// ============================================================================

using Microsoft.Office.Interop.Excel;
using ScriptLang.Runtime;

namespace ExcelScriptLoader
{
    public partial class ExcelModule
    {

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

        internal static ObjectValue WorkbookToObject(Workbook wb)
        {
            var w = wb; // 闭包捕获，避免 lambda 中引用参数
            return new ObjectValue(new()
            {
                // -- 属性 --
                ["name"] = StringValue.Create(w.Name),
                ["path"] = StringValue.Create(w.Path ?? ""),
                ["fullName"] = StringValue.Create(w.FullName ?? ""),

                // -- 方法 --
                ["save"] = F("save", _ =>
                {
                    try { w.Save(); } catch { }
                }),
                ["saveAs"] = F("saveAs", args =>
                {
                    if (args.Count > 0 && args[0] is StringValue p)
                        try { w.SaveAs(p.Value); } catch { }
                }),
                ["close"] = F("close", _ =>
                {
                    try { w.Close(false); } catch { }
                }),
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
                    try
                    {
                        foreach (Worksheet ws in w.Sheets)
                            list.Add(WorksheetToObject(ws));
                    }
                    catch { }
                    return new ArrayValue(list);
                }),
            });
        }

    }
}
