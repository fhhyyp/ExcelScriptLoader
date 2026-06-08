using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;

namespace ExcelScriptLoader
{

/// <summary>
/// CustomXMLParts 存储层 — 使用纯反射 COM 调用，零 office.dll 依赖
/// </summary>
public static class MacroStorage
{
    private const string NsUri = "http://sereinscript/excel-macros/v1";

    private static bool _verbose = false;

    private static void Info(string msg)
    {
        Debug.WriteLine($"[MacroStorage] {msg}");
        if (_verbose) MessageBox.Show(msg, "MacroStorage 诊断", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static void Error(string msg)
    {
        Debug.WriteLine($"[MacroStorage] ERROR: {msg}");
        MessageBox.Show(msg, "MacroStorage 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    // ==================== 公开 API ====================

    public static List<ExcelMacro> LoadMacros(Workbook workbook)
    {
        var macros = new List<ExcelMacro>();

        try
        {
            var customXmlParts = GetCustomXmlParts(workbook);
            if (customXmlParts == null)
            {
                Info("LoadMacros: CustomXMLParts=null (正常，工作簿尚无宏数据)");
                return macros;
            }

            // SelectByNamespace 通过反射 COM 不可靠（返回全部部件）→ 遍历全部手动过滤
            int totalCount = (int)GetProperty(customXmlParts, "Count");
            Debug.WriteLine($"[MacroStorage] 遍历全部 {totalCount} 个部件查找宏数据...");
            int readOk = 0, readFail = 0;

            for (int i = 1; i <= totalCount; i++)
            {
                try
                {
                    var part = GetIndexedItem(customXmlParts, i);
                    var xml = (string)GetProperty(part, "XML");
                    readOk++;
                    // 快速过滤：XML 中必须包含我们的命名空间
                    if (!xml.Contains(NsUri)) continue;

                    var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
                    XNamespace ns = NsUri;
                    foreach (var me in doc.Root?.Elements(ns + "macro") ?? [])
                    {
                        try { macros.Add(ExcelMacro.FromXElement(me)); }
                        catch (Exception ex) { Debug.WriteLine($"[MacroStorage] 损坏的宏: {ex.Message}"); }
                    }
                }
                catch (Exception ex) { readFail++; Debug.WriteLine($"[MacroStorage] Part[{i}] 跳过: {ex.Message}"); }
            }
            Debug.WriteLine($"[MacroStorage] 读取成功 {readOk}, 失败 {readFail}, 宏 {macros.Count}");
        }
        catch (Exception ex)
        {
            Error($"LoadMacros 失败: {ex.Message}\n{ex.StackTrace}");
            _verbose = false; // 不再弹窗
        }

        Info($"LoadMacros 完成: 共 {macros.Count} 个宏");
        return macros;
    }

    public static void SaveMacros(Workbook workbook, List<ExcelMacro> macros)
    {
        if (workbook == null) { Error("SaveMacros: workbook 为 null"); return; }

        try
        {
            if (macros.Count == 0)
            {
                Info("SaveMacros: 宏列表为空，仅清理旧部件");
                DeleteAllMacroParts(workbook);
                return;
            }

            // 1. 清理旧部件
            DeleteAllMacroParts(workbook);

            // 2. 构建 XML
            XNamespace ns = NsUri;
            var root = new XElement(ns + "macros");
            foreach (var m in macros)
            {
                m.ModifiedAt = DateTime.UtcNow;
                root.Add(m.ToXElement());
            }

            var xmlString = new XDocument(new XDeclaration("1.0", "utf-8", null), root).ToString();
            Info($"SaveMacros: XML 已构建 ({xmlString.Length} 字符)\n: {xmlString}");

            // 3. 获取 CustomXMLParts
            var customXmlParts = GetCustomXmlParts(workbook);
            if (customXmlParts == null)
            {
                Error("SaveMacros: 无法获取 CustomXMLParts（可能需要先保存工作簿）");
                return;
            }

            // 4. 添加新部件
            var newPart = InvokeMethod(customXmlParts, "Add", xmlString);
            if (newPart == null)
            {
                Error("SaveMacros: CustomXMLParts.Add 返回 null（添加失败）");
                return;
            }

            Info($"SaveMacros: 成功保存 {macros.Count} 个宏到工作簿 \"{workbook.Name}\"");
        }
        catch (Exception ex)
        {
            Error($"SaveMacros 失败: {ex.Message}\n\n详细:\n{ex}");
        }
    }

    public static bool DeleteMacro(Workbook workbook, string macroId)
    {
        var macros = LoadMacros(workbook);
        var removed = macros.RemoveAll(m => m.Id == macroId);
        if (removed > 0) { SaveMacros(workbook, macros); return true; }
        return false;
    }

    public static bool NameExists(Workbook workbook, string name, string? excludeId = null)
        => LoadMacros(workbook).Any(m =>
            string.Equals(m.Name, name.Trim(), StringComparison.OrdinalIgnoreCase) && m.Id != excludeId);

    public static ExcelMacro? FindById(Workbook workbook, string id)
        => LoadMacros(workbook).FirstOrDefault(m => m.Id == id);

    // ==================== 反射 COM 访问 ====================

    private static object? GetCustomXmlParts(Workbook workbook)
    {
        try
        {
            return workbook.GetType().InvokeMember(
                "CustomXMLParts",
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null, workbook, null);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[MacroStorage] GetCustomXmlParts 异常: {ex.Message}");
            return null;
        }
    }

    private static object? InvokeMethod(object target, string methodName, params object?[] args)
    {
        try
        {
            return target.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                null, target, args);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[MacroStorage] InvokeMethod({methodName}): {ex.Message}");
            return null;
        }
    }

    private static object GetProperty(object target, string propertyName)
    {
        return target.GetType().InvokeMember(
            propertyName,
            BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
            null, target, null)!;
    }

    private static object GetIndexedItem(object target, int index)
    {
        // 尝试 default member（COM 默认索引器，如 parts(1)）
        try
        {
            return target.GetType().InvokeMember(
                "",
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null, target, [index])!;
        }
        catch
        {
            // 回退 Item 方法
            return target.GetType().InvokeMember(
                "Item",
                BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                null, target, [index])!;
        }
    }

    private static void DeleteAllMacroParts(Workbook workbook)
    {
        try
        {
            var customXmlParts = GetCustomXmlParts(workbook);
            if (customXmlParts == null) return;

            // SelectByNamespace 不可靠，改为遍历全部并删除匹配命名空间的部件
            int totalCount = (int)GetProperty(customXmlParts, "Count");
            var toDelete = new List<int>();
            for (int i = 1; i <= totalCount; i++)
            {
                try
                {
                    var xml = (string)GetProperty(GetIndexedItem(customXmlParts, i), "XML");
                    if (xml.Contains(NsUri)) toDelete.Add(i);
                }
                catch { }
            }
            Debug.WriteLine($"[MacroStorage] DeleteAll: 删除 {toDelete.Count} 个旧部件");
            for (int j = toDelete.Count - 1; j >= 0; j--)
            {
                try { InvokeMethod(GetIndexedItem(customXmlParts, toDelete[j]), "Delete"); }
                catch (Exception ex) { Debug.WriteLine($"[MacroStorage] Delete: {ex.Message}"); }
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[MacroStorage] DeleteAll: {ex.Message}"); }
    }

    private static class Debug { public static void WriteLine(string msg) => System.Diagnostics.Debug.WriteLine(msg); }
}
}
