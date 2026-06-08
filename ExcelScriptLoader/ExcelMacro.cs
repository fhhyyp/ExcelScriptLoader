using System.Xml.Linq;

namespace ExcelScriptLoader;

/// <summary>
/// 宏数据模型，对应 CustomXMLParts 中的一条 macro 记录
/// </summary>
public class ExcelMacro
{
    /// <summary>GUID，唯一标识</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("B").ToUpperInvariant();

    /// <summary>宏名称（显示用，同一工作簿内唯一）</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>描述（可选）</summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>SereinScript 脚本代码</summary>
    public string ScriptCode { get; set; } = string.Empty;

    /// <summary>快捷键（预留）</summary>
    public string ShortcutKey { get; set; } = string.Empty;

    /// <summary>创建时间</summary>
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    /// <summary>修改时间</summary>
    public DateTime ModifiedAt { get; set; } = DateTime.UtcNow;

    // ==================== XML 序列化 ====================

    private const string Ns = "http://sereinscript/excel-macros/v1";

    /// <summary>从 XML Element 反序列化</summary>
    public static ExcelMacro FromXElement(XElement element)
    {
        XName xn(string local) => XName.Get(local, Ns);

        return new ExcelMacro
        {
            Id = element.Element(xn("id"))?.Value ?? Guid.NewGuid().ToString("D"),
            Name = element.Element(xn("name"))?.Value ?? "未命名宏",
            Description = element.Element(xn("description"))?.Value ?? string.Empty,
            ScriptCode = ReadCodeLines(element.Element(xn("code"))),
            ShortcutKey = element.Element(xn("shortcutKey"))?.Value ?? string.Empty,
            CreatedAt = ParseDateTime(element.Element(xn("createdAt"))?.Value),
            ModifiedAt = ParseDateTime(element.Element(xn("modifiedAt"))?.Value),
        };
    }

    /// <summary>序列化为 XML Element</summary>
    public XElement ToXElement()
    {
        XName xn(string local) => XName.Get(local, Ns);

        return new XElement(xn("macro"),
            new XElement(xn("id"), Id),
            new XElement(xn("name"), Name),
            new XElement(xn("description"), Description),
            new XElement(xn("code"),
                ScriptCode.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').Select(line =>
                    new XElement(xn("line"), line))),
            new XElement(xn("shortcutKey"), ShortcutKey),
            new XElement(xn("createdAt"), CreatedAt.ToString("O")),
            new XElement(xn("modifiedAt"), ModifiedAt.ToString("O"))
        );
    }

    /// <summary>读取按行存储的代码（新格式: &lt;line&gt;...&lt;/line&gt;；兼容旧 CDATA）</summary>
    private static string ReadCodeLines(XElement? codeElement)
    {
        if (codeElement == null) return string.Empty;

        XName xn(string local) => XName.Get(local, Ns);

        // 新格式：<line> 元素
        var lines = codeElement.Elements(xn("line")).Select(e => e.Value).ToList();
        if (lines.Count > 0)
            return string.Join("\r\n", lines);

        // 兼容旧格式：CDATA
        var cdata = codeElement.Nodes().OfType<XCData>().FirstOrDefault();
        if (cdata != null)
            return cdata.Value.Replace("\r\n", "\n").Replace("\r", "\n");

        return codeElement.Value.Replace("\r\n", "\n").Replace("\r", "\n");
    }

    private static DateTime ParseDateTime(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return DateTime.UtcNow;
        return DateTime.TryParse(value, null, System.Globalization.DateTimeStyles.RoundtripKind, out var result)
            ? result
            : DateTime.UtcNow;
    }

    /// <summary>验证宏名称有效性</summary>
    public static string? ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "宏名称不能为空";
        if (name.Trim().Length == 0)
            return "宏名称不能为纯空格";
        if (name.IndexOfAny(['<', '>', '&', '"', '\'']) >= 0)
            return "宏名称包含非法字符（< > & \" '）";
        return null; // 验证通过
    }

    public override string ToString() => $"{Name} ({Id[..8]}...)";
}
