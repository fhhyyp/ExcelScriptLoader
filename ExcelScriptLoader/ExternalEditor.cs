using System.Diagnostics;

namespace ExcelScriptLoader;

/// <summary>
/// 外部编辑器集成（VS Code + SereinScript LSP）
/// </summary>
public static class ExternalEditor
{
    private static string? _vsCodePath;

    /// <summary>VS Code 是否可用</summary>
    public static bool IsVsCodeAvailable => GetVsCodePath() != null;

    /// <summary>
    /// 在 VS Code 中编辑代码，等待关闭后返回修改后的内容
    /// </summary>
    /// <param name="initialCode">初始代码</param>
    /// <param name="macroName">宏名称（用作文件名）</param>
    /// <returns>修改后的代码，如果用户取消则返回 null</returns>
    public static string? EditInVsCode(string initialCode, string macroName)
    {
        var codePath = GetVsCodePath();
        if (codePath == null)
            return null; // VS Code 不可用

        try
        {
            // 创建临时 .script 文件（LSP 根据扩展名激活）
            var tempDir = Path.Combine(Path.GetTempPath(), "ExcelScriptLoader");
            Directory.CreateDirectory(tempDir);

            // 清理文件名中的非法字符
            var safeName = string.Join("_", macroName.Split(Path.GetInvalidFileNameChars()));
            if (string.IsNullOrWhiteSpace(safeName))
                safeName = "macro";

            var tempFile = Path.Combine(tempDir, $"{safeName}.script");

            // 写入初始代码
            File.WriteAllText(tempFile, initialCode);

            // 获取初始写入时间以检测修改
            var lastWriteTime = File.GetLastWriteTimeUtc(tempFile);

            // 启动 VS Code（--wait：等待用户关闭文件标签页）
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = codePath,
                    Arguments = $"--wait \"{tempFile}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                },
            };

            process.Start();
            process.WaitForExit();

            // 读取修改后的内容
            if (File.Exists(tempFile))
            {
                var newContent = File.ReadAllText(tempFile);

                // 清理临时文件
                try { File.Delete(tempFile); }
                catch { /* 忽略清理失败 */ }

                return newContent;
            }

            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ExternalEditor] VS Code 编辑失败: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 查找 VS Code 安装路径
    /// </summary>
    private static string? GetVsCodePath()
    {
        if (_vsCodePath != null)
            return _vsCodePath;

        // 按优先级查找
        string[] candidates =
        [
            // 1. PATH 中的 code 命令
            "code",

            // 2. 标准用户安装路径
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "Microsoft VS Code", "Code.exe"),

            // 3. 标准系统安装路径
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "Microsoft VS Code", "Code.exe"),

            // 4. 旧版安装路径
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                "Microsoft VS Code", "Code.exe"),

            // 5. Cursor (VS Code fork)
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "Cursor", "Cursor.exe"),
        ];

        foreach (var candidate in candidates)
        {
            try
            {
                // 如果是裸命令名，尝试通过 PATH 解析
                if (!candidate.Contains(Path.DirectorySeparatorChar) && !candidate.Contains('/'))
                {
                    var fullPath = FindInPath(candidate);
                    if (fullPath != null)
                    {
                        _vsCodePath = fullPath;
                        Debug.WriteLine($"[ExternalEditor] 找到 VS Code: {_vsCodePath}");
                        return _vsCodePath;
                    }
                }
                else if (File.Exists(candidate))
                {
                    _vsCodePath = candidate;
                    Debug.WriteLine($"[ExternalEditor] 找到 VS Code: {_vsCodePath}");
                    return _vsCodePath;
                }
            }
            catch
            {
                // 继续尝试下一个
            }
        }

        Debug.WriteLine("[ExternalEditor] 未找到 VS Code");
        return null;
    }

    /// <summary>在 PATH 环境变量中查找可执行文件</summary>
    private static string? FindInPath(string fileName)
    {
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        var extensions = Environment.GetEnvironmentVariable("PATHEXT")?.Split(';') ?? [".exe"];

        foreach (var dir in pathEnv.Split(';'))
        {
            foreach (var ext in extensions)
            {
                var fullPath = Path.Combine(dir.Trim(), fileName + ext.Trim());
                if (File.Exists(fullPath))
                    return fullPath;
            }
        }

        return null;
    }
}
