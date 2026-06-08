# ALIGNMENT: ExcelScriptLoader v1 — 需求对齐分析

## 1. 原始需求

将 SereinScript（C# 脚本引擎）嵌入 Excel，使其像 VBA 一样在 Excel 进程内运行并操作工作表。交付物为一个 XLL 插件文件。

## 2. 项目理解

### 2.1 当前项目状态

| 项目 | 状态 |
|------|------|
| **ExcelScriptLoader** | 骨架项目 — .NET 10 Console App，仅含 `Program.cs`（Hello World） |
| **SereinScript (ScriptLang)** | 成熟的脚本引擎 — 词法分析 → 语法分析 → 字节码编译 → VM 执行 |

### 2.2 SereinScript 引擎 API 分析

```
ScriptEngine
├── CreateTask(string filePath, Scope? scope) → ScriptTask
├── CreateTask(ByteCodeChunk chunk, string? filePath) → ScriptTask
├── RegisterGlobal(string name)              // 预注册全局变量（编译前）
├── SetGlobal(string name, Value value)      // 设置全局变量值
├── GlobalScope: Scope                       // 全局作用域
├── ClearCache()
└── PrototypeManager                         // 原型扩展管理器

Scope
├── DefineClrObject(string name, object data) // 注入 CLR 对象（关键！）
├── Define(string name, Value value)
├── CreateChildScope() → Scope

ScriptTask
├── RunAsync() → Task<Value>
└── Cancel()

Value 子类型
├── NullValue, BoolValue, StringValue, NumberValue<T>
├── ObjectValue, ArrayValue
├── ClrObjectValue     ← 包装任意 C# 对象给脚本访问
├── ClrMethodValue     ← 包装 C# 方法
├── FunctionValue      ← 原生函数（多种委托签名）
└── CompiledFunctionValue
```

### 2.3 Excel 对象注入路径

```csharp
// 在插件启动时将 Excel 对象注入脚本作用域
var scope = new Scope(engine.GlobalScope);
scope.DefineClrObject("app", excelApplication);       // Excel.Application
scope.DefineClrObject("workbook", activeWorkbook);     // Active Workbook
scope.DefineClrObject("sheet", activeWorksheet);       // Active Worksheet

// 脚本中即可直接使用：
// app.Selection.Value = "Hello";
// workbook.Sheets.Count;
```

### 2.4 关键发现：SereinScript 缺少字符串代码执行 API

当前 `ScriptEngine.CreateTask()` 仅接受**文件路径**或**预编译 ByteCodeChunk**，不支持直接执行内存中的代码字符串。

**影响**：宏代码存储在 CustomXMLParts 中（字符串形式），需要一个转换层。

**备选方案**：
- **A**：在 ScriptEngine 中新增 `CreateTaskFromSource(string code, string sourceName, Scope? scope)` 方法
- **B**：利用 `SourceManager.AddSource(fakePath, code)` 预注册 + 用伪路径调用 CreateTask
- **C**：将宏代码写入临时 .script 文件再执行

## 3. 任务边界

### 包含范围
- Excel-DNA XLL 插件框架搭建
- CustomXMLParts 宏存储读写
- Ribbon 自定义标签页（WinForms 对话框）
- SereinScript 引擎适配层（含字符串执行 API）
- Excel 对象上下文注入（Application/Workbook/Worksheet/Range）
- 宏管理完整 CRUD

### 不包含范围
- VBA 代码注入（Alt+F8 注册）— 高风险，作为备选
- 数字签名（P2 优先级）
- 跨进程通信
- Excel 版本兼容性测试（仅针对 Excel 2019+/Microsoft 365）

## 4. 风险与假设

| 风险 | 等级 | 缓解 |
|------|------|------|
| .NET 10 运行时未安装 | 高 | 打包时检测并提示安装链接 |
| Excel-DNA 与 .NET 10 兼容性 | 中 | Excel-DNA 1.7+ 已支持 .NET 6/8，需验证 .NET 10 |
| SereinScript 字节码与 Excel 线程模型冲突 | 中 | 所有脚本执行在后台 Task 中，避免阻塞 Excel UI 线程 |
| CustomXMLParts 大小限制 | 低 | 单个脚本建议 <5MB |
| VBA Project Model 未启用 | 中 | Alt+F8 集成作为备选方案 |

## 5. 待确认问题

### Q1: 字符串执行 API 方案
SereinScript 目前不支持直接执行字符串代码。需要修改 ScriptLang 项目添加此能力。推荐方案 A（新增 `CreateTaskFromSource` 方法），侵入性最小。是否同意？

### Q2: Alt+F8 宏列表集成优先级
通过 VBProject 动态注入 VBA 模块可实现 Alt+F8 显示自定义宏，但需要用户启用"信任对 VBA 项目对象模型的访问"。是否接受此限制？或降级为 Ribbon-Only 方案？

### Q3: 项目框架选择
需求文档指定 .NET 10 + Excel-DNA。确认：
- Excel-DNA 版本：1.7+（支持 .NET 6+）
- UI 框架：WinForms（已指定 UseWindowsForms=true）
- 不使用 WPF

### Q4: SereinScript 引用方式
ExcelScriptLoader 如何引用 SereinScript？
- A) 项目引用（同解决方案）
- B) NuGet 包引用（需先发布 SereinScript 包）
- C) DLL 直接引用

推荐 A，因为两个项目在同一开发环境中。

### Q5: 脚本中 Excel 对象命名约定
需求文档中使用 `app.Selection.Value = "Hello"`。确认全局变量命名：
- `app` → Excel.Application
- `workbook` → 当前活动工作簿
- `sheet` → 当前活动工作表
- `selection` / `range` → 当前选中区域

是否需要额外暴露？（如 `cells`, `rows`, `columns`）
