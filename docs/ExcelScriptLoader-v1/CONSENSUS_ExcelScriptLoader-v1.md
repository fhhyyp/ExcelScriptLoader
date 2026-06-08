# CONSENSUS: ExcelScriptLoader v1 — 需求共识

## 需求描述

将 SereinScript C# 脚本引擎嵌入 Excel，通过 Excel-DNA XLL 插件实现宏的创建、存储、编辑与运行。宏代码存储在 Excel 文件的 CustomXMLParts 中，随工作簿自动保存和加载。

## 验收标准

1. ✅ 用户可通过 Ribbon 菜单新建/编辑/删除/运行宏
2. ✅ 宏代码保存在工作簿 CustomXMLParts 中，关闭后重新打开仍存在
3. ✅ 脚本可访问 `app`/`workbook`/`sheet`/`cell`/`selection` 等 Excel 对象
4. ✅ 脚本执行错误时返回清晰的错误信息
5. ✅ 插件作为 XLL 在 Excel 启动时自动加载

## 技术方案

### 技术栈
| 层 | 技术选型 |
|-----|----------|
| 插件框架 | Excel-DNA 1.7+ |
| 运行时 | .NET 10.0-windows |
| 脚本引擎 | SereinScript (ScriptLang) — 项目引用 |
| UI 框架 | WinForms |
| 宏存储 | CustomXMLParts（XML 格式） |
| 宏触发 | Ribbon 菜单（非 Alt+F8 / VBA 注入） |

### 核心架构决策

**D1: 脚本执行流程**
```
宏代码(字符串) → ScriptEngine.CreateTaskFromSource() [新增API]
  → Lexer → Parser → Compiler → ByteCodeChunk → VM.ExecuteAsync() → Value
```
- 在 ScriptLang 中新增 `ScriptEngine.CreateTaskFromSource(string source, string sourceName, Scope? scope)` 方法
- 利用已有 `SourceManager.AddSource()` 预注册 + 复用现有编译管道

**D2: Excel 对象注入**
```csharp
var scope = new Scope(engine.GlobalScope);
scope.DefineClrObject("app", Globals.ThisAddIn.Application);
scope.DefineClrObject("workbook", activeWorkbook);
scope.DefineClrObject("sheet", activeWorksheet);
scope.DefineClrObject("selection", application.Selection);
scope.DefineClrObject("cell", application.ActiveCell);
```
- 全局变量名：`app`, `workbook`, `sheet`, `cell`, `selection`
- 通过 `ClrObjectValue` 包装，脚本可直接访问 .NET 属性与方法

**D3: 宏触发方式**
- 仅 Ribbon 菜单触发（非 Alt+F8 VBA 注入）
- Ribbon 标签页名称：**"脚本宏"**

**D4: 项目结构**
- ExcelScriptLoader 引用 ScriptLang 项目（Solution 级别）
- 两个项目在同一个 Solution 中

## 技术约束

- TargetFramework: `net10.0-windows`
- UseWindowsForms: `true`
- PlatformTarget: `x64`
- Excel-DNA 通过 NuGet 安装
- 不依赖 VBA Project Model

## 集成方案

```
ExcelScriptLoader.sln
├── ExcelScriptLoader/          # Excel-DNA XLL 插件
│   ├── AddIn.cs                # IExcelAddIn 入口
│   ├── MacroStorage.cs         # CustomXMLParts 读写
│   ├── ExcelMacro.cs           # 宏数据模型
│   ├── RibbonController.cs     # 自定义 Ribbon
│   ├── ScriptEngineAdapter.cs  # 脚本引擎适配层
│   └── Dialogs/                # WinForms 对话框
└── ScriptLang/                 # SereinScript 引擎（已有）
    └── ScriptEngine.cs         # 需新增 CreateTaskFromSource()
```
