# 🚀 ExcelScriptLoader

**Excel C# 脚本引擎插件** — 像 VBA 一样在 Excel 中运行脚本，但用 JavaScript 风格语法和现代 C# 引擎。

<p align="center">
  <img src="https://img.shields.io/badge/.NET-10.0-512BD4?logo=dotnet" />
  <img src="https://img.shields.io/badge/Excel--DNA-1.9-217346?logo=microsoft-excel" />
  <img src="https://img.shields.io/badge/SereinScript-v1.0-blue" />
  <img src="https://img.shields.io/badge/license-MIT-green" />
  <img src="https://img.shields.io/badge/QQ%E7%BE%A4-955830545-eb1923?logo=tencent-qq" />
</p>

---

## 📖 项目介绍

**ExcelScriptLoader** 是一个 Excel XLL 插件，让 [SereinScript](https://github.com/your-org/SereinScript)（基于 .NET 的 JS 风格动态脚本语言）能够嵌入 Excel 进程内运行，直接操作工作表、单元格、表格等 Excel 对象。

> 类似 VBA，但语法更现代，开发体验更好。

### 为什么选择 ExcelScriptLoader？

| 特性 | 说明 |
|------|------|
| 🚀 **现代语法** | JavaScript 风格，支持 lambda、闭包、解构 |
| 📦 **随文件保存** | 脚本存储在 xlsx 的 CustomXMLParts 中，无需额外文件 |
| ⚡ **编译执行** | 字节码编译 + VM 执行，非解释执行 |
| 🔧 **VS Code 集成** | 可选在 VS Code 中编辑脚本，借助 LSP 获得语法高亮和自动补全 |
| 🪶 **轻量** | 基于 Excel-DNA，纯 C# 实现 |

### 快速体验

```js
import { excel } from "excel"

// 读取数据
let data = excel.sheet("销售数据").read()

// 数据汇总
let summary = {}
for i in range(1, len(data)) {
    let dept = data[i][1]
    let amt  = data[i][3]
    summary[dept] = (summary[dept] ?? 0) + amt
}

// 写入结果
let ws = excel.active.addSheet("汇总")
ws.write("A1", [["部门","合计"]] + ...)
excel.active.save()
```

---

## 🏗️ 项目架构

```
ExcelScriptLoader.slnx
├── ExcelScriptLoader/            ← Excel-DNA XLL 插件
│   ├── AddIn.cs                  IExcelAddIn 入口，生命周期管理
│   ├── ExcelModule.cs            Excel 脚本 API（ObjectValue+FunctionValue 模式）
│   ├── ExcelMacro.cs             宏数据模型 + XML 序列化
│   ├── MacroStorage.cs           CustomXMLParts 存储层
│   ├── ScriptEngineAdapter.cs    脚本引擎适配 + Console 重定向
│   ├── RibbonController.cs       自定义 Ribbon 标签页 "脚本宏"
│   ├── ExternalEditor.cs         VS Code 外部编辑器集成
│   ├── Dialogs/
│   │   ├── MacroEditorDialog.cs  宏编辑器
│   │   ├── MacroSelectorDialog.cs 宏选择器（运行/编辑/删除）
│   │   └── OutputWindow.cs       脚本输出窗口
│   └── ExcelScriptLoader-AddIn.dna  Excel-DNA 清单
│
└── ScriptLang/                   ← SereinScript 脚本引擎
    ├── ScriptEngine.cs            引擎入口 + CreateTaskFromSource()
    ├── Runtime/ImportResolver.cs  模块解析 + RegisterBuiltinModule()
    └── System/*Module.cs          Timer/File/Console 等内置模块
```

### 技术栈

| 层 | 技术 |
|----|------|
| 插件宿主 | Excel-DNA 1.9 |
| 运行时 | .NET 10.0-windows |
| 脚本引擎 | SereinScript（AST→字节码→VM） |
| UI | WinForms |
| 宏存储 | CustomXMLParts（随 xlsx 自动保存） |

---

## 🚀 如何使用

### 安装

1. 下载 `ExcelScriptLoader-AddIn64-packed.xll`（[Releases](https://github.com/your-org/ExcelScriptLoader/releases)）
2. 打开 Excel → 文件 → 选项 → 加载项
3. 管理：Excel 加载项 → 转到 → 浏览
4. 选择 `.xll` 文件 → 确定

> 需要安装 [.NET 10 Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)

### 基本用法

1. 打开任意 xlsx 工作簿
2. 点击 Ribbon 标签页 **"脚本宏"** → **"新建宏"**
3. 输入宏名称和脚本代码：

```js
import { excel } from "excel"

let cell = excel.cell()
cell.value("Hello, Excel Script!")
```

4. 点击 **"运行"** → 当前单元格显示 `Hello, Excel Script!`
5. 点击 **"保存"** → 宏随工作簿保存

### Ribbon 菜单

| 按钮 | 功能 |
|------|------|
| 🆕 新建宏 | 打开编辑器创建新宏 |
| ✏️ 编辑宏 | 查看并编辑已有宏 |
| ▶️ 运行宏 | 选择宏并执行 |
| 🗑️ 删除宏 | 删除选中的宏 |
| 📋 宏列表 | 查看当前工作簿所有宏 |
| 📤 输出 | 打开脚本输出窗口 |

---

## 📚 Excel API

### 导入

```js
import { excel } from "excel"
```

### 顶层对象

| API | 说明 |
|-----|------|
| `excel.active` | 当前宿主工作簿 |
| `excel.sheet(name?)` | 获取工作表 |
| `excel.cell()` | 当前活动单元格 |
| `excel.selection()` | 当前选中区域 |
| `excel.open(path)` | 打开外部文件 |
| `excel.read(path)` | 快捷读取外部文件 |
| `excel.write(path, data)` | 快捷写入外部文件 |

### 对象链

```
excel.active          → Workbook
  .sheet("数据")       → Worksheet
    .cell("A1")        → Cell
      .value("Hello")   → 设置值
      .format({...})    → 设置格式
```

### 输出调试

使用 `print()` 输出日志，结果显示在**输出窗口**中：

```js
print("开始处理...")
print("共 " + len(data) + " 行")
```

> 完整 API 文档：[docs/ExcelScriptLoader-v1/API_REFERENCE.md](docs/ExcelScriptLoader-v1/API_REFERENCE.md)

---

## 🔧 如何二次开发

### 环境要求

- .NET 10.0 SDK
- Visual Studio 2022+ 或 VS Code + C# Dev Kit
- Microsoft Excel 2019+ (64-bit)
- SereinScript 项目（项目引用）

### 克隆并构建

```bash
git clone https://github.com/your-org/ExcelScriptLoader
git clone https://github.com/your-org/SereinScript

# 构建（关闭 Excel 后再执行）
cd ExcelScriptLoader
dotnet build

# 输出: ExcelScriptLoader/bin/Debug/net10.0-windows/publish/ExcelScriptLoader-AddIn64-packed.xll
```

### 项目结构速览

| 文件 | 职责 |
|------|------|
| `AddIn.cs` | 插件入口：AutoOpen/AutoClose、工作簿事件 |
| `ExcelModule.cs` | 核心 API：所有 `excel.xxx()` 方法的实现 |
| `MacroStorage.cs` | 宏持久化：CustomXMLParts 读写 |
| `ScriptEngineAdapter.cs` | 引擎适配：编译执行、Console 重定向 |
| `RibbonController.cs` | Ribbon UI 回调 |
| `Dialogs/` | WinForms 对话框 |

### 修改 API

1. 编辑 `ExcelModule.cs`：内部 `XxxToObject()` 工厂方法将 COM 对象包装为 `ObjectValue`
2. 每个方法通过 `F("name", args => ...)` 创建 `FunctionValue`
3. 无需修改其他文件——API 通过原型分派自动暴露给脚本

### VS Code 集成

插件自动检测 VS Code 安装。编辑脚本时点击 **"VS Code"** 按钮：
- 代码保存为临时 `.script` 文件
- VS Code 打开（LSP 扩展提供语法高亮、补全）
- 关闭标签页后代码自动回填

---

## 📂 文档目录

| 文档 | 说明 |
|------|------|
| [API_REFERENCE.md](docs/ExcelScriptLoader-v1/API_REFERENCE.md) | Excel 脚本 API 完整参考 |
| [API_DEVPLAN.md](docs/ExcelScriptLoader-v1/API_DEVPLAN.md) | API 开发计划（v1.1 路线图） |
| [API开发指导.md](docs/ExcelScriptLoader-v1/API开发指导.md) | API 设计规范与开发指南 |
| [ALIGNMENT_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/ALIGNMENT_ExcelScriptLoader-v1.md) | 需求对齐分析 |
| [CONSENSUS_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/CONSENSUS_ExcelScriptLoader-v1.md) | 需求共识文档 |
| [DESIGN_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/DESIGN_ExcelScriptLoader-v1.md) | 架构设计文档 |
| [TASK_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/TASK_ExcelScriptLoader-v1.md) | 任务拆分与依赖图 |
| [ACCEPTANCE_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/ACCEPTANCE_ExcelScriptLoader-v1.md) | 验收报告 |
| [FINAL_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/FINAL_ExcelScriptLoader-v1.md) | 实现总结 |
| [TODO_ExcelScriptLoader-v1.md](docs/ExcelScriptLoader-v1/TODO_ExcelScriptLoader-v1.md) | 待办事项与后续建议 |

---

## 👥 社群

**QQ 群：955830545** 提供技术交流与支持，欢迎加入。

> 因为个人是社畜，所以可能不会及时回复，请谅解。

---

## 📄 许可证

本项目基于 [MIT 许可证](LICENSE) 开源。
