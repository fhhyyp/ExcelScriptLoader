# Excel 脚本 API 风格与命名约定 v1.0

> 本文档基于 [ExcelModule.cs](../ExcelScriptLoader/ExcelModule.cs) 与 [API_REFERENCE.md](API_REFERENCE.md) 的既有实现，
> 提取 API 设计模式、命名规则、代码骨架，作为后续所有 API 开发的强制性参考。
>
> **原则：新增 API 必须与现有风格一致。如有冲突，以本文档为准。**

---

## 一、脚本侧命名约定（用户可见 API）

### 1.1 大小写

| 规则 | 示例 | 说明 |
|------|------|------|
| 全部 camelCase | `excel.active`, `sheet.read()`, `range.rowCount()` | 无例外。不使用 snake_case 或 PascalCase |
| 模块名小写 | `import { excel } from "excel"` | 模块名始终全小写 |
| 全局函数名小写 | `msgbox()`, `inputbox()` | 全局函数同样 camelCase |

### 1.2 方法命名分类

#### A. 访问器 — 获取子对象

模式：**名词**，表示"获取某个东西"

```
excel.sheet(name?)       → 获取工作表
excel.cell()             → 获取当前单元格
excel.selection()        → 获取选中区域
workbook.sheet(name)     → 获取工作表
worksheet.cell(addr)     → 获取单元格
worksheet.range(addr)    → 获取区域
worksheet.table(name)    → 获取表格
worksheet.usedRange()    → 获取已用区域
```

新增 API 示例：
```
worksheet.name()         → 获取命名区域              ← 名词
workbook.names()         → 获取所有命名区域列表       ← 名词复数
range.entireRow()        → 获取整行范围              ← 名词
range.entireColumn()     → 获取整列范围              ← 名词
```

#### B. Getter/Setter — 读写属性

模式：**同名方法**，无参 = 读，有参 = 写

```
cell.value()             → 读值
cell.value(v)            → 写值
cell.formula()           → 读公式
cell.formula(f)          → 写公式
cell.format()            → 读格式
cell.format(fmt)         → 写格式
range.values()           → 读全部值
range.values(arr)        → 写全部值
```

新增 API 示例：
```
excel.screenUpdating()   → 读屏幕刷新状态
excel.screenUpdating(b)  → 写屏幕刷新状态
worksheet.visible()      → 读可见性
worksheet.visible(v)     → 写可见性
excel.calculation()      → 读计算模式
excel.calculation(mode)  → 写计算模式
```

**选名规则**：
- 如果语义上是一个"属性"（状态、开关、配置），用 getter/setter 同名方法
- 不要因"有副作用"而改用动词 — `screenUpdating` 而非 `setScreenUpdating`
- 参数可选，方法名不带 `get`/`set` 前缀

#### C. 动作 — 执行操作（无返回值 / 返回 void）

模式：**动词**，表示"做某事"

```
workbook.save()          → 保存
workbook.close()         → 关闭
range.clear()            → 清空
table.clear()            → 清空数据
```

新增 API 示例：
```
worksheet.activate()     → 激活工作表                ← 动词
worksheet.select()       → 选择工作表                ← 动词
range.merge()            → 合并单元格                ← 动词
range.unmerge()          → 取消合并                  ← 动词 un+动词
excel.calculate()        → 重新计算                  ← 动词
```

#### D. 动作+宾语 — 创建/删除/操作指定对象

模式：**动词 + 名词**（camelCase 中动词和名词连写，名词首字母大写）

```
workbook.addSheet(name)     → 新建工作表     (add + Sheet)
workbook.removeSheet(name)  → 删除工作表     (remove + Sheet)
table.addAll(rows)          → 批量新增行     (add + All)
```

新增 API 示例：
```
workbook.addName(name, ref)     → 添加命名区域    (add + Name)
worksheet.addChart(...)         → 添加图表        (add + Chart)
worksheet.addPicture(...)       → 添加图片        (add + Picture)
range.addComment(text)          → 添加批注        (add + Comment)
range.addHyperlink(url)         → 添加超链接      (add + Hyperlink)
range.addValidation(type, opt)  → 添加数据验证    (add + Validation)
range.addConditionalFormat(...) → 添加条件格式    (add + ConditionalFormat)
```

**能愿动词命名**：
```
不推荐: removeSheet, deleteSheet → 已有 removeSheet 先例，沿用 remove
不推荐: createSheet, makeSheet  → 已有 addSheet 先例，沿用 add
```

#### E. 查询 — 返回数值/字符串/布尔值

模式：**名词 + Count / 名词复数 / 名词**

```
range.rowCount()         → 行数         (名词 + Count)
range.colCount()         → 列数         (名词 + Count)
range.rows()             → 所有行       (名词复数)
workbook.sheets()        → 所有工作表   (名词复数)
table.rows()             → 所有数据行   (名词复数)
```

新增 API 示例：
```
range.column()           → 列号         (名词)
range.row()              → 行号         (名词)
range.address()          → 地址字符串   (名词)
worksheet.index()        → 索引号       (名词)
workbook.path()          → 文件路径     (名词)
workbook.fullName()      → 完整路径     (名词)
```

#### F. 导航 — 返回同级对象

模式：**形容词/方位词**

```
worksheet.previous()     → 上一个工作表 (方位词)
worksheet.next()         → 下一个工作表 (方位词)
range.offset(r, c)       → 偏移区域     (方位词)
```

#### G. 数据操作 — 有副作用的查询

模式：**动词 + 可选 options 对象**

```
range.find(what, opts?)  → 查找         (动词)
range.sort(key, opts?)   → 排序         (动词)
range.autoFilter(f?,c?)  → 筛选         (verb+名词)
```

### 1.3 参数命名

| 参数 | 命名 | 类型 | 示例 |
|------|------|------|------|
| 地址/区域字符串 | 位置参数，不命名 | String | `"A1"`, `"A1:C10"`, `"1:1"` |
| 工作表名 | 位置参数 | String | `"销售数据"` |
| 文件路径 | 位置参数 | String | `"d:/data.xlsx"` |
| 格式配置 | 位置参数 | Object | `{ bold: true, color: "red" }` |
| 查找选项 | `options` 对象 | Object | `{ lookAt: "whole", matchCase: true }` |
| 排序选项 | `options` 对象 | Object | `{ order: "desc", header: true }` |
| 行偏移 | `rowOffset` 或位置 | Number | `1`, `-1` |
| 列偏移 | `colOffset` 或位置 | Number | `2`, `0` |

**原则**：
- 必选参数用位置参数（不依赖对象 key）
- 可选参数超过 2 个时，打包为一个 options 对象
- options 对象的字段名也使用 camelCase
- 不使用 VBA 风格的命名参数（`LookAt:=xlWhole`）

### 1.4 反例 — 不符合规范的命名

```
❌ excel.ActiveWorkbook       → 应为 excel.active
❌ excel.getSheet("name")     → 应为 excel.sheet("name")
❌ excel.setScreenUpdating()  → 应为 excel.screenUpdating(bool)
❌ range.RowCount             → 应为 range.rowCount()（带括号的方法）
❌ worksheet.AddSheet         → 应为 worksheet.addSheet()
❌ worksheet.delete_sheet     → 应为 worksheet.removeSheet()
```

---

## 二、C# 侧实现约定

### 2.1 注册模式选择

| 场景 | 使用模式 | 参考代码 |
|------|----------|----------|
| `excel` 顶层模块的方法/属性 | **PrototypeExtension** | `ExcelModule` class, 行19-127 |
| Workbook / Worksheet / Cell / Range / Table 的子成员 | **ObjectValue + FunctionValue** | `WorkbookToObject()` 等工厂, 行131-394 |
| 全局函数（msgbox 等） | **GlobalScope.Define()** | `ScriptEngineAdapter.cs` 行313-332 |

**决策树**：
```
新增 API 属于哪个层级？
  ├─ excel 模块本身 (excel.xxx)
  │   └─ 在 ExcelModule class 中添加 [PrototypeProperty] 或 [PrototypeFunction]
  ├─ workbook / worksheet / cell / range / table 的子成员
  │   └─ 在对应的 XxxToObject() 工厂 Dictionary 中添加条目
  └─ 全局独立函数
      └─ 在 ScriptEngineAdapter.RegisterBuiltinFunctions() + ReRegisterCustomFunctions()
```

### 2.2 ObjectValue Dictionary 条目编写模板

#### 只读属性（静态值，构造时即确定）

```csharp
["name"] = StringValue.Create(w.Name),
["path"] = StringValue.Create(w.Path ?? ""),
```

适用：字符串属性、不会变化的值。

#### 只读属性（需要 COM 调用）

```csharp
["rowCount"] = F("rowCount", _ =>
    NumberValueFactory.Create(r.Rows.Count)),
["column"] = F("column", _ =>
    { try { return NumberValueFactory.Create(r.Column); } catch { return Value.Null; } }),
```

适用：数值属性、需要实时从 COM 对象读取的值。

#### 无参动作（void）

```csharp
["save"] = F("save", _ => { try { w.Save(); } catch { } }),
["clear"] = F("clear", _ => { try { r.Clear(); } catch { } }),
```

适用：无参数的纯操作。

#### 单参数动作（string）

```csharp
["saveAs"] = F("saveAs", args =>
{
    if (args.Count > 0 && args[0] is StringValue p)
        try { w.SaveAs(p.Value); } catch { }
}),
```

#### 单参数动作（object / options）

```csharp
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
```

#### Getter/Setter 同名方法（无参=读，有参=写）

```csharp
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
```

#### 返回子对象

```csharp
["offset"] = F("offset", args =>
{
    if (args.Count >= 2
        && args[0] is NumberValue<double> ro
        && args[1] is NumberValue<double> co)
    {
        try { return RangeToObject(r.Offset[(int)ro.Value, (int)co.Value]); }
        catch { }
    }
    return Value.Null;
}),
```

### 2.3 参数类型匹配速查

| 脚本侧类型 | C# 类型匹配 | 取值方式 |
|-----------|------------|---------|
| 字符串 | `StringValue sv` | `sv.Value` |
| 整数 | `NumberValue<int> ni` | `ni.Value` |
| 浮点数 | `NumberValue<double> nd` | `nd.Value`（转为 int: `(int)nd.Value`） |
| 布尔 | `BoolValue bv` | `bv.Value` |
| 对象 | `ObjectValue obj` | `obj.Properties` (Dictionary) |
| 数组 | `ArrayValue av` | `av.Elements` (List\<Value\>) |
| null | `NullValue` 或 `Value.Null` | — |

**检查顺序**：先判断 `args.Count`，再按类型从具体到宽泛匹配。

### 2.4 返回值类型映射

| 脚本侧类型 | C# 构造方式 |
|-----------|-----------|
| String | `StringValue.Create(value)` |
| Number (int) | `NumberValueFactory.Create(intValue)` |
| Number (double) | `NumberValueFactory.Create(doubleValue)` |
| Boolean | `BoolValue.Create(boolValue)` |
| null | `Value.Null` |
| Array | `new ArrayValue(List<Value>)` |
| Object | `new ObjectValue(Dictionary<string, Value>)` |
| 子对象 (Workbook等) | `WorkbookToObject(wb)` / `RangeToObject(rng)` / ... |

### 2.5 错误处理

**铁律：每个 COM 调用必须包裹在 try/catch 中。**

```csharp
// 正确 ✅
["save"] = F("save", _ => { try { w.Save(); } catch { } }),

// 错误 ❌ — COM 调用未捕获
["save"] = F("save", _ => { w.Save(); }),
```

**On failure:**
- 返回值的函数 → `return Value.Null`
- void 函数 → 静默吞下异常
- 永远不要向脚本层抛出 C# 异常

### 2.6 闭包捕获

```csharp
// 正确 ✅ — 先捕获再使用
internal static ObjectValue WorkbookToObject(Workbook wb)
{
    var w = wb;  // capture
    return new ObjectValue(new()
    {
        ["save"] = F("save", _ => { try { w.Save(); } catch { } }),
    });
}

// 错误 ❌ — 直接使用参数可能导致闭包引用问题
internal static ObjectValue WorkbookToObject(Workbook wb)
{
    return new ObjectValue(new()
    {
        ["save"] = F("save", _ => { try { wb.Save(); } catch { } }),
    });
}
```

### 2.7 Factory 方法命名

| 模式 | 示例 |
|------|------|
| `{类型}ToObject` | `WorkbookToObject`, `WorksheetToObject`, `CellToObject`, `RangeToObject`, `TableToObject` |
| 新增类型遵循此模式 | `ChartToObject`, `PivotTableToObject`, `NameToObject` |

### 2.8 工具方法

| 方法 | 用途 | 签名 |
|------|------|------|
| `F(name, Func<...>)` | 创建有返回值的 FunctionValue | `FunctionValue F(string, Func<List<Value>, Value>)` |
| `F(name, Action<...>)` | 创建无返回值的 FunctionValue | `FunctionValue F(string, Action<List<Value>>)` |
| `WrapCell(object?)` | COM 值 → ScriptLang Value | `Value WrapCell(object? v)` |
| `UnwrapV(Value)` | ScriptLang Value → COM object | `object UnwrapV(Value v)` |
| `ReadRange(ExcelRange)` | ExcelRange → ArrayValue (二维) | `Value ReadRange(ExcelRange rng)` |
| `WriteRange(ExcelRange, Value)` | ArrayValue → Excel COM 写入 | `void WriteRange(ExcelRange, Value)` |
| `WriteObjects(ExcelRange, Value)` | 对象数组 → Excel COM 写入（含表头） | `void WriteObjects(ExcelRange, Value)` |
| `ApplyFormat(ExcelRange, ObjectValue)` | 对象配置 → 单元格格式 | `void ApplyFormat(ExcelRange, ObjectValue)` |
| `ParseColor(Value)` | string/int Value → COM Color int | `object ParseColor(Value c)` |

---

## 三、全局函数约定

### 3.1 注册位置

```csharp
// ScriptEngineAdapter.cs — RegisterBuiltinFunctions()
_engine.GlobalScope.Define("funcName", new ClrObjectValue(new Func<...>(...)), isMutable: false);

// 同时在 ReRegisterCustomFunctions() 中添加恢复逻辑
if (!_engine.GlobalScope.IsDefinedLocally("funcName"))
{
    _engine.GlobalScope.Define("funcName", new ClrObjectValue(new Func<...>(...)), isMutable: false);
}
```

### 3.2 命名

- 全部小写 camelCase
- 不与 ScriptLang 内置函数（`print`, `len`, `range`, `keys` 等）冲突
- 不与 `excel` 模块导出名冲突

---

## 四、模块注册约定

### 4.1 模块导出名

```csharp
// ScriptEngineAdapter.cs — RegisterExcelModule()
var properties = new Dictionary<string, Value>
{
    { "excel", new ClrObjectValue(excel) },  // 导出名 = 模块名
};
var module = new ObjectValue(properties);
_engine.ImportResolver.RegisterBuiltinModule("excel", module);
```

- 模块名全小写
- export 对象 key 与模块名一致
- 如果未来新增模块（如 `chart`），遵循同样规则

### 4.2 Prototype 注册

```csharp
// ScriptEngineAdapter.cs — RegisterWrapperPrototypes()
var proto = (IPrototype)new ExcelModule(_excelApp!);
_engine?.PrototypeManager.Register(proto);
```

- 仅对使用 `[PrototypeExtension]` 的类注册原型
- ObjectValue+FunctionValue 模式的对象无需注册原型

---

## 五、选项对象（Options Object）约定

当方法需要 3 个以上可选参数时，打包为一个 options 对象。

### 5.1 字段命名

全部 camelCase，与顶级 API 命名一致：

```
{ lookAt: "whole", matchCase: true, searchOrder: "rows" }       ← 查找选项
{ order: "desc", header: true }                                  ← 排序选项
{ bold: true, color: "red", bgColor: 0x4472C4, size: 14 }      ← 格式选项
```

### 5.2 枚举值

使用小写字符串代替数字常量：

```
"whole" / "part"              ← 匹配模式
"asc" / "desc"                ← 排序方向
"rows" / "cols"               ← 搜索/偏移方向
"hidden" / "veryHidden"       ← 可见性
"auto" / "manual" / "semiauto" ← 计算模式
"red" / "green" / "blue" ...  ← 颜色名
```

### 5.3 C# 侧解析

```csharp
if (args.Count > 1 && args[1] is ObjectValue opts)
{
    if (opts.Properties.TryGetValue("fieldName", out var val) && val is ExpectedType tv)
    {
        // 使用 tv.Value
    }
}
```

---

## 六、文件组织

### 6.1 ExcelModule.cs 内部结构

```
ExcelModule class
├── [PrototypeProperty] 属性              ← excel.xxx 顶层属性
├── [PrototypeFunction] 方法              ← excel.xxx 顶层方法
├── internal static WorkbookToObject()    ← 工厂方法（按字母排列）
├── internal static WorksheetToObject()
├── internal static CellToObject()
├── internal static RangeToObject()
├── internal static TableToObject()
├── (新增) internal static ChartToObject()    ← 新类型工厂
├── internal static F()                   ← 工具方法
├── internal static WrapCell()
├── internal static UnwrapV()
├── internal static ReadRange()
├── internal static WriteRange()
├── internal static WriteObjects()
├── internal static ApplyFormat()
└── internal static ParseColor()
```

### 6.2 新增类型工厂的位置

在现有工厂方法后面、工具方法前面插入新工厂：

```csharp
    // 在 TableToObject() 之后
    internal static ObjectValue ChartToObject(Chart chart) { ... }
    internal static ObjectValue PivotToObject(PivotTable pt) { ... }

    // ---- 内部工具 ----  (保持不变)
    internal static FunctionValue F(...)
```

---

## 七、API 新增检查清单

新增任何 API 前，逐项确认：

- [ ] 脚本侧方法名使用 camelCase
- [ ] 方法命名符合分类（访问器 / getter-setter / 动作 / 动作+宾语 / 查询 / 导航）
- [ ] 可选参数 ≤2 个时用位置参数，>2 个时打包为 options 对象
- [ ] options 对象的字段名使用 camelCase
- [ ] 枚举值使用小写字符串而非数字
- [ ] 每个 COM 调用包裹在 try/catch 中
- [ ] 失败时返回 `Value.Null`（而非抛异常）
- [ ] 使用 `var x = obj;` 闭包捕获模式
- [ ] 类型匹配使用 `is StringValue sv` 模式取值的标准写法
- [ ] 返回值使用正确工厂方法（`StringValue.Create` / `NumberValueFactory.Create` / `BoolValue.Create` / `Value.Null`）
- [ ] 已添加对应的 `ReRegisterCustomFunctions()` 逻辑（如果是全局函数）
- [ ] 命名不与已有 API 冲突
- [ ] 保持与周围代码相同的缩进和空行风格

---

## 八、快速参考卡片

### 命名速查

```
场景                          模式                  示例
──────────────────────────────────────────────────────────
获取子对象                     名词                  .sheet("name"), .cell("A1")
可读写属性                     getter/setter同名     .value()/.value(v), .visible()/.visible(v)
只读属性                       F() 返回静态值        .rowCount(), .address()
无参操作                      动词                  .save(), .clear(), .activate()
创建对象                       add+名词             .addSheet(name), .addChart(...)
删除对象                       remove+名词          .removeSheet(name)
批量操作                       addAll / 名词复数    .addAll(rows), .sheets()
查找/搜索                      动词+options         .find(what, opts?)
相对移动                       方位词                .offset(r,c), .next(), .previous()
状态开关                       getter/setter同名     .screenUpdating()/.screenUpdating(b)
选项配置                       对象字面量            { field: value, ... }
```

### 代码骨架速查

```
只读静态属性:  ["name"] = StringValue.Create(w.Name),
只读COM属性:   ["rowCount"] = F("rowCount", _ => NumberValueFactory.Create(r.Rows.Count)),
无参动作:      ["save"] = F("save", _ => { try { w.Save(); } catch { } }),
单参动作:      ["saveAs"] = F("saveAs", args => { if (args.Count > 0 && args[0] is StringValue p) try { ... } catch { } }),
GetterSetter: ["value"] = F("value", args => { if (args.Count == 0) { /*get*/ } else { /*set*/ } }),
返回子对象:    ["offset"] = F("offset", args => { ... try { return RangeToObject(r.Offset[...]); } catch { } return Value.Null; }),
对象参数:      ["add"] = F("add", args => { if (args.Count > 0 && args[0] is ObjectValue obj) try { ... } catch { } }),
全局函数:      _engine.GlobalScope.Define("name", new ClrObjectValue(new Func<...>(...)), isMutable: false);
```
