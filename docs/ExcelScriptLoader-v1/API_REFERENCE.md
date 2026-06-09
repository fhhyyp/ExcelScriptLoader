# Excel 脚本 API 参考 v1.0

## 导入

```js
import { excel, msgbox, inputbox } from "excel"
```

`msgbox` / `inputbox` 与 `excel` 同级从模块导入，调用时无需 `excel.` 前缀：
```js
msgbox("处理完成！")
let name = inputbox("请输入姓名：")
```

---

## 一、excel（顶层入口）

### 属性

| 成员 | 类型 | 说明 |
|------|------|------|
| `excel.active` | Workbook | 当前宿主工作簿 |

### 方法

#### `excel.sheet(name?)` → Worksheet
获取工作表。无参=活动工作表，有参=按名称。
```js
let s = excel.sheet()             // 活动工作表
let s2 = excel.sheet("销售数据")  // 按名称
```

#### `excel.cell()` → Cell
当前活动单元格。
```js
let c = excel.cell()
c.value("Hello")
```

#### `excel.selection()` → Range
当前选中区域。
```js
let rng = excel.selection()
rng.values([["A","B"],[1,2]])
```

#### `excel.open(path)` → Workbook
打开外部工作簿。
```js
let wb = excel.open("d:/sales.xlsx")
let data = wb.sheet("Sheet1").read()
wb.close()
```

#### `excel.create()` → Workbook
新建空白工作簿。

#### `excel.read(path, sheetName?)` → Array
快捷读取：打开→读取→关闭。返回二维数组。
不指定工作表名时默认读取第一个工作表（Sheets[1]），而非活动工作表。
```js
let data = excel.read("d:/data.xlsx")
let data2 = excel.read("d:/data.xlsx", "报表")
```

#### `excel.write(path, data)` → void
快捷写入：打开文件（不存在则创建）→ 写入第一张工作表 → 保存 → 关闭。
注意：始终写入 Sheets[1]，不支持指定其他工作表。
```js
excel.write("d:/result.xlsx", [["姓名","年龄"],["Tom",18]])
```

#### `excel.screenUpdating()` / `excel.screenUpdating(bool)` → bool / void
获取/设置屏幕刷新。批量写入前关闭可大幅提升性能。
```js
let prev = excel.screenUpdating()
excel.screenUpdating(false)
// ... 批量操作 ...
excel.screenUpdating(prev)
```

#### `excel.displayAlerts()` / `excel.displayAlerts(bool)` → bool / void
获取/设置弹窗警告。关闭后保存提示、删除确认等不会中断脚本。
```js
excel.displayAlerts(false)
excel.active.close()
excel.displayAlerts(true)
```

#### `excel.calculation()` / `excel.calculation(mode)` → str / void
获取/设置计算模式。mode: `"auto"` / `"manual"` / `"semiauto"`。
```js
excel.calculation("manual")
// ... 大批量写入 ...
excel.calculate()
excel.calculation("auto")
```

#### `excel.calculate()` → void
强制重新计算所有打开的工作簿。
```js
excel.calculate()
```

#### `excel.enableEvents()` / `excel.enableEvents(bool)` → bool / void
获取/设置事件触发开关。
```js
excel.enableEvents(false)
// ... 脚本操作不触发 Excel 内置事件 ...
excel.enableEvents(true)
```

#### `excel.copyText(text)` → void
复制文本到系统剪贴板。
```js
excel.copyText("Hello World")
```

#### `excel.pasteText()` → String
从系统剪贴板粘贴文本。
```js
let text = excel.pasteText()
excel.sheet().cell("A1").value(text)
```

---

## 二、Workbook（工作簿）

`excel.active` / `excel.open()` / `excel.create()` 返回。

| 成员 | 类型 | 参数 | 说明 |
|------|------|------|------|
| `.name` | String | — | 文件名（属性） |
| `.path()` | String | — | 文件所在目录路径 |
| `.fullName()` | String | — | 完整路径（含文件名） |
| `.save()` | void | — | 保存 |
| `.saveAs(path)` | void | path | 另存为 |
| `.close()` | void | — | 关闭 |
| `.sheet(name)` | Worksheet | name | 按名称获取工作表 |
| `.addSheet(name)` | Worksheet | name | 新建并命名工作表 |
| `.removeSheet(name)` | void | name | 删除工作表 |
| `.sheets()` | Array | — | 所有工作表数组 |

```js
let wb = excel.active
wb.addSheet("汇总")
for s in wb.sheets() { print(s.name) }
```

---

## 三、Worksheet（工作表）

`excel.sheet()` / `workbook.sheet()` 返回。

| 成员 | 类型 | 参数 | 说明 |
|------|------|------|------|
| `.name` | String | — | 工作表名（属性） |
| `.activate()` | void | — | 激活/切换到当前工作表 |
| `.cell(addr)` | Cell | "A1" | 获取单元格 |
| `.range(addr)` | Range | "A1:C10" | 获取区域 |
| `.table(name)` | Table | "表名" | 获取 Excel 表格 |
| `.usedRange()` | Range | — | 已用区域 |
| `.read(addr?)` | Array | —/"A1:B20" | 读取数据 |
| `.write(addr, data)` | void | addr, arr | 写入二维数组 |
| `.writeObjects(addr, data)` | void | addr, arr | 写入对象数组 |
| `.insertRow(index)` | void | 行号 | 在指定行前插入一行 |
| `.removeRow(index)` | void | 行号 | 删除指定行 |
| `.insertColumn(index)` | void | 列号 | 在指定列前插入一列 |
| `.removeColumn(index)` | void | 列号 | 删除指定列 |
| `.hideRow(index)` | void | 行号 | 隐藏指定行 |
| `.hideColumn(index)` | void | 列号 | 隐藏指定列 |
| `.rowHeight()` / `.rowHeight(row)` / `.rowHeight(row, h)` | Num | row?, h? | 读标准行高/指定行高/设置行高 |
| `.columnWidth()` / `.columnWidth(col)` / `.columnWidth(col, w)` | Num | col?, w? | 读标准列宽/指定列宽/设置列宽 |
| `.visible()` / `.visible(mode)` | bool/void | bool\|str | 读可见性/设置 (true/false/"hidden"/"veryHidden") |
| `.showAllData()` | void | — | 清除工作表中所有筛选 |

```js
let ws = excel.sheet("数据")
let all = ws.read()                      // 全部数据
let part = ws.read("A1:B20")             // 指定区域
ws.write("A1", [["名称","值"],["A",1]])
```

---

## 四、Cell（单元格）

`worksheet.cell()` / `excel.cell()` 返回。

所有值通过同名方法 get/set：**无参=读取，有参=写入**。

| 方法 | 示例 |
|------|------|
| `.value()` / `.value(v)` | `cell.value()` `cell.value("Hello")` |
| `.formula()` / `.formula(f)` | `cell.formula()` `cell.formula("=SUM(A:A)")` |
| `.format()` / `.format(fmt)` | `cell.format()` `cell.format({bold=true})` |
| `.address()` | `cell.address()` — 返回如 "$A$1" |
| `.row()` | `cell.row()` — 行号 |
| `.column()` | `cell.column()` — 列号 |
| `.hasFormula()` | `cell.hasFormula()` — 是否为公式 |

### format 支持字段
| 字段 | 类型 | 说明 |
|------|------|------|
| `bold` | bool | 加粗 |
| `color` | str\|int | 字体色 ("red"\|0xFF0000) |
| `bgColor` | str\|int | 背景色 |
| `size` | number | 字号 |
| `numberFormat` | str | 数字格式 |

```js
let c = excel.sheet().cell("A1")
c.value("标题")
c.format({
    bold = true,
    color = "white",
    bgColor = 0x4472C4,
    size = 14,
})
```

---

## 五、Range（区域）

`worksheet.range()` / `worksheet.usedRange()` / `excel.selection()` 返回。

| 成员 | 类型 | 参数 | 说明 |
|------|------|------|------|
| `.values()` / `.values(arr)` | Array | —/arr | 读/写全部值 |
| `.formulas()` / `.formulas(arr)` | Array | —/arr | 读/写全部公式 |
| `.rows()` | Array | — | 所有行 |
| `.rowCount()` | Number | — | 行数 |
| `.colCount()` | Number | — | 列数 |
| `.clear()` | void | — | 清空内容和格式 |
| `.offset(rowOffset, colOffset)` | Range | row, col | 偏移区域 |
| `.find(what, opts?)` | Range\|null | what, {lookAt?, matchCase?} | 查找单元格 |
| `.findNext()` | Range\|null | — | 继续查找下一个 |
| `.copy(target)` | void | addr | 复制到目标 |
| `.cut()` / `.cut(target)` | void | —/addr | 剪切到剪贴板 / 剪切到目标 |
| `.insert(direction?)` | void | "down"\|"right" | 插入单元格 |
| `.delete(direction?)` | void | "up"\|"left" | 删除单元格 |
| `.entireRow()` | Range | — | 整行范围 |
| `.entireColumn()` | Range | — | 整列范围 |
| `.merge()` | void | — | 合并单元格 |
| `.unmerge()` | void | — | 取消合并 |
| `.autoFit()` | void | — | 自动调整行高列宽 |
| `.border()` / `.border(opts)` | obj/void | {style, edges, color?} | 读/写边框 |
| `.wrapText()` / `.wrapText(bool)` | bool/void | bool | 读/写自动换行 |
| `.alignment()` / `.alignment(opts)` | obj/void | {horizontal?, vertical?} | 读/写对齐 |
| `.autoFilter(field?, criteria?)` | void | field?, crit? | 自动筛选 |
| `.sort(key, opts?)` | void | key, {order?, header?} | 排序 |

```js
let rng = excel.sheet().range("A1:C5")
let data = rng.values()           // 读取 → 二维数组
rng.values(newData)               // 写入
print(rng.rowCount(), rng.colCount())
rng.copy("E1")                    // 复制

// 查找
let found = rng.find("关键词", { lookAt: "whole" })
if (found) { print(found.address()) }

// 偏移
let target = rng.offset(1, 2)

// 格式化
rng.merge()
rng.border({ style: "thin", edges: "all", color: "black" })
rng.alignment({ horizontal: "center", vertical: "middle" })

// 筛选排序
rng.autoFilter(1, ">100")
rng.sort("B", { order: "desc" })
```

---

## 六、Table（表格）

`worksheet.table(name)` 返回（Excel ListObject）。

| 成员 | 类型 | 参数 | 说明 |
|------|------|------|------|
| `.name` | String | — | 表名（属性） |
| `.rows()` | Array\<Object\> | — | 所有数据行 |
| `.add(row)` | void | Object | 新增一行 |
| `.addAll(rows)` | void | Array | 批量新增 |
| `.clear()` | void | — | 清空数据 |

```js
let t = excel.sheet("数据").table("用户表")

// 遍历
for row in t.rows() {
    print(row["姓名"] + ": " + row["金额"])
}

// 新增
t.add({ 姓名 = "Jack", 年龄 = 25, 部门 = "研发" })
t.addAll([{ 姓名 = "A" }, { 姓名 = "B" }])
t.clear()
```

---

## 七、完整示例

### 数据汇总
```js
import { excel } from "excel"

let data = excel.sheet("明细").read()
let summary = {}
for i in range(1, len(data)) {
    let dept = data[i][1]
    let amt = data[i][3]
    summary[dept] = (summary[dept] ?? 0) + amt
}

let report = [["部门","合计"]]
for dept in keys(summary) { report.push([dept, summary[dept]]) }

let ws = excel.active.addSheet("汇总")
ws.write("A1", report)
ws.cell("A1").format({ bold = true, color = "blue" })
```

### 条件格式
```js
import { excel } from "excel"

let ws = excel.sheet()
let rng = ws.usedRange()

for i in range(2, rng.rowCount() + 1) {
    let amount = ws.cell("C" + i).value()
    if amount > 1000 {
        ws.cell("C" + i).format({ color = "red", bold = true })
    }
}
ws.range("1:1").format({ bold = true, bgColor = 0x4472C4, color = "white" })
```

### 跨文件合并
```js
import { excel } from "excel"

for file in ["jan.xlsx","feb.xlsx","mar.xlsx"] {
    let data = excel.read(file)
    for i in range(1, len(data)) {
        // 处理...
    }
}
```

---

## 八、速查表

```
excel.active             → Workbook
excel.sheet(name?)       → Worksheet
excel.cell()             → Cell
excel.selection()        → Range
excel.open(path)         → Workbook
excel.create()           → Workbook
excel.read(path,sh?)     → Array    快捷读
excel.write(path,arr)    → void     快捷写
excel.screenUpdating()/()→ bool/void  屏幕刷新
excel.displayAlerts()/() → bool/void  弹窗警告
excel.calculation()/()   → str/void   计算模式
excel.calculate()        → void       强制重算
excel.enableEvents()/()  → bool/void  事件开关
excel.copyText(text)     → void       复制到剪贴板
excel.pasteText()        → String     从剪贴板粘贴

msgbox(message)          → void       弹出消息框（模块导入）
inputbox(prompt)         → String     弹出输入框（模块导入）

workbook.name .path() .fullName()
workbook.save() .saveAs() .close()
workbook.sheet(name) .addSheet(name) .removeSheet(name) .sheets()

worksheet.name .activate()
worksheet.cell("A1") .range("A:C") .table("名") .usedRange()
worksheet.read(a?) .write(a,arr) .writeObjects(a,arr)
worksheet.insertRow(i) .removeRow(i) .insertColumn(i) .removeColumn(i)
worksheet.hideRow(i) .hideColumn(i)
worksheet.rowHeight()/.rowHeight(r)/.rowHeight(r,h)
worksheet.columnWidth()/.columnWidth(c)/.columnWidth(c,w)
worksheet.visible()/.visible(v) .showAllData()

cell.value()/.value(v) .formula()/.formula(f) .format()/.format(f)
cell.address() .row() .column() .hasFormula()

range.values()/.values(arr) .formulas()/.formulas(arr)
range.rows() .rowCount() .colCount() .clear()
range.offset(r,c) .find(what,opts?) .findNext()
range.copy(to) .cut()/.cut(to)
range.insert(dir?) .delete(dir?) .entireRow() .entireColumn()
range.merge() .unmerge() .autoFit()
range.border()/.border(opts) .wrapText()/.wrapText(b) .alignment()/.alignment(opts)
range.autoFilter(f?,c?) .sort(key,opts?)

table.name .rows() .add(row) .addAll(rows) .clear()
```

---

