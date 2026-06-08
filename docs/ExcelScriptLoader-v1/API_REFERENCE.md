# Excel 脚本 API 参考 v1.0

## 导入

```js
import { excel } from "excel"
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
```js
let data = excel.read("d:/data.xlsx")
let data2 = excel.read("d:/data.xlsx", "报表")
```

#### `excel.write(path, data)` → void
快捷写入：打开→写入→保存→关闭。
```js
excel.write("d:/result.xlsx", [["姓名","年龄"],["Tom",18]])
```

---

## 二、Workbook（工作簿）

`excel.active` / `excel.open()` / `excel.create()` 返回。

| 成员 | 类型 | 参数 | 说明 |
|------|------|------|------|
| `.name` | String | — | 文件名（属性） |
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
| `.cell(addr)` | Cell | "A1" | 获取单元格 |
| `.range(addr)` | Range | "A1:C10" | 获取区域 |
| `.table(name)` | Table | "表名" | 获取 Excel 表格 |
| `.usedRange()` | Range | — | 已用区域 |
| `.read(addr?)` | Array | —/"A1:B20" | 读取数据 |
| `.write(addr, data)` | void | addr, arr | 写入二维数组 |
| `.writeObjects(addr, data)` | void | addr, arr | 写入对象数组 |

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
| `.rows()` | Array | — | 所有行 |
| `.rowCount()` | Number | — | 行数 |
| `.colCount()` | Number | — | 列数 |
| `.clear()` | void | — | 清空内容和格式 |
| `.copy(target)` | void | addr | 复制到目标 |

```js
let rng = excel.sheet().range("A1:C5")
let data = rng.values()           // 读取 → 二维数组
rng.values(newData)               // 写入
print(rng.rowCount(), rng.colCount())
rng.copy("E1")                    // 复制
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
excel.active          → Workbook
excel.sheet(name?)    → Worksheet
excel.cell()          → Cell
excel.selection()     → Range
excel.open(path)      → Workbook
excel.create()        → Workbook
excel.read(path,sh?)  → Array    快捷读
excel.write(path,arr) → void     快捷写

workbook.name         → String
workbook.save()       .saveAs() .close()
workbook.sheet(name)  .addSheet(name) .removeSheet(name) .sheets()

worksheet.name        → String
worksheet.cell("A1")  → Cell
worksheet.range("A:C")→ Range
worksheet.table("名")  → Table
worksheet.usedRange() → Range
worksheet.read(a?)    .write(a,arr) .writeObjects(a,arr)

cell.value()/.value(v)  .formula()/.formula(f)  .format()/.format(f)

range.values()/.values(arr)  .rows() .rowCount() .colCount() .clear() .copy(to)

table.name .rows() .add(row) .addAll(rows) .clear()
```
