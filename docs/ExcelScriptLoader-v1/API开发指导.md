# Excel 标准库设计规范

## 1. 设计目标

Excel 模块用于提供：

- 当前宿主 Excel 的访问能力
- 外部 Excel 文件的读写能力
- 工作表操作能力
- 单元格与区域操作能力
- 表格（Table）操作能力
- 对象集合与二维数组之间的转换能力

设计原则：

1. 简洁优先
2. 数据优先
3. 尽量减少对象层级
4. 默认处理常见场景
5. 与 JavaScript 使用习惯保持一致
6. 宿主 Excel 与外部 Excel 保持统一模型

------

# 2. 导入

```js
import { excel } from "excel"
```

------

# 3. 顶层 API

## 当前工作簿

```js
excel.active
```

返回当前宿主工作簿。

示例：

```js
book = excel.active
```

------

## 获取工作表

获取当前活动工作表：

```js
excel.sheet()
```

等价于：

```js
excel.active.activeSheet
```

获取指定工作表：

```js
excel.sheet("销售数据")
```

------

## 当前单元格

```js
excel.cell()
```

返回当前活动单元格。

------

## 当前选区

```js
excel.selection()
```

返回当前选中区域。

------

## 打开文件

```js
excel.open(path)
```

示例：

```js
book = excel.open("d:/sales.xlsx")
```

------

## 创建工作簿

```js
excel.create()
```

示例：

```js
book = excel.create()
```

------

## 快捷读取

读取整个工作表：

```js
excel.read(path)
```

读取指定工作表：

```js
excel.read(path, sheetName)
```

示例：

```js
data = excel.read("sales.xlsx")
```

返回：

```js
[
    ["姓名", "金额"],
    ["Tom", 100]
]
```

------

## 快捷写入

```js
excel.write(path, data)
```

示例：

```js
excel.write("result.xlsx", data)
```

------

# 4. Workbook

工作簿对象。

------

## 获取工作表

```js
book.sheet(name)
```

示例：

```js
sheet = book.sheet("销售")
```

------

## 创建工作表

```js
book.addSheet(name)
```

示例：

```js
sheet = book.addSheet("统计")
```

------

## 删除工作表

```js
book.removeSheet(name)
```

------

## 工作表集合

```js
book.sheets
```

示例：

```js
for sheet in book.sheets {
    println(sheet.name)
}
```

------

## 保存

```js
book.save()
```

------

## 另存为

```js
book.saveAs(path)
```

------

## 关闭

```js
book.close()
```

------

# 5. Worksheet

工作表对象。

------

## 单元格

```js
sheet.cell(address)
```

示例：

```js
sheet.cell("A1")
```

------

## 区域

```js
sheet.range(address)
```

示例：

```js
sheet.range("A1:C10")
```

------

## 获取表格

```js
sheet.table(name)
```

示例：

```js
table = sheet.table("销售表")
```

------

## 已使用区域

```js
sheet.usedRange()
```

示例：

```js
data = sheet.usedRange().values
```

------

## 读取数据

读取整个已使用区域：

```js
sheet.read()
```

读取指定区域：

```js
sheet.read("A1:C10")
```

示例：

```js
data = sheet.read()
```

------

## 写入数据

```js
sheet.write(address, data)
```

示例：

```js
sheet.write("A1", data)
```

自动扩展区域大小。

------

## 写入对象集合

```js
sheet.writeObjects(address, objects)
```

示例：

```js
sheet.writeObjects("A1", users)
```

自动生成表头。

------

# 6. Cell

单元格对象。

------

## 值

读取：

```js
cell.value
```

写入：

```js
cell.value = "Tom"
```

------

## 公式

```js
cell.formula
```

示例：

```js
cell.formula = "=SUM(A:A)"
```

------

## 格式

```js
cell.format
```

示例：

```js
cell.format = {
    bold = true,
    color = "red"
}
```

------

# 7. Range

区域对象。

------

## 值集合

```js
range.values
```

返回二维数组：

```js
[
    ["姓名", "年龄"],
    ["Tom", 18]
]
```

------

## 写入

```js
range.values = data
```

------

## 行集合

```js
range.rows
```

------

## 列集合

```js
range.columns
```

------

## 清空

```js
range.clear()
```

------

## 复制

```js
range.copy(target)
```

示例：

```js
sheet.range("A1:C10")
    .copy("E1")
```

------

# 8. Table

Excel 表格对象。

------

## 获取所有记录

```js
table.rows
```

返回：

```js
[
    {姓名="Tom", 金额=100},
    {姓名="Lucy", 金额=200}
]
```

------

## 新增记录

```js
table.add(row)
```

示例：

```js
table.add({
    姓名="Jack",
    金额=300
})
```

------

## 批量新增

```js
table.addAll(rows)
```

------

## 清空数据

```js
table.clear()
```

------

## 转换为对象集合

```js
table.objects()
```

等价于：

```js
table.rows
```

------

## 从对象集合加载

```js
table.objects(data)
```

示例：

```js
table.objects(users)
```

------

# 9. 推荐数据模型

推荐优先使用对象集合：

```js
[
    {
        姓名 = "Tom",
        年龄 = 18
    },
    {
        姓名 = "Lucy",
        年龄 = 20
    }
]
```

而非二维数组：

```js
[
    ["姓名", "年龄"],
    ["Tom", 18],
    ["Lucy", 20]
]
```

原因：

- 可读性更高
- 更适合过滤与转换
- 更符合脚本语言风格

------

# 10. 典型示例

## 当前工作表读取

```js
data = excel.sheet().read()
```

------

## 当前选区读取

```js
data = excel.selection().values
```

------

## 读取外部文件

```js
data = excel.read("sales.xlsx")
```

------

## 写入外部文件

```js
excel.write("result.xlsx", data)
```

------

## 根据对象生成报表

```js
users = [
    {
        姓名 = "Tom",
        年龄 = 18
    },
    {
        姓名 = "Lucy",
        年龄 = 20
    }
]

excel.sheet("结果")
     .writeObjects("A1", users)
```

------

## 表格过滤

```js
rows = excel
    .sheet("用户")
    .table("用户表")
    .rows
    .where(x => x.年龄 >= 18)
```

------

# 11. 核心 API 一览

```js
excel.sheet(name?)
excel.cell()
excel.selection()

excel.open(path)
excel.create()

excel.read(path, sheet?)
excel.write(path, data)

book.sheet(name)
book.addSheet(name)

book.save()
book.saveAs(path)
book.close()

sheet.cell(address)
sheet.range(address)
sheet.table(name)

sheet.read(address?)
sheet.write(address, data)
sheet.writeObjects(address, objects)

sheet.usedRange()

cell.value
cell.formula
cell.format

range.values
range.clear()
range.copy(target)

table.rows
table.add(row)
table.addAll(rows)
table.clear()
```

该集合应作为 Excel 标准库的最小可用 API 集。