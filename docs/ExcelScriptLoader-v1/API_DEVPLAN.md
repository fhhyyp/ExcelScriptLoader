# Excel API 开发计划 v1.1

## 当前状态 (v1.0)

基于 `ObjectValue` + `FunctionValue` 模式，已实现 6 类对象共 30+ API。

---

## 阶段 A：补充核心 API（P0）

### A1. Cell 增强
| API | 说明 | 优先级 |
|-----|------|--------|
| `cell.address()` | 返回 "$A$1" 格式地址 | P0 |
| `cell.row()` / `cell.column()` | 行号/列号 | P0 |
| `cell.offset(r,c)` | 偏移单元格 | P0 |
| `cell.clear()` | 清除单个单元格 | P1 |

### A2. Range 增强
| API | 说明 | 优先级 |
|-----|------|--------|
| `range.find(what)` / `findAll(what)` | 查找 | P0 |
| `range.replace(what, repl)` | 替换 | P1 |
| `range.sort(col, asc?)` | 排序 | P0 |
| `range.autoFilter()` / `range.filterBy(col, val)` / `range.clearFilter()` | 筛选 | P0 |
| `range.currentRegion()` | Ctrl+A 区域 | P1 |
| `range.end(direction)` | Ctrl+方向键 | P1 |
| `range.address()` | 返回地址 | P1 |

### A3. Worksheet 增强
| API | 说明 | 优先级 |
|-----|------|--------|
| `worksheet.activate()` | 切换到该工作表 | P0 |
| `worksheet.delete()` | 删除自身 | P1 |
| `worksheet.count()` | 工作表数量（Workbook 级） | P1 |

---

## 阶段 B：数据查询分析（P1）

### B1. 查找模块
```js
// 建议：excel.find / excel.findAll 作为顶层快捷方法
excel.find("关键词")              // 当前表查找
excel.findAll("关键词")           // 全表查找所有匹配
excel.replace("old", "new")       // 全表替换
```

### B2. 数据导航
```js
let rng = excel.sheet().range("A1")
rng.end(Direction.Down)           // Ctrl+↓
rng.currentRegion()               // Ctrl+A 连续区域
```

### B3. 筛选排序
```js
excel.sheet().range("A1").autoFilter()
excel.sheet().range("A1").filterBy(2, "研发部")
excel.sheet().range("A1:D100").sort(1, false)  // 按第1列降序
```

---

## 阶段 C：格式与样式（P1）

### C1. 增强 format
| API | 说明 |
|-----|------|
| `range.format(fmt)` | 批量格式 |
| `range.autoFit()` | 自动列宽 |
| `range.colWidth(w)` / `range.rowHeight(h)` | 列宽/行高 |
| `range.addBorder()` | 边框 |
| `range.merge()` | 合并单元格 |
| `range.hAlign(align)` | 水平对齐 |
| `cell.fontName(name)` / `cell.fontSize(sz)` | 字体 |

### C2. 颜色常量
```js
import { Color } from "excel"
Color.Red    // 0xFF0000
Color.Blue   // 0x0070C0
Color.Green  // 0x00FF00
```

---

## 阶段 D：事件与交互（P2）

### D1. 用户交互
| API | 说明 |
|-----|------|
| `excel.msgbox(msg)` | 弹消息框 |
| `excel.inputbox(prompt)` | 弹输入框，返回字符串 |
| `excel.confirm(msg)` | 确认框，返回 bool |

### D2. 进度
| API | 说明 |
|-----|------|
| `excel.statusBar(msg)` | 设置状态栏文字 |

---

## 阶段 E：高级功能（P2）

### E1. 图表
```js
let chart = excel.sheet().range("A1:B10").addChart("柱形图")
chart.title("销售数据")
```

### E2. 数据透视表
```js
let pivot = excel.sheet().range("A1:D100").createPivotTable("汇总表")
pivot.rowField("部门")
pivot.dataField("金额", "sum")
```

### E3. 外部文件批量处理
```js
excel.batch(["f1.xlsx","f2.xlsx"], (wb, i) => {
    // 处理每个工作簿
})
```

### E4. 数组工具
```js
import { utils } from "excel"
utils.transpose(arr)      // 转置二维数组
utils.toObjects(arr)      // 二维数组 → 对象数组
utils.toArray(objects)    // 对象数组 → 二维数组
```

---

## 阶段 F：性能与体验（P2）

### F1. 批量操作
| API | 说明 |
|-----|------|
| `excel.suspendEvents()` / `resumeEvents()` | 暂停屏幕刷新 |
| `excel.calculation(mode)` | `"manual"` / `"automatic"` |

### F2. 脚本包管理
| API | 说明 |
|-----|------|
| `excel.include(path)` | 引用外部 .script 文件 |
| `excel.require(name)` | 引用已注册模块 |

---

## 优先级汇总

| 阶段 | 内容 | 新增 API 数 |
|------|------|------------|
| A | 核心补充 | ~10 |
| B | 数据查询分析 | ~8 |
| C | 格式样式 | ~8 |
| D | 事件交互 | ~3 |
| E | 高级功能 | ~5 |
| F | 性能体验 | ~4 |

**建议优先实现 A + B，覆盖 VBA 80% 常用数据操作场景。**
