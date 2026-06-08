# TODO: ExcelScriptLoader v1 — 待办事项与后续建议

## 必须完成（P0）

### 1. Excel 环境实测
- [ ] 在安装了 Excel 2019+ 的 Windows 机器上加载 XLL
- [ ] 验证 .NET 10 运行时在 Excel 进程内加载是否正常
- [ ] 测试新建宏 → 保存 → 关闭 → 重新打开 → 宏仍存在
- [ ] 测试脚本中访问 app/workbook/sheet/cell/selection

### 2. Excel-DNA 兼容性
- [ ] 确认 RuntimeVersion="v10.0" 是否正确（可能需要 "v6.0" + rollForward）
- [ ] 如 .NET 10 不被 Excel-DNA 原生支持，调整为 .NET 8 作为 TargetFramework
- [ ] 验证 64 位 Excel 加载 64 位 XLL 正常

### 3. office.dll 路径
- [ ] 当前 office.dll HintPath 硬编码为 Office 15.0 GAC
- [ ] Office 2016/365 可能使用不同版本路径
- [ ] 建议：使用 NuGet 包 `Microsoft.Office.Core` 替代 GAC 引用

## 建议优化（P1）

### 4. 线程安全改进
- [ ] 当前 `Execute()` 使用 `Task.Run().GetAwaiter().GetResult()` 在 UI 线程同步等待
- [ ] 考虑改为 async void 事件处理 + loading 指示器

### 5. 编译缓存优化
- [ ] 当前每次执行前 `ClearCache()`，浪费了编译缓存
- [ ] 实现基于代码 Hash 的缓存键，代码不变时不重新编译

### 6. 语法高亮编辑器
- [ ] 使用 ScintillaNET 或 RichTextBox 着色实现代码语法高亮
- [ ] 至少实现关键字着色和行号显示

### 7. 错误行号定位
- [ ] 解析脚本异常时提取行列信息
- [ ] 在编辑器中高亮错误行

## 后续功能（P2）

### 8. 宏导入/导出
- [ ] 导出宏为 .script 文件
- [ ] 从 .script 文件导入宏
- [ ] 批量导入/导出

### 9. 快捷键绑定
- [ ] 快捷键配置 UI
- [ ] 全局键盘钩子监听
- [ ] 快捷键冲突检测

### 10. VBA Alt+F8 集成（备选）
- [ ] 研究 VBProject.VBComponents 动态模块注入
- [ ] 实现 VBA 存根代码生成
- [ ] 处理信任中心设置检测和提示

### 11. 脚本预编译
- [ ] 宏保存时预编译为 ByteCodeChunk (.ssc)
- [ ] 运行时直接加载 .ssc，跳过编译阶段
- [ ] 显著加快重复执行速度

### 12. NuGet 包发布
- [ ] 将 SereinScript 发布为 NuGet 包
- [ ] ExcelScriptLoader 引用 NuGet 包替代项目引用
- [ ] 便于独立打包和版本管理

## 已知限制

- **单线程执行**：当前脚本在后台 Task 执行但同步等待，大脚本会阻塞 UI
- **无沙箱**：脚本可访问所有 CLR 对象（包括文件系统、网络），存在安全风险
- **Office 版本**：未在 Office 2016 以下版本测试
- **32 位 Excel**：当前仅配置 x64，32 位需额外配置
- **无自动完成**：代码编辑器不支持智能提示
