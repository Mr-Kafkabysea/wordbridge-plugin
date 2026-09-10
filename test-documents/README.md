# 准备 Word 测试文档

本目录用于真实 Word 集成验证；`tests/` 保存自动化测试代码。仓库不分发个人测试文档，首次使用需要自行准备。

1. 在桌面 Word 中新建文档，输入几行虚构的普通正文。
2. 保存为本目录的 `basic-selection.docx`，保持打开并处于激活状态。
3. 不使用正式工作文档、敏感内容、编辑保护、修订或内容控件；文末不要放表格。
4. 在项目根目录运行只读检查：

```powershell
.\.venv\Scripts\python.exe scripts/check_active_document.py --expected-document "test-documents/basic-selection.docx"
```

检查 `status=matched` 和完整路径。身份匹配不授权写入，`ReadOnly=false` 也不代表全部编辑条件已通过。

读取选区前需手动选中文字。追加按 README 的预览与确认流程执行，不要求选区；程序不自动打开、切换、保存或关闭文档。用户自行核对修改并在 Word 中撤销。

本目录直接存放的 `.docx` 和 Word 的 `~$` 锁文件由 `.gitignore` 排除，不强制提交个人文件。新增嵌套测试目录时先检查忽略规则；迁移本目录前须同步检查代码的路径允许范围。
