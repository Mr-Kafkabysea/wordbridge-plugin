# WordBridge MCP

通过本地 MCP 服务连接正在运行的 Microsoft Word，读取当前文档及选区，并在用户确认后追加文字。

当前为 `0.2.1-dev` 开发验证版本，仅支持 Windows 桌面 Word 和已保存本地 .docx 的普通正文场景。没有替换选区、文档枚举或自动撤销工具，也不支持网页版 Word、macOS 和 Linux。尚未发布安装包或建立完整客户端兼容性矩阵。

## 环境与安装

当前验证环境为 Python 3.14.7（64 位）和桌面 Word 16.0；其他版本尚未系统验收。取得源码后，在项目根目录的 PowerShell 中运行：

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip check
```

已有 `.venv` 时跳过创建命令。无需激活环境；安装依赖需要网络。这些命令不操作 Word。

## 配置 MCP 客户端

客户端需要支持本地 stdio MCP 服务；写入还要求支持并向真实用户展示确认表单。以下是启动参数，不是特定客户端的配置文件。将 `<项目绝对路径>` 替换为实际路径：

| 设置 | 内容 |
| --- | --- |
| 服务名称 | `wordbridge` |
| 传输方式 | `stdio` |
| 启动程序 | `<项目绝对路径>\.venv\Scripts\python.exe` |
| 参数列表 | 一个参数：`<项目绝对路径>\scripts\mcp_server.py` |
| 工作目录（若支持） | `<项目绝对路径>` |
| 环境 | 保留 Windows Python 所需系统环境，包括 `WINDIR` |

脚本路径应作为单个参数传入。Word 和服务需在可相互访问的用户会话中运行。程序不自动启动 Word、打开文件或切换当前文档。直接运行服务会等待协议消息，没有普通终端输出是正常现象。

先完成下方只读检查，再测试确认写入。客户端不得自动同意确认；不支持确认的客户端不能写入。

## 已有工具

| 工具 | 用途 | 写入 |
| --- | --- | --- |
| `word_status` | 检查可访问的 Word COM 实例 | 否 |
| `get_active_document` | 核对当前文档与期望路径 | 否 |
| `get_selection` | 读取当前普通正文选区 | 否 |
| `preview_append_text` | 预览文末追加及状态摘要 | 否 |
| `append_text` | 用户确认后追加文字 | 是 |

## 验证范围

开发版读取接口支持省略路径：`get_active_document()` 自动识别当前文档，
`get_selection()` 自动核对当前文档并读取选中文字，无需用户手动提供路径。
也可传入 `expected_document`，严格核对之前识别的目标；相对路径仍以项目根目录为准。
接受任意目录内已保存的本地 `.docx`，不自动打开或切换文档。
只有光标、没有选中文字时返回 `empty_selection`，不会自动读取光标所在段落。
预览和写入仍必须传入目标路径，写入确认流程不变。
更新后需重新加载 MCP 服务及工具列表；旧进程仍使用旧接口。

截至 2026-09-10，64 项自动化模拟测试通过，依赖检查通过。此前 `0.1.0-dev` 在本机普通正文场景下完成真实 MCP 确认追加，用户确认原生撤销可用；该次 MCP 撤销后没有独立读取正文摘要。本版新增的无路径读取与目录范围扩大尚未完成真实 Word / 客户端验收。主动取消确认的真实界面测试、复杂格式、多实例、多客户端并发和另一台电脑的独立安装验收仍待验证。

目录放开不等于正式文档已验收。建议先使用文档副本；下方 `test-documents` 路径是安全练习示例，不是程序的目录白名单。

## 项目文档

产品配套的最小 [WordBridge Skill](skills/wordbridge/SKILL.md) 提供工具发现、任务路由和安全调用指引。它与 MCP 服务分开加载，仅把文件放在仓库里不会自动启用；当前未打包成插件，也未完成跨模型触发验收。Skill 不负责安装 Python、Word 或 MCP 服务。

- [开发规范](AGENTS.md)
- [更新日志](CHANGELOG.md)
- [第三方依赖](docs/第三方依赖.md)
- [MCP 确认表单兼容性问题](docs/MCP确认表单兼容性问题.md)

项目自身许可证尚未选定，当前没有 `LICENSE` 文件；依赖许可证不等于本项目许可证。正式公开前仍需完成许可证选择与发布检查。

## 只读连接检查

专用 Word 测试文件存放在 [test-documents/](test-documents/README.md)，自动化测试代码存放在 `tests/`。

手动打开 Word（可新建空白文档并保存为 `test-documents/basic-selection.docx`），切换回终端，在项目根目录运行：

```powershell
.\.venv\Scripts\python.exe scripts/check_word_connection.py
```

- `connected`：已读取一个现有 Word 实例的版本，退出码为 0；不表示已识别所有 Word 实例或目标文档。
- `not_available`：当前会话没有可访问的活动 Word COM 对象，退出码为 1；不能仅凭此状态断言 Word 未运行。
- `com_error`：COM 初始化、连接或版本读取失败，输出阶段与 HRESULT 错误码，退出码为 1。

模拟测试命令（不连接真实 Word）：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

实现依据：微软 [GetActiveObject](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-getactiveobject) 与 [Word Application.Version](https://learn.microsoft.com/en-us/office/vba/api/word.application.version) 文档。此连接程序独立编写，未提取第三方项目源码。

## MCP 通信检查

在项目根目录运行测试客户端，它会启动并关闭本地 MCP 服务，只调用 `word_status`：

```powershell
.\.venv\Scripts\python.exe scripts/check_mcp.py
```

当前服务入口为 `scripts/mcp_server.py`。已注册 `word_status`、`get_active_document`、`get_selection`、`preview_append_text` 和 `append_text`。前四个只读，最后一个需客户端向用户发起确认。

通过 MCP 追加时，先调用 `preview_append_text(expected_document, text)` 获取 `state`，再用同一目标、文字和 `expected_state=state` 调用 `append_text`。服务重新检查预览，要求客户端展示确认表单；用户接受且勾选确认后，再次检查文档状态才写入。客户端不得自动同意。不支持确认的客户端拒绝写入；状态过期时重新预览，结果不确定时禁止自动重试。

如需通过 MCP 检查真实测试文档及追加预览（仍不写入）：

```powershell
.\.venv\Scripts\python.exe scripts/check_mcp.py --expected-document "test-documents/basic-selection.docx" --preview-text "MCP 预览测试，不执行写入。"
```

`mcp_check=passed` 只表示协议调用完成；仍须检查各工具的业务 `status`。测试客户端不会调用 `append_text`，也不会自动接受写入确认；选区文字不显示在该检查脚本的汇总输出中，但实际 MCP 工具会将所选文字返回给调用者。

## 当前文档身份检查

在 Word 中激活测试文档，然后在项目根目录运行：

```powershell
.\.venv\Scripts\python.exe scripts/check_active_document.py --expected-document "test-documents/basic-selection.docx"
```

程序只连接已有 Word 实例，通过 `ActiveDocument` 读取名称、完整路径及 `ReadOnly`。期望文件必须是已存在的本地 `.docx`，不限制所在目录；程序不会打开文件、切换文档或读取正文。

- `matched`：完整路径匹配，退出码为 0。同名但不同目录的文件不会匹配。
- `different_document`：当前文档与期望文件不同，或没有本地完整路径，退出码为 1。
- `no_document`：连接的 Word 实例中 `Documents.Count` 为 0，退出码为 1。
- `invalid_document_file`、`not_available`、`com_error`、`path_error`：输入或检查失败，退出码为 1。

结果是一次检查时的元数据快照，不授权写入，也不能保证之后的当前文档不变。`ReadOnly=false` 不代表已通过编辑保护、选区或写入权限检查。此程序不处理受保护视图，也不枚举多个 Word 实例。

实现依据：微软 [ActiveDocument](https://learn.microsoft.com/en-us/office/vba/api/word.application.activedocument)、[FullName](https://learn.microsoft.com/en-us/office/vba/api/word.document.fullname)、[ReadOnly](https://learn.microsoft.com/en-us/office/vba/api/word.document.readonly) 文档。

## 读取测试文档的选中文字

在测试文档正文中选中一小段文字，切回终端，在项目根目录运行：

```powershell
.\.venv\Scripts\python.exe scripts/read_selection.py --expected-document "test-documents/basic-selection.docx"
```

- `selected`：返回选中文字、`start`、`end`、`story_type` 和 `character_count`，退出码为 0。文字只作为调用结果输出，不另写日志文件；终端或调用客户端仍可能保留输出。
- `empty_selection`：仅有光标或没有可返回的文字，退出码为 1，不自动选择内容。
- `selection_document_mismatch`：选区所属文档与期望路径不符，拒绝读取文字。
- `unsupported_selection`：首版只支持正文中的普通文字选区，暂不处理页眉、页脚、形状及其他选区类型。
- `selection_too_large`：选区跨度超过 10000 个 Word 位置单位，拒绝读取文字。

`start` 和 `end` 是 Word 在当前 story 中的位置，正文从 0 开始；`character_count` 是 Python 对返回字符串的计数。两者在包含 emoji 等内容时不必相等。原始段落标记等字符会保留，以 JSON 转义形式显示。

结果代表一次读取的选区范围，不保证用户后续编辑后仍有效。读取流程不设置选区、不替换文字、不保存或关闭文档。当前文档检查命令默认仍只读取元数据。

实现依据：微软 [Selection.Range](https://learn.microsoft.com/en-us/office/vba/api/word.selection.range)、[Range.Text](https://learn.microsoft.com/en-us/office/vba/api/word.range.text)、[Range.Start](https://learn.microsoft.com/en-us/office/vba/api/word.range.start) 文档。

## 在文末追加文字（试验功能）

默认只生成预览；在项目根目录运行：

```powershell
.\.venv\Scripts\python.exe scripts/append_text.py --expected-document "test-documents/basic-selection.docx" --text "这是一段由 WordBridge 追加的测试文字。"
```

预览返回完整路径、文末位置、待追加文字及 `state` 摘要。获得用户对目标与文字的明确批准后，使用相同文字和预览返回的摘要执行：

```powershell
.\.venv\Scripts\python.exe scripts/append_text.py --expected-document "test-documents/basic-selection.docx" --text "这是一段由 WordBridge 追加的测试文字。" --apply --expected-state "<预览返回的 state>"
```

- 使用零长度 Range 在正文最后一个段落标记前插入，保留已有文字，不调用 Selection 替换。不会自动添加新段落；如需换段，应在输入文字中明确包含换行。
- 不保存、不关闭文档，不切换或修改其他文档。输入的 LF/CRLF 换行统一为 Word 的 CR 段落标记。
- 期望文件必须是已保存的本地 `.docx`，并与 Word 当前文档匹配。只读、编辑保护、受保护视图、修订记录/修订开启、内容控件或文末表格等场景暂不支持。
- 试验阶段最多追加 10000 个 Python 字符，正文长度上限为 100000 个 Word 位置单位。预览会在内存读取目标文档正文以计算摘要，但不输出或保存原正文；待追加文字会显示在预览中。
- `state` 绑定目标路径、正文、正文长度和待追加文字，用于检测预览过期，不代表用户批准，也不覆盖所有格式或外部协作变化。调用者仍负责取得用户批准；执行时请暂停编辑同一文档，不并行运行写入脚本。
- 成功写入后检查新插入范围的文字，并使用 `WordBridge append text` 作为自定义撤销记录名称。撤销由用户在 Word 中执行；不保证视觉选区保持原样。
- `stale_preview` 表示预览已过期，应重新预览。`write_outcome_unknown` 或 `verification_failed` 表示可能已经写入，应先检查 Word，禁止自动重试或盲目撤销。

追加代码独立编写，未复制第三方项目代码，未新增依赖。实现依据：微软 [Range.InsertAfter](https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter)、[StartCustomRecord](https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.startcustomrecord)、[EndCustomRecord](https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.endcustomrecord) 文档。

写入不确定、超时或连接中断时，先检查 Word，不自动重试。单服务实例串行调用 COM，但不能阻止用户或其他进程编辑；写入期间不要同时编辑同一文档。确认表单展示目标路径和待追加文字，客户端可能保留输出，请仅使用虚构测试内容。
