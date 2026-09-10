# MCP 确认表单兼容性问题

记录：2026-09-10。更新：2026-09-10。状态：schema 修复及自动化回归通过；真实 Codex 确认追加成功，用户确认原生撤销可用。修复后主动取消确认的真实界面测试仍待完成。

## 现象与根因

WordBridge 已加载，文档身份与追加预览通过，但两次实际 MCP 追加调用均返回 `approval_cancel`、`write_attempted=false`。不能据此断言用户手动取消。

已观察到的客户端解析错误为：

```text
failed to parse typed MCP elicitation schema
unknown field `title`, expected one of `$schema`, `type`, `properties`, `required`
```

原因链：Pydantic 为 `AppendApproval` 自动生成顶层 `title` → MCP SDK 保留该字段发出确认表单 → 当前 Codex 的严格解析器拒绝 → 客户端返回取消 → WordBridge 在写入前安全退出。此故障不在 Word COM 的插入阶段，也不是缺少 `append_text` 工具或需要自动保存。

这是本次实测客户端与 `mcp==2.2.0` / `pydantic==2.13.5` 组合的兼容边界，不应推广为所有 JSON Schema 或所有 MCP 客户端都禁止 `title`。

## 最小修复

在 `scripts/mcp_server.py` 的 `AppendApproval` 配置 `json_schema_extra` 回调，只删除确认表单顶层 `title`。保留 `properties.confirm.title`（中文复选框标签）、`required=["confirm"]` 和 `StrictBool`；不递归删字段、不更改第三方安装包、不移除确认机制或文档摘要检查。

修复后 SDK 实际生成：

```json
{
  "properties": {
    "confirm": {
      "title": "我确认追加上述文字到上述文档",
      "type": "boolean"
    }
  },
  "required": ["confirm"],
  "type": "object"
}
```

## 验证与后续检查

- 已完成：先添加回归测试，在修复前确认因多出的根 `title` 失败；修复后全部 56 项自动化测试通过，`pip check` 通过。
- 已完成：测试调用 SDK 的 `render_elicitation_schema()` 检查真实生成结构，并在 MCP 确认回调中检查实际收到的 `requested_schema`，而不是只检查 Python 类型定义。
- 已完成：保留严格必填布尔值校验；取消、拒绝、未勾选、伪造确认、预览过期和不确定结果不自动重试等原有测试继续通过。
- 已完成：真实 Codex MCP 确认写入链路返回成功，用户确认写入及原生撤销可用；证据及边界见下节。
- 待完成：修复后在真实确认界面主动取消并核对不写入；此前解析失败触发的自动取消不能替代这项验收。中文标签的视觉呈现及保存状态也未在本次通过独立界面或元数据检查核验。

## 真实 MCP 闭环验收

2026-09-10，本机普通正文测试中，真实 MCP 确认追加成功，用户确认原生撤销可用。撤销后未独立读取正文摘要，保存状态亦未单独核验；不代表复杂格式、多客户端或其他客户端均已验收。

## 升级回归检查

运行自动化检查（项目根目录，不连接真实 Word）：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m pip check
```

今后新增确认表单或升级 SDK/Pydantic/客户端时，遵守项目 `AGENTS.md` 的公共验证要求：检查 SDK 渲染结果及协议上的表单结构，维持确认与状态校验，再做真实客户端验证。当前兼容性测试仅允许根字段 `$schema`、`type`、`properties`、`required`；新增其他根字段时应先核对客户端支持，不能直接放宽测试。字段内部的显示标签不属于本次禁止范围。

如果再次出现 `approval_cancel`，先检查客户端是否报告表单解析失败；不要反复触发写入、推断用户点了取消，或通过自动同意、换旧服务、直接调用底层写入来绕过确认。

## 参考与边界

[OpenAI Docs：App Server](https://learn.chatgpt.com/docs/app-server) 说明 MCP 确认请求及 accept/decline/cancel 响应机制；本次根字段限制以本机实际错误和生成结果为依据，不是从文档泛化出的通用限制。

以后升级仍须重新验证真实客户端，不能仅以模拟测试通过代替。
