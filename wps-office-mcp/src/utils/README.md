error.ts, logger.ts
MCP 服务端通用工具与日志目录。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。

logger.ts 日志目录解析（resolveLogDir）：
WPS_MCP_LOG_DIR → $HOME/.wps-office-mcp/logs → os.tmpdir()/wps-office-mcp-logs，
首个通过「mkdir + 写探针文件」验证可写的目录生效；全部不可写时不挂文件 transport
（winston silent=true，绝不写 stdout 污染 MCP stdio 协议）。
这是为受限沙箱（如部分 WSL/容器 HOME 不可写）增加的回退，不影响普通 Windows/macOS/Linux 行为。
