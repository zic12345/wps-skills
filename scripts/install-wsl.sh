#!/usr/bin/env bash
# Input: WSL 环境（已装 Windows 版 WPS Office、Node.js >= 18）
# Output: 编译好的 wps-office-mcp dist 产物与 MCP 配置提示
# Pos: WSL 一键安装脚本。一旦我被修改，请更新 INSTALL.md 中的 WSL 章节。
#
# WSL（Windows Subsystem for Linux）下使用宿主 Windows 版 WPS Office：
# MCP Server 运行在 Linux 侧，通过 WSL interop 启动宿主机 powershell.exe，
# 复用 wps-office-mcp/scripts/wps-com.ps1 的 COM 通道（Ket/Kwps/Kwpp）。
# COM 通道不依赖 WPS JS 加载项，无需安装 wps-claude-addon。
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
MCP_DIR="$PROJECT_ROOT/wps-office-mcp"

c_info()  { printf '\033[36m%s\033[0m\n' "$*"; }
c_ok()    { printf '\033[32m✓ %s\033[0m\n' "$*"; }
c_warn()  { printf '\033[33m⚠ %s\033[0m\n' "$*"; }
c_err()   { printf '\033[31m✗ %s\033[0m\n' "$*"; }

c_info ">>> [1/5] 检测 WSL 环境..."
if ! grep -qiE 'microsoft|wsl' /proc/version 2>/dev/null \
   && [ ! -e /proc/sys/fs/binfmt_misc/WSLInterop ] \
   && [ ! -e /proc/sys/fs/binfmt_misc/WSLInterop-late ]; then
  c_err "当前不是 WSL 环境（/proc/version 无 microsoft 标记）。原生 Linux 请改用 scripts/install.sh"
  exit 1
fi
c_ok "WSL 环境确认"

c_info ">>> [2/5] 检测 WSL interop 与 powershell.exe..."
PS_BIN=""
if command -v powershell.exe >/dev/null 2>&1; then
  PS_BIN="powershell.exe"
elif [ -x /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe ]; then
  PS_BIN="/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
else
  c_err "找不到 powershell.exe。请确认 WSL interop 已启用（/etc/wsl.conf 中 [interop] enabled=true）"
  exit 1
fi
PS_VERSION="$("$PS_BIN" -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>/dev/null | tr -d '\r')"
c_ok "PowerShell 可用: $PS_BIN (v$PS_VERSION)"

c_info ">>> [3/5] 检测 Windows 侧 WPS Office 与 COM 注册..."
COM_CHECK="$("$PS_BIN" -NoProfile -Command '
$ok = $true
foreach ($id in "Ket.Application","Kwps.Application","Kwpp.Application") {
  if (-not (Test-Path "HKCU:\Software\Classes\$id") -and -not (Test-Path "HKLM:\Software\Classes\$id")) {
    Write-Output "MISSING:$id"; $ok = $false
  }
}
if ($ok) { Write-Output "ALL_REGISTERED" }
' 2>/dev/null | tr -d '\r')"
if [ "$COM_CHECK" != "ALL_REGISTERED" ]; then
  c_warn "部分 WPS COM ProgID 未注册（$COM_CHECK）"
  c_warn "请先在 Windows 侧完整安装并至少启动过一次 WPS Office（表格/文字/演示各一次）"
fi

# COM 冒烟（不依赖 WPS 是否已启动；GetActiveObject 失败时脚本会新建实例）
PS_SCRIPT_WIN="$(wslpath -w "$MCP_DIR/scripts/wps-com.ps1")"
PONG="$("$PS_BIN" -ExecutionPolicy Bypass -NoProfile -File "$PS_SCRIPT_WIN" -Action ping -Params '{}' 2>/dev/null | tr -d '\r')"
if echo "$PONG" | grep -q '"success":true'; then
  c_ok "wps-com.ps1 COM 桥接冒烟通过: $PONG"
else
  c_err "COM 桥接冒烟失败: $PONG"
  c_warn "可能原因：PowerShell 执行策略限制，或 WPS Office 未正确安装"
  exit 1
fi

c_info ">>> [4/5] 检测 Node.js 并安装依赖、编译..."
if ! command -v node >/dev/null 2>&1; then
  c_err "未找到 node，请先安装 Node.js >= 18（建议用 nvm）"
  exit 1
fi
NODE_MAJOR="$(node -p 'process.versions.node.split(".")[0]')"
if [ "$NODE_MAJOR" -lt 18 ]; then
  c_err "Node.js 版本过低: $(node -v)，需要 >= 18"
  exit 1
fi
c_ok "Node.js $(node -v)"

# 某些受限环境 ~/.npm 属主不对，用本地缓存目录规避（无需 sudo）
NPM_CACHE="${NPM_CACHE:-/tmp/npm-cache-wps-mcp}"
( cd "$MCP_DIR" && npm install --cache "$NPM_CACHE" )
( cd "$MCP_DIR" && npm run build )
[ -f "$MCP_DIR/dist/index.js" ] || { c_err "编译产物 dist/index.js 不存在"; exit 1; }
c_ok "编译完成: $MCP_DIR/dist/index.js"

c_info ">>> [5/5] 生成 MCP Client 配置..."
DIST_WIN="$(wslpath -w "$MCP_DIR/dist/index.js")"
cat <<EOF

✓✓✓ 安装完成 ✓✓✓

MCP Server (stdio) 启动命令：
  command: node
  args:    ["$MCP_DIR/dist/index.js"]

Claude Code settings.json 示例：
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["$MCP_DIR/dist/index.js"]
    }
  }
}

说明：
- WSL 走 PowerShell COM 通道，不需要安装 WPS JS 加载项，也不需要轮询服务。
- 日志默认写到 ~/.wps-office-mcp/logs；该目录不可写时自动回退到 ${TMPDIR:-/tmp}/wps-office-mcp-logs，
  也可用环境变量 WPS_MCP_LOG_DIR 指定。
- 可选环境变量：
    WPS_POWERSHELL_BIN  覆盖 powershell.exe 路径
    MCP_CONSOLE_LOG=1  同时输出控制台日志（调试用，勿在 stdio 生产模式开启）
- Windows 脚本路径: $PS_SCRIPT_WIN
- Windows 入口路径: $DIST_WIN

下一步：重启你的 MCP 客户端（Claude Code / Cursor / DSH 等），打开一个 WPS 文档后调用 wps_common_ping 验证。
EOF
