mac-poll-server.ts, wps-client.ts, wps-keepalive.ts
WPS 客户端与平台通信实现目录。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。

平台路由（wps-client.ts）：
- Windows（os.platform()=win32）：spawn powershell 执行 scripts/wps-com.ps1，走 WPS COM（Ket/Kwps/Kwpp）。
- WSL（Linux 且检测到 WSL：/proc/sys/fs/binfmt_misc/WSLInterop[‑late] 存在，或 /proc/version 含 microsoft）：
  spawn 宿主机 powershell.exe（PATH 优先，回退 /mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe，
  可用环境变量 WPS_POWERSHELL_BIN 覆盖），脚本路径经 `wslpath -w` 转成 \\wsl.localhost\... 或盘符 UNC；
  params 中路径类键（path/filePath/outputPath/imagePath/source/target/destination/sources 等）递归转 Windows 路径。
  COM 通道不依赖 JS 加载项，无需轮询。
- macOS / 原生 Linux：反向轮询（mac-poll-server，默认端口 58891，WPS_MCP_PORT 可改）。
