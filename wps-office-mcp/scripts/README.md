wps-com.ps1
Windows PowerShell COM 桥接脚本目录（含图表/去重/上下文等扩展实现）。
一旦这里的结构发生变化，请务必更新我... 就像重新标记领地一样。

## 编码约定（改动必读）

- 文件必须保存为 **带 BOM 的 UTF-8 + CRLF**。Windows PowerShell 5.1 对无 BOM 的 .ps1
  会按系统 ANSI 代码页（中文系统为 GBK）解码，中文字符串/注释会乱码并引发连锁
  ParserError。用 node/python 改本文件后，务必确认首字节为 `EF BB BF`。
- 脚本由 Node 端以 `-File wps-com.ps1 -Action <name> -Params <json>` 调用，
  只向 stdout 输出一行压缩 JSON；错误也要走 `Output-Json`，避免向 stdout 写杂讯污染 stdio MCP。

## WSL 联调修复（2026-10-19）

- **Resolve-Sheet($workbook, $sheetSelector)**：统一解析 sheet 参数。
  - `$null`/空串/`0` → ActiveSheet（MCP 层未指定时用 `sheet || 0`）；
  - 数值或纯数字串 ≥1 → 按 1 基索引；其他字符串 → 按名称。
  - 修复未指定 sheet 时 `Sheets.Item(0)` 必抛 `DISP_E_BADINDEX (0x8002000B)` 的问题。
  - 全脚本所有 `$wb.Sheets.Item($p.sheet)` 类调用均已替换为 Resolve-Sheet。
- **switchSheet**：MCP 工具层传的参数名是 `name`（原脚本读 `$p.sheet`，参数名不匹配）。
  现读 `$p.name`（回退 `$p.sheet`），空值报错。脚本中有两个同名 switchSheet 块，均已修。
- **setRangeData**：逐格 `Value2 =` 在 WPS COM 下对 JSON 反序列化出的 Int32/Decimal
  报 `InvalidCastException`（字符串可写、数字丢失）。改为构造 `object[,]` 二维数组
  （`[Array]::CreateInstance([object], rows, cols)` + SetValue）后对整个 Range 一次性
  `Value2 =` 赋值，数字与字符串均正确写入。
