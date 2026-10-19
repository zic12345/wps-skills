/**
 * Input: 平台信息与WPS调用参数
 * Output: WPS API 调用结果
 * Pos: 跨平台 WPS 客户端。一旦我被修改，请更新我的头部注释（Updated: 2026-10-19 新增 WSL interop → powershell.exe COM 通道），以及所属文件夹的md。
 * WPS通信客户端 - 老王的跨平台版
 * Windows: 通过PowerShell调用WPS COM接口
 * WSL: 通过 WSL interop 调用宿主机 powershell.exe，走同一条 Windows COM 通道（脚本与路径自动转 Windows 形式）
 * Mac/原生Linux: 通过反向轮询服务器（MCP Server当服务端，WPS加载项来轮询）
 *
 * 丢，为了兼容Mac老王可是费了老大劲了
 * WPS Mac加载项在沙箱里启动不了HTTP服务器，只能反过来搞！
 */

import { spawn, spawnSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import {
  WpsEndpointConfig,
  WpsApiRequest,
  WpsApiResponse,
  WpsAppType,
  WpsClientStatus,
  DocumentInfo,
  WorkbookInfo,
  PresentationInfo,
} from '../types/wps';
import { log, logRequest, logResponse } from '../utils/logger';
import { errorUtils } from '../utils/error';
import { macPollServer } from './mac-poll-server';

// 平台判断
const PLATFORM = os.platform();
const IS_MAC = PLATFORM === 'darwin';
const IS_WINDOWS = PLATFORM === 'win32';
const IS_LINUX = PLATFORM === 'linux';

// WSL（适用于 Linux 的 Windows 子系统）检测：binfmt 互操作注册存在，或 /proc/version 含 microsoft
function detectWsl(): boolean {
  if (!IS_LINUX) return false;
  try {
    if (fs.existsSync('/proc/sys/fs/binfmt_misc/WSLInterop') ||
        fs.existsSync('/proc/sys/fs/binfmt_misc/WSLInterop-late')) {
      return true;
    }
    return /microsoft|wsl/i.test(fs.readFileSync('/proc/version', 'utf8'));
  } catch {
    return false;
  }
}
const IS_WSL = detectWsl();

// Mac 和 原生Linux 通过反向轮询（wps-claude-assistant 加载项 HTTP 轮询模式）；
// Windows 与 WSL（interop 调宿主机 powershell.exe）走 PowerShell COM
const USE_POLL = IS_MAC || (IS_LINUX && !IS_WSL);

// 传输方式标识（日志/错误信息用）
const TRANSPORT_LABEL = IS_WINDOWS
  ? 'PowerShell COM'
  : IS_WSL
    ? 'WSL interop → powershell.exe COM'
    : `HTTP Poll (${PLATFORM})`;

// ==================== WSL interop 适配 ====================

// WSL 下优先直接用 powershell.exe（依赖 PATH 中存在 Windows 互操作目录），
// 找不到时回退到 System32 完整路径；均可通过环境变量 WPS_POWERSHELL_BIN 覆盖
function resolvePowerShellBin(): string {
  const override = process.env.WPS_POWERSHELL_BIN?.trim();
  if (override) return override;
  if (IS_WINDOWS) return 'powershell';
  const sysNative = '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe';
  return fs.existsSync(sysNative) ? sysNative : 'powershell.exe';
}
const POWERSHELL_BIN = resolvePowerShellBin();

// PowerShell脚本路径（dist/client 向上两级到包根再进 scripts）
const PS_SCRIPT_PATH_NIX = path.join(__dirname, '../../scripts/wps-com.ps1');

// wslpath 路径转换结果缓存（同步转换，只在首次遇到每个 WSL 路径时付出一次进程开销）
const winPathCache = new Map<string, string>();

/**
 * WSL 路径 -> Windows 路径（通过 wslpath -w）
 * /mnt/d/foo -> D:\foo；/home/... -> \\wsl.localhost\<distro>\home\...
 * 转换失败时原样返回
 */
function toWindowsPath(p: string): string {
  if (typeof p !== 'string' || !p) return p;
  // 已经是 Windows 路径（盘符或 UNC）则不处理
  if (/^[A-Za-z]:[\\/]/.test(p) || p.startsWith('\\\\')) return p;
  const cached = winPathCache.get(p);
  if (cached !== undefined) return cached;
  let result = p;
  try {
    const r = spawnSync('wslpath', ['-w', p], { encoding: 'utf8', timeout: 3000 });
    if (r.status === 0 && r.stdout.trim()) {
      result = r.stdout.trim();
    } else {
      log.warn('wslpath conversion failed, keep original path', { path: p, stderr: r.stderr });
    }
  } catch (err) {
    log.warn('wslpath unavailable, keep original path', { path: p, error: (err as Error).message });
  }
  winPathCache.set(p, result);
  return result;
}

// 参数中携带文件/目录路径的键名（WSL 下需转换为 Windows 路径，PS 端 COM 只认 Windows 路径）
const PATH_PARAM_KEYS = new Set([
  'path', 'filePath', 'outputPath', 'imagePath', 'savePath',
  'source', 'target', 'destination', 'sources', 'templatePath', 'exportPath',
]);

/**
 * 转换参数对象中的路径字段（支持字符串与数组，如 sources）
 */
function convertParamsPaths(params: Record<string, unknown>): Record<string, unknown> {
  if (!IS_WSL) return params;
  const convert = (v: unknown): unknown => {
    if (typeof v === 'string') return toWindowsPath(v);
    if (Array.isArray(v)) return v.map(convert);
    return v;
  };
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(params)) {
    out[k] = PATH_PARAM_KEYS.has(k) ? convert(v) : v;
  }
  return out;
}

// Mac轮询服务器端口（支持环境变量 WPS_MCP_PORT 配置）
const MAC_POLL_PORT = parseInt(process.env.WPS_MCP_PORT || '58891', 10);

/**
 * 执行Mac轮询调用
 * 通过轮询服务器发送命令，等待WPS加载项取走并返回结果
 */
async function execMacPoll(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  log.debug('Executing Mac Poll', { action, params });

  try {
    // 确保轮询服务器已启动
    if (!macPollServer.isRunning) {
      log.info('[Mac] Starting poll server...');
      await macPollServer.start(MAC_POLL_PORT);
    }

    // 通过轮询服务器执行命令
    const result = await macPollServer.executeCommand(action, params);
    return result;
  } catch (error) {
    log.error('Mac Poll call failed', { action, error });
    throw error;
  }
}

// PowerShell 默认超时（毫秒）
const PS_TIMEOUT = 30000;

// ==================== PPT 目标文稿锁定（避免多文稿打开时 ActivePresentation 漂移）====================
// 通过 wps_ppt_set_active_target 设置后，所有 PRESENTATION 类调用（executeMethod 带 WpsAppType.PRESENTATION）
// 自动注入 presentationName，由 wps-com.ps1 的 Get-TargetPres 精确定位目标文稿；单次调用显式传 presentationName 可覆盖。
let pptTargetName: string | undefined;
export function setPptTarget(name?: string): void {
  pptTargetName = name && name.trim() ? name.trim() : undefined;
}
export function getPptTarget(): string | undefined {
  return pptTargetName;
}

/**
 * 执行PowerShell命令 (Windows / WSL)
 * WSL 下通过 interop 启动宿主机 powershell.exe：
 * - 脚本路径由 wslpath -w 转成 Windows 路径（\\wsl.localhost\... 或盘符路径）
 * - params 中的路径字段在 convertParamsPaths 内转换
 */
async function execPowerShell(action: string, rawParams: Record<string, unknown> = {}): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const params = convertParamsPaths(rawParams);
    // WSL 下脚本在 Linux 文件系统，需转换；Windows 原生路径原样使用
    const scriptPath = IS_WSL ? toWindowsPath(PS_SCRIPT_PATH_NIX) : PS_SCRIPT_PATH_NIX;
    // JSON参数通过spawn args数组传递，Node自动处理Windows引号转义
    const paramsJson = JSON.stringify(params);
    const args = [
      '-ExecutionPolicy', 'Bypass',
      '-NoProfile',
      '-File', scriptPath,
      '-Action', action,
      '-Params', paramsJson
    ];

    log.debug('Executing PowerShell', { action, params, bin: POWERSHELL_BIN, scriptPath });

    const ps = spawn(POWERSHELL_BIN, args, {
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
      // WSL interop 下保证输出为 UTF-8（PS 脚本自身也会设置 Console.OutputEncoding）
      env: { ...process.env }
    });

    let stdout = '';
    let stderr = '';
    let killed = false;

    // 超时保护：防止PowerShell进程挂起，超时时主动 SIGTERM 并记录 PID
    const timeoutHandle = setTimeout(() => {
      killed = true;
      log.warn('PowerShell timeout, killing process', { pid: ps.pid, action, timeoutMs: PS_TIMEOUT });
      ps.kill('SIGTERM');
      reject(new Error(`PowerShell 执行超时（${PS_TIMEOUT}ms）: ${action} (PID: ${ps.pid})`));
    }, PS_TIMEOUT);

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('close', (code) => {
      clearTimeout(timeoutHandle);
      if (killed) return; // 已超时处理，忽略后续事件

      if (code !== 0) {
        // 非零退出码：分别处理 stderr 有内容 / 为空两种路径，避免空 stderr 误入 JSON.parse
        if (stderr) {
          log.error('PowerShell error', { stderr, code, pid: ps.pid, action });
          reject(new Error(stderr));
        } else {
          log.error('PowerShell exited with non-zero code', { code, stdout: stdout.substring(0, 200), pid: ps.pid, action });
          reject(new Error(`PowerShell 退出码 ${code}: ${stdout.substring(0, 200) || '(无输出)'}`));
        }
        return;
      }

      try {
        const result = JSON.parse(stdout.trim());
        resolve(result);
      } catch (_e) {
        log.error('Failed to parse PowerShell output', { stdout: stdout.substring(0, 200), pid: ps.pid, action });
        reject(new Error(`PowerShell 输出解析失败（非有效JSON）: ${stdout.substring(0, 200)}`));
      }
    });

    ps.on('error', (err) => {
      clearTimeout(timeoutHandle);
      if (killed) return;
      log.error('PowerShell spawn error', { error: err.message, pid: ps.pid, action });
      reject(new Error(`无法启动 PowerShell 进程: ${err.message}（请确认 WSL 互操作已启用、可访问 powershell.exe，或用环境变量 WPS_POWERSHELL_BIN 指定路径）`));
    });
  });
}

/**
 * 统一执行接口 - 根据平台选择调用方式
 * Mac / 原生Linux: 反向轮询模式（MCP Server是服务端，WPS加载项来取命令）
 * Windows / WSL: PowerShell调用COM接口（WSL经 interop 调宿主机 powershell.exe）
 */
async function execWpsAction(action: string, params: Record<string, unknown> = {}): Promise<unknown> {
  if (USE_POLL) {
    return execMacPoll(action, params);
  } else {
    return execPowerShell(action, params);
  }
}

/**
 * WPS客户端类 - 跨平台通信
 * Windows: PowerShell COM桥接
 * WSL: PowerShell COM 桥接（interop 到 Windows 宿主机）
 * Mac/原生Linux: 反向轮询 WPS加载项
 */
export class WpsClient {
  private status: WpsClientStatus;

  constructor(_config?: Partial<WpsEndpointConfig>) {
    this.status = { connected: false };
    log.info('WPS Client initialized', { method: TRANSPORT_LABEL, platform: PLATFORM, isWsl: IS_WSL, powershell: USE_POLL ? undefined : POWERSHELL_BIN });
  }

  /**
   * 调用WPS接口（跨平台）
   */
  async invokeAction<T = unknown>(action: string, params: Record<string, unknown> = {}): Promise<WpsApiResponse<T>> {
    const startTime = Date.now();
    logRequest(action, params);

    try {
      const result = await execWpsAction(action, params) as WpsApiResponse<T>;
      const duration = Date.now() - startTime;
      logResponse(action, result.success, duration);

      if (result.success) {
        this.status.connected = true;
        this.status.lastHeartbeat = new Date();
      }

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      logResponse(action, false, duration);
      this.status.connected = false;
      throw errorUtils.wrap(error, `WPS 调用失败（${USE_POLL ? (IS_MAC ? 'macOS轮询' : 'Linux轮询') : (IS_WSL ? 'WSL PowerShell COM' : 'PowerShell COM')}）: ${action}`);
    }
  }

  /**
   * 兼容旧API
   */
  async callApi<T = unknown>(request: WpsApiRequest): Promise<WpsApiResponse<T>> {
    const actionMap: Record<string, string> = {
      // Excel 旧API
      'workbook.getActive': 'getActiveWorkbook',
      'cell.getValue': 'getCellValue',
      'cell.setValue': 'setCellValue',
      'range.getData': 'getRangeData',
      'range.setData': 'setRangeData',
      // Word 旧API
      'document.getActive': 'getActiveDocument',
      'document.getText': 'getDocumentText',
      'document.insertText': 'insertText',
      // PPT 旧API
      'presentation.getActive': 'getActivePresentation',
      'presentation.addSlide': 'addSlide',
      // 通用旧API
      'file.save': 'save',
      'file.saveAs': 'saveAs',
      'file.open': 'openFile',
      'ping': 'ping',
    };
    const action = actionMap[request.method] || request.method;
    return this.invokeAction<T>(action, request.params || {});
  }

  /**
   * 检查WPS连接状态
   */
  async checkConnection(): Promise<boolean> {
    try {
      const result = await this.invokeAction('ping');
      this.status.connected = result.success;
      return result.success;
    } catch {
      this.status.connected = false;
      this.status.error = 'Connection check failed';
      return false;
    }
  }

  /**
   * 获取客户端状态
   */
  getStatus(): WpsClientStatus {
    return { ...this.status };
  }

  // ==================== 表格操作 (WPS表格) ====================

  async getActiveWorkbook(): Promise<WorkbookInfo | null> {
    const response = await this.invokeAction<WorkbookInfo>('getActiveWorkbook');
    return response.success ? response.data || null : null;
  }

  async getCellValue(sheet: string | number, row: number, col: number): Promise<unknown> {
    const response = await this.invokeAction<{ value: unknown }>('getCellValue', { sheet, row, col });
    return response.data?.value;
  }

  async setCellValue(sheet: string | number, row: number, col: number, value: unknown): Promise<boolean> {
    const response = await this.invokeAction('setCellValue', { sheet, row, col, value });
    return response.success;
  }

  async getRangeData(sheet: string | number, range: string): Promise<unknown[][]> {
    const response = await this.invokeAction<{ data: unknown[][] }>('getRangeData', { sheet, range });
    return response.data?.data || [];
  }

  async setRangeData(sheet: string | number, range: string, data: unknown[][]): Promise<boolean> {
    const response = await this.invokeAction('setRangeData', { sheet, range, data });
    return response.success;
  }

  async setFormula(sheet: string | number, row: number, col: number, formula: string): Promise<boolean> {
    const response = await this.invokeAction('setFormula', { sheet, row, col, formula });
    return response.success;
  }

  // ==================== 文档操作 (WPS文字) ====================

  async getActiveDocument(): Promise<DocumentInfo | null> {
    const response = await this.invokeAction<DocumentInfo>('getActiveDocument');
    return response.success ? response.data || null : null;
  }

  async createDocument(): Promise<boolean> {
    const response = await this.invokeAction('createDocument');
    return response.success;
  }

  async insertText(text: string, position?: number): Promise<boolean> {
    const response = await this.invokeAction('insertText', { text, position });
    return response.success;
  }

  async getDocumentText(): Promise<string> {
    const response = await this.invokeAction<{ text: string }>('getDocumentText');
    return response.data?.text || '';
  }

  // ==================== 演示操作 (WPS演示) ====================

  async getActivePresentation(): Promise<PresentationInfo | null> {
    const response = await this.invokeAction<PresentationInfo>('getActivePresentation');
    return response.success ? response.data || null : null;
  }

  async createPresentation(): Promise<boolean> {
    const response = await this.invokeAction('createPresentation');
    return response.success;
  }

  async addSlide(layout?: string): Promise<boolean> {
    const response = await this.invokeAction('addSlide', { layout });
    return response.success;
  }

  // ==================== 通用操作 ====================

  async executeMethod<T = unknown>(
    method: string,
    params?: Record<string, unknown>,
    appType?: WpsAppType
  ): Promise<WpsApiResponse<T>> {
    let finalParams: Record<string, unknown> = params || {};
    // 锁定目标文稿后，为演示类调用自动注入 presentationName（已显式传入则不覆盖）
    if (appType === WpsAppType.PRESENTATION && pptTargetName && finalParams.presentationName === undefined) {
      finalParams = { ...finalParams, presentationName: pptTargetName };
    }
    return this.invokeAction<T>(method, finalParams);
  }

  async openFile(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('openFile', { path: filePath });
    return response.success;
  }

  async saveFile(_appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('save');
    return response.success;
  }

  async saveFileAs(filePath: string, _appType?: WpsAppType): Promise<boolean> {
    const response = await this.invokeAction('saveAs', { path: filePath });
    return response.success;
  }
}

// 导出单例
export const wpsClient = new WpsClient();

export default WpsClient;
