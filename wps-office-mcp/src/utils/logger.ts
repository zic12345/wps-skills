/**
 * Input: 日志消息与上下文
 * Output: 结构化日志输出
 * Pos: MCP 日志系统模块。一旦我被修改，请更新我的头部注释（Updated: 2026-10-19 WSL 支持：日志目录可写性探测+tmp回退），以及所属文件夹的md。
 * 日志工具 - 老王的日志系统
 * 写日志比写代码还规范，报错了至少知道是哪个SB搞出来的
 */

import winston from 'winston';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

/**
 * 日志级别枚举
 */
export enum LogLevel {
  ERROR = 'error',
  WARN = 'warn',
  INFO = 'info',
  DEBUG = 'debug',
}

/**
 * 自定义日志格式 - 老王风格
 */
const wangFormat = winston.format.printf(({ level, message, timestamp, ...meta }) => {
  const metaStr = Object.keys(meta).length ? ` | ${JSON.stringify(meta)}` : '';
  return `[${timestamp}] [${level.toUpperCase()}] ${message}${metaStr}`;
});

/**
 * 解析布尔环境变量
 */
const isEnvTrue = (value?: string): boolean => {
  if (!value) return false;
  return ['1', 'true', 'yes', 'on'].includes(value.toLowerCase());
};

/**
 * 解析可写的日志目录
 * 优先级：WPS_MCP_LOG_DIR 环境变量 → $HOME/.wps-office-mcp/logs → os.tmpdir() 回退
 * 全部不可写时返回 null（仅控制台输出，MCP stdio 下即等价于无文件日志，不影响协议）
 */
function resolveLogDir(): string | null {
  const candidates: string[] = [];
  if (process.env.WPS_MCP_LOG_DIR) {
    candidates.push(process.env.WPS_MCP_LOG_DIR);
  }
  const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir();
  if (homeDir) {
    candidates.push(path.join(homeDir, '.wps-office-mcp', 'logs'));
  }
  candidates.push(path.join(os.tmpdir(), 'wps-office-mcp-logs'));

  for (const dir of candidates) {
    try {
      fs.mkdirSync(dir, { recursive: true });
      // 触碰测试文件确认确实可写（某些沙箱允许 mkdir 但禁止写文件）
      const probe = path.join(dir, '.write-probe');
      fs.writeFileSync(probe, '');
      fs.rmSync(probe, { force: true });
      return dir;
    } catch {
      // 尝试下一个候选
    }
  }
  return null;
}

/**
 * 创建Logger实例
 */
const createLogger = (name: string): winston.Logger => {
  // MCP 默认走 stdio 协议，stdout 必须保持纯净；仅在显式开启时才输出 Console 日志
  const enableConsoleLog = isEnvTrue(process.env.MCP_CONSOLE_LOG);
  const transports: winston.transport[] = [];

  const logDir = resolveLogDir();
  if (logDir) {
    // 文件输出 - 错误单独记录，出问题好找
    transports.push(
      new winston.transports.File({
        filename: path.join(logDir, 'error.log'),
        level: 'error',
      }),
    );
    // 所有日志
    transports.push(
      new winston.transports.File({
        filename: path.join(logDir, 'combined.log'),
      }),
    );
  }

  if (enableConsoleLog) {
    transports.unshift(
      new winston.transports.Console({
        format: winston.format.combine(
          winston.format.colorize({ all: true }),
          winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
          wangFormat
        ),
      })
    );
  }

  return winston.createLogger({
    // 无可用 transport（日志目录不可写且未开控制台）时静默丢弃，禁止 winston 自行打日志到 stdout 污染 MCP 协议
    silent: transports.length === 0,
    level: process.env.LOG_LEVEL || 'info',
    format: winston.format.combine(
      winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
      winston.format.errors({ stack: true }),
      wangFormat
    ),
    defaultMeta: { service: name },
    transports,
  });
};

/**
 * 主Logger - MCP Server专用
 */
export const logger = createLogger('wps-mcp');

/**
 * 创建子Logger - 各模块用自己的
 */
export const createChildLogger = (moduleName: string): winston.Logger => {
  return logger.child({ module: moduleName });
};

/**
 * 快捷日志函数 - 懒人专用
 */
export const log = {
  /**
   * 记录信息日志
   */
  info: (message: string, meta?: Record<string, unknown>): void => {
    logger.info(message, meta);
  },

  /**
   * 记录警告日志
   */
  warn: (message: string, meta?: Record<string, unknown>): void => {
    logger.warn(message, meta);
  },

  /**
   * 记录错误日志 - 艹，又出错了
   */
  error: (message: string, error?: Error | unknown, meta?: Record<string, unknown>): void => {
    if (error instanceof Error) {
      logger.error(message, { ...meta, error: error.message, stack: error.stack });
    } else {
      logger.error(message, { ...meta, error });
    }
  },

  /**
   * 记录调试日志 - 排查问题用的
   */
  debug: (message: string, meta?: Record<string, unknown>): void => {
    logger.debug(message, meta);
  },
};

/**
 * 请求日志中间件风格的函数
 */
export const logRequest = (
  method: string,
  params?: Record<string, unknown>
): void => {
  logger.info(`[REQUEST] ${method}`, { params });
};

/**
 * 响应日志
 */
export const logResponse = (
  method: string,
  success: boolean,
  duration: number,
  data?: unknown
): void => {
  const level = success ? 'info' : 'error';
  logger.log(level, `[RESPONSE] ${method} - ${success ? 'SUCCESS' : 'FAILED'} (${duration}ms)`, {
    success,
    duration,
    ...(process.env.LOG_LEVEL === 'debug' ? { data } : {}),
  });
};

export default logger;
