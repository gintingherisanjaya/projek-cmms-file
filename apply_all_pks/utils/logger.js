import fs from 'node:fs';

const wrap = (code) => (text) => `\u001b[${code}m${text}\u001b[0m`;

export const gray = wrap('90');
export const cyan = wrap('36');
export const yellow = wrap('33');
export const green = wrap('32');
export const red = wrap('31');
export const bold = wrap('1');

/** Strip ANSI escape codes for plain log file output. */
function stripAnsi(str) {
  return String(str).replace(/\u001b\[[0-9;]*m/g, '');
}

/** Optional log file stream; when set, all log* also append (plain text) to this file. */
let logFileStream = null;
let dashboardSink = null;

/**
 * Start teeing all log output to a file. Call once at script start.
 * @param {string} filePath - Path to .log or .txt file.
 */
export function setLogFile(filePath) {
  if (logFileStream) logFileStream.end();
  logFileStream = fs.createWriteStream(filePath, { flags: 'a' });
  logFileStream.write(`[${new Date().toISOString()}] Log started.\n`);
}

/** Stop writing to the log file and close it. Call when script finishes. */
export function closeLogFile() {
  if (logFileStream) {
    logFileStream.end();
    logFileStream = null;
  }
}

export function setDashboardSink(sink) {
  dashboardSink = sink || null;
}

export function clearDashboardSink() {
  dashboardSink = null;
}

function tee(level, ...args) {
  const raw = args.join(' ');
  const line = `[${new Date().toISOString()}] [${level}] ${stripAnsi(raw)}\n`;
  if (logFileStream && logFileStream.writable) {
    logFileStream.write(line);
  }
}

function emitConsoleOrDashboard(level, message) {
  if (dashboardSink?.pushLog) {
    dashboardSink.pushLog(level ? `[${level}] ${message}` : message);
    return;
  }
  console.log(message);
}

export const log = (...args) => {
  const raw = args.join(' ');
  emitConsoleOrDashboard('', raw);
  if (logFileStream && logFileStream.writable) {
    logFileStream.write(`[${new Date().toISOString()}] ${stripAnsi(raw)}\n`);
  }
};
export const logInfo = (...args) => {
  const msg = cyan(args.join(' '));
  emitConsoleOrDashboard('INFO', msg);
  tee('INFO', args.join(' '));
};
export const logWarn = (...args) => {
  const msg = yellow(args.join(' '));
  emitConsoleOrDashboard('WARN', msg);
  tee('WARN', args.join(' '));
};
export const logError = (...args) => {
  const msg = red(args.join(' '));
  emitConsoleOrDashboard('ERROR', msg);
  tee('ERROR', args.join(' '));
};
export const logSuccess = (...args) => {
  const msg = green(args.join(' '));
  emitConsoleOrDashboard('SUCCESS', msg);
  tee('SUCCESS', args.join(' '));
};
