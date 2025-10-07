export const logLevels = {
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
};

class Logger {
  constructor() {
    this.level = parseInt(localStorage.getItem('logLevel') || logLevels.INFO);
  }
  
  setLevel(level) {
    this.level = level;
    localStorage.setItem('logLevel', level);
  }
  
  debug(message, ...args) {
    if (this.level >= logLevels.DEBUG) {
      console.debug(`[DEBUG] ${message}`, ...args);
      this._saveLog('DEBUG', message, args);
    }
  }
  
  info(message, ...args) {
    if (this.level >= logLevels.INFO) {
      console.info(`[INFO] ${message}`, ...args);
      this._saveLog('INFO', message, args);
    }
  }
  
  warn(message, ...args) {
    if (this.level >= logLevels.WARN) {
      console.warn(`[WARN] ${message}`, ...args);
      this._saveLog('WARN', message, args);
    }
  }
  
  error(message, error, ...args) {
    if (this.level >= logLevels.ERROR) {
      console.error(`[ERROR] ${message}`, error, ...args);
      this._saveLog('ERROR', message, [error, ...args]);
    }
  }
  
  _saveLog(level, message, args) {
    try {
      const logs = JSON.parse(localStorage.getItem('appLogs') || '[]');
      logs.push({
        time: new Date().toISOString(),
        level,
        message,
        details: args.map(arg => {
          if (arg instanceof Error) {
            return {message: arg.message, stack: arg.stack};
          }
          return String(arg);
        })
      });
      
      // 保留最近100条日志
      while (logs.length > 100) {
        logs.shift();
      }
      
      localStorage.setItem('appLogs', JSON.stringify(logs));
    } catch (e) {
      console.error('保存日志失败:', e);
    }
  }
  
  getLogs() {
    return JSON.parse(localStorage.getItem('appLogs') || '[]');
  }
  
  clearLogs() {
    localStorage.removeItem('appLogs');
  }
}

export default new Logger();

