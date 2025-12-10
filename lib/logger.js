const fs = require('fs');
const path = require('path');
const os = require('os');

/**
 * Simple structured logger for SheetFlow
 */

const LOG_LEVELS = {
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
};

const LOG_LEVEL_NAMES = {
  0: 'ERROR',
  1: 'WARN',
  2: 'INFO',
  3: 'DEBUG'
};

class Logger {
  constructor(options = {}) {
    this.level = options.level || LOG_LEVELS.INFO;
    this.console = options.console !== false; // Default to true
    this.file = options.file || null;
    this.maxFileSize = options.maxFileSize || 10 * 1024 * 1024; // 10MB default
    this.maxFiles = options.maxFiles || 5;

    // Create log directory if file logging is enabled
    if (this.file) {
      const logDir = path.dirname(this.file);
      if (!fs.existsSync(logDir)) {
        fs.mkdirSync(logDir, { recursive: true });
      }
    }
  }

  formatMessage(level, message, meta = {}) {
    const timestamp = new Date().toISOString();
    const levelName = LOG_LEVEL_NAMES[level];

    const logEntry = {
      timestamp,
      level: levelName,
      message,
      pid: process.pid,
      ...meta
    };

    return {
      formatted: `${timestamp} [${levelName}] ${message}${meta && Object.keys(meta).length ? ' ' + JSON.stringify(meta) : ''}`,
      structured: logEntry
    };
  }

  log(level, message, meta = {}) {
    if (level > this.level) {
      return; // Skip if level is below threshold
    }

    const { formatted, structured } = this.formatMessage(level, message, meta);

    // Console output
    if (this.console) {
      switch (level) {
        case LOG_LEVELS.ERROR:
          console.error(formatted);
          break;
        case LOG_LEVELS.WARN:
          console.warn(formatted);
          break;
        case LOG_LEVELS.DEBUG:
          console.debug(formatted);
          break;
        default:
          console.log(formatted);
      }
    }

    // File output
    if (this.file) {
      try {
        this.writeToFile(formatted + '\n');
      } catch (error) {
        console.error('Failed to write to log file:', error.message);
      }
    }
  }

  writeToFile(message) {
    // Check file size and rotate if necessary
    if (fs.existsSync(this.file)) {
      const stats = fs.statSync(this.file);
      if (stats.size >= this.maxFileSize) {
        this.rotateLogFile();
      }
    }

    // Append to log file
    fs.appendFileSync(this.file, message);
  }

  rotateLogFile() {
    try {
      // Move current log files
      for (let i = this.maxFiles - 1; i > 0; i--) {
        const oldFile = `${this.file}.${i}`;
        const newFile = `${this.file}.${i + 1}`;

        if (fs.existsSync(oldFile)) {
          if (i === this.maxFiles - 1) {
            // Delete oldest file
            fs.unlinkSync(oldFile);
          } else {
            fs.renameSync(oldFile, newFile);
          }
        }
      }

      // Move current log to .1
      if (fs.existsSync(this.file)) {
        fs.renameSync(this.file, `${this.file}.1`);
      }
    } catch (error) {
      console.error('Failed to rotate log file:', error.message);
    }
  }

  error(message, meta = {}) {
    this.log(LOG_LEVELS.ERROR, message, meta);
  }

  warn(message, meta = {}) {
    this.log(LOG_LEVELS.WARN, message, meta);
  }

  info(message, meta = {}) {
    this.log(LOG_LEVELS.INFO, message, meta);
  }

  debug(message, meta = {}) {
    this.log(LOG_LEVELS.DEBUG, message, meta);
  }

  // Performance timing helper
  time(label) {
    const startTime = Date.now();
    return {
      end: (message = '') => {
        const duration = Date.now() - startTime;
        this.debug(`${label} completed${message ? ': ' + message : ''}`, { duration: `${duration}ms` });
        return duration;
      }
    };
  }

  // Security event logging
  security(event, details = {}) {
    this.warn(`Security Event: ${event}`, {
      security: true,
      event,
      ...details,
      user: os.userInfo().username,
      hostname: os.hostname()
    });
  }
}

// Create default logger instance
const defaultLogger = new Logger({
  level: process.env.LOG_LEVEL ? LOG_LEVELS[process.env.LOG_LEVEL.toUpperCase()] : LOG_LEVELS.INFO,
  file: process.env.LOG_FILE || null
});

// Export both the class and a default instance
module.exports = {
  Logger,
  logger: defaultLogger,
  LOG_LEVELS
};