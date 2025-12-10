const { Worker } = require('worker_threads');
const path = require('path');
const { EventEmitter } = require('events');
const { logger } = require('./logger');

/**
 * Excel Worker Manager for background processing
 * Handles large Excel file operations without blocking the main thread
 */
class ExcelWorkerManager extends EventEmitter {
  constructor() {
    super();
    this.workers = new Map(); // workerId -> worker instance
    this.currentWorkerId = 0;
  }

  /**
   * Convert Excel worksheet to CSV in background
   * @param {string} filePath - Path to Excel file
   * @param {string} sheetName - Worksheet name to convert
   * @param {string} outputPath - Output CSV file path
   * @param {object} options - Conversion options
   * @returns {Promise} - Promise with conversion result
   */
  async convertWorksheetToCSV(filePath, sheetName, outputPath, options = {}) {
    return this.processExcelFile(filePath, {
      operation: 'convertToCSV',
      sheetName,
      outputPath,
      maxRows: options.maxRows,
      maxFileSize: options.maxFileSize
    });
  }

  /**
   * Process Excel file in background worker
   * @param {string} filePath - Path to Excel file
   * @param {object} options - Processing options
   * @returns {Promise} - Promise with worker ID for tracking
   */
  async processExcelFile(filePath, options = {}) {
    return new Promise((resolve, reject) => {
      const workerId = ++this.currentWorkerId;
      const timer = logger.time(`Excel processing worker ${workerId}`);

      logger.info('Starting Excel worker', { workerId, filePath, options });

      // Create worker with the Excel processing script
      const worker = new Worker(path.join(__dirname, 'excelWorkerThread.js'), {
        workerData: {
          filePath,
          options,
          workerId
        }
      });

      // Store worker reference
      this.workers.set(workerId, worker);

      // Handle worker messages
      worker.on('message', (message) => {
        const { type, data, error, workerId: msgWorkerId } = message;

        switch (type) {
          case 'progress':
            logger.debug('Excel worker progress', { workerId: msgWorkerId, progress: data });
            this.emit('progress', { workerId: msgWorkerId, ...data });
            break;

          case 'worksheets':
            logger.info('Excel worker found worksheets', {
              workerId: msgWorkerId,
              count: data.worksheets.length
            });
            this.emit('worksheets', { workerId: msgWorkerId, ...data });
            break;

          case 'success':
            timer.end('completed successfully');
            logger.info('Excel worker completed', { workerId: msgWorkerId });
            this.cleanup(workerId);
            resolve({ workerId: msgWorkerId, ...data });
            break;

          case 'error':
            timer.end('failed');
            logger.error('Excel worker failed', { workerId: msgWorkerId, error });
            this.cleanup(workerId);
            reject(new Error(error));
            break;

          default:
            logger.warn('Unknown worker message type', { type, workerId: msgWorkerId });
        }
      });

      // Handle worker errors
      worker.on('error', (error) => {
        timer.end('worker error');
        logger.error('Excel worker error', { workerId, error: error.message });
        this.cleanup(workerId);
        reject(error);
      });

      // Handle worker exit
      worker.on('exit', (code) => {
        if (code !== 0) {
          timer.end('worker exit with error');
          logger.error('Excel worker exited with error', { workerId, code });
          this.cleanup(workerId);
          reject(new Error(`Worker exited with code ${code}`));
        }
      });

      // Start processing
      worker.postMessage({ type: 'start' });
    });
  }

  /**
   * Cancel a running worker
   * @param {number} workerId - Worker ID to cancel
   */
  cancelWorker(workerId) {
    const worker = this.workers.get(workerId);
    if (worker) {
      logger.info('Cancelling Excel worker', { workerId });
      worker.terminate();
      this.cleanup(workerId);
      this.emit('cancelled', { workerId });
    }
  }

  /**
   * Cancel all running workers
   */
  cancelAllWorkers() {
    logger.info('Cancelling all Excel workers', { count: this.workers.size });
    for (const [workerId] of this.workers) {
      this.cancelWorker(workerId);
    }
  }

  /**
   * Get status of all workers
   */
  getWorkersStatus() {
    return Array.from(this.workers.keys()).map(workerId => ({
      workerId,
      status: 'running'
    }));
  }

  /**
   * Clean up worker resources
   * @private
   */
  cleanup(workerId) {
    const worker = this.workers.get(workerId);
    if (worker) {
      this.workers.delete(workerId);
    }
  }

  /**
   * Shutdown all workers
   */
  shutdown() {
    logger.info('Shutting down Excel worker manager', { activeWorkers: this.workers.size });
    this.cancelAllWorkers();
  }
}

// Create singleton instance
const excelWorkerManager = new ExcelWorkerManager();

// Clean up on process exit
process.on('exit', () => {
  excelWorkerManager.shutdown();
});

process.on('SIGTERM', () => {
  excelWorkerManager.shutdown();
  process.exit(0);
});

module.exports = {
  ExcelWorkerManager,
  excelWorkerManager
};