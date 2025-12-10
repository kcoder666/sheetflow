const { validateFilePath, validateDirectoryPath, sanitizeSheetName } = require('../lib/security');
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');

describe('Security Module', () => {
  let testDir;
  let testFile;

  beforeEach(() => {
    // Create a temporary directory for testing
    testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheetflow-test-'));

    // Create a test Excel file
    testFile = path.join(testDir, 'test.xlsx');
    fs.writeFileSync(testFile, 'fake excel content'); // Not a real Excel file, but good for path tests
  });

  afterEach(() => {
    // Clean up test files
    try {
      if (fs.existsSync(testFile)) fs.unlinkSync(testFile);
      if (fs.existsSync(testDir)) fs.rmdirSync(testDir);
    } catch (error) {
      // Ignore cleanup errors
    }
  });

  describe('validateFilePath', () => {
    test('should validate existing Excel files', () => {
      expect(() => validateFilePath(testFile)).not.toThrow();
    });

    test('should reject null or undefined paths', () => {
      expect(() => validateFilePath(null)).toThrow('Invalid file path');
      expect(() => validateFilePath(undefined)).toThrow('Invalid file path');
      expect(() => validateFilePath('')).toThrow('Invalid file path');
    });

    test('should reject non-string paths', () => {
      expect(() => validateFilePath(123)).toThrow('Invalid file path');
      expect(() => validateFilePath({})).toThrow('Invalid file path');
    });

    test('should reject non-existent files', () => {
      expect(() => validateFilePath('/nonexistent/file.xlsx')).toThrow('File does not exist');
    });

    test('should reject directories', () => {
      expect(() => validateFilePath(testDir)).toThrow('Path is not a file');
    });

    test('should reject non-Excel file extensions', () => {
      const txtFile = path.join(testDir, 'test.txt');
      fs.writeFileSync(txtFile, 'text content');

      expect(() => validateFilePath(txtFile)).toThrow('Invalid file type');

      fs.unlinkSync(txtFile);
    });

    test('should accept valid Excel extensions', () => {
      const extensions = ['.xlsx', '.xls', '.xlsm'];

      extensions.forEach(ext => {
        const file = path.join(testDir, `test${ext}`);
        fs.writeFileSync(file, 'content');

        expect(() => validateFilePath(file)).not.toThrow();

        fs.unlinkSync(file);
      });
    });
  });

  describe('validateDirectoryPath', () => {
    test('should validate existing directories', () => {
      expect(() => validateDirectoryPath(testDir)).not.toThrow();
    });

    test('should reject null or undefined paths', () => {
      expect(() => validateDirectoryPath(null)).toThrow('Invalid directory path');
      expect(() => validateDirectoryPath(undefined)).toThrow('Invalid directory path');
      expect(() => validateDirectoryPath('')).toThrow('Invalid directory path');
    });

    test('should reject non-existent directories', () => {
      expect(() => validateDirectoryPath('/nonexistent/directory')).toThrow('Directory does not exist');
    });

    test('should reject files', () => {
      expect(() => validateDirectoryPath(testFile)).toThrow('Path is not a directory');
    });
  });

  describe('sanitizeSheetName', () => {
    test('should return valid sheet names unchanged', () => {
      expect(sanitizeSheetName('Sheet1')).toBe('Sheet1');
      expect(sanitizeSheetName('Data_2023')).toBe('Data_2023');
    });

    test('should reject null or undefined names', () => {
      expect(() => sanitizeSheetName(null)).toThrow('Invalid sheet name');
      expect(() => sanitizeSheetName(undefined)).toThrow('Invalid sheet name');
      expect(() => sanitizeSheetName('')).toThrow('Invalid sheet name');
    });

    test('should reject non-string names', () => {
      expect(() => sanitizeSheetName(123)).toThrow('Invalid sheet name');
      expect(() => sanitizeSheetName({})).toThrow('Invalid sheet name');
    });

    test('should remove dangerous characters', () => {
      expect(sanitizeSheetName('Sheet<1>')).toBe('Sheet1');
      expect(sanitizeSheetName('Data:Sheet')).toBe('DataSheet');
      expect(sanitizeSheetName('Path/To\\File')).toBe('PathToFile');
    });

    test('should trim whitespace', () => {
      expect(sanitizeSheetName('  Sheet1  ')).toBe('Sheet1');
    });

    test('should reject names that become empty after sanitization', () => {
      expect(() => sanitizeSheetName('<>')).toThrow('Sheet name cannot be empty after sanitization');
      expect(() => sanitizeSheetName('   ')).toThrow('Sheet name cannot be empty after sanitization');
    });
  });
});

// Helper function to run tests if this file is executed directly
if (require.main === module) {
  console.log('Running security tests...');

  // Simple test runner
  const runTests = async () => {
    const tests = [
      () => {
        console.log('✓ Testing validateFilePath with valid file');
        const testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'test-'));
        const testFile = path.join(testDir, 'test.xlsx');
        fs.writeFileSync(testFile, 'content');
        validateFilePath(testFile);
        fs.unlinkSync(testFile);
        fs.rmdirSync(testDir);
      },
      () => {
        console.log('✓ Testing sanitizeSheetName');
        if (sanitizeSheetName('Sheet<1>') !== 'Sheet1') throw new Error('Failed');
      }
    ];

    for (const test of tests) {
      try {
        test();
      } catch (error) {
        console.error('✗ Test failed:', error.message);
      }
    }

    console.log('Tests completed');
  };

  runTests();
}