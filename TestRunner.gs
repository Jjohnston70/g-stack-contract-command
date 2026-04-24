/**
 * TNDS Test Runner Framework
 *
 * Automated testing framework for Google Apps Script modules.
 * Can be run via clasp run or from the Apps Script editor.
 *
 * Usage:
 *   clasp run runAllTests
 *   clasp run runModuleTests --params '["contract-tracker"]'
 */

// ============================================================================
// TEST RUNNER CORE
// ============================================================================

/**
 * Test result object
 */
class TestResult {
  constructor(testId, testName, module) {
    this.testId = testId;
    this.testName = testName;
    this.module = module;
    this.status = 'pending';
    this.message = '';
    this.duration = 0;
    this.timestamp = new Date();
  }

  pass(message = 'Passed') {
    this.status = 'passed';
    this.message = message;
    return this;
  }

  fail(message) {
    this.status = 'failed';
    this.message = message;
    return this;
  }

  skip(message = 'Skipped') {
    this.status = 'skipped';
    this.message = message;
    return this;
  }

  error(message) {
    this.status = 'error';
    this.message = message;
    return this;
  }
}

/**
 * Test Suite class
 */
class TestSuite {
  constructor(moduleName) {
    this.moduleName = moduleName;
    this.tests = [];
    this.results = [];
    this.beforeAll = null;
    this.afterAll = null;
    this.beforeEach = null;
    this.afterEach = null;
  }

  /**
   * Add a test case
   */
  addTest(testId, testName, testFn) {
    this.tests.push({ testId, testName, testFn });
    return this;
  }

  /**
   * Run all tests in the suite
   */
  run() {
    console.log(`\n========== Running ${this.moduleName} Tests ==========`);

    // Run beforeAll
    if (this.beforeAll) {
      try {
        this.beforeAll();
      } catch (e) {
        console.error('beforeAll failed: ' + e.message);
      }
    }

    // Run each test
    this.tests.forEach(test => {
      const result = new TestResult(test.testId, test.testName, this.moduleName);
      const startTime = Date.now();

      try {
        // Run beforeEach
        if (this.beforeEach) this.beforeEach();

        // Run the test
        const testOutput = test.testFn();

        // Check result
        if (testOutput === true || testOutput === undefined) {
          result.pass();
        } else if (testOutput === false) {
          result.fail('Test returned false');
        } else if (typeof testOutput === 'string') {
          result.fail(testOutput);
        } else if (testOutput && testOutput.success === false) {
          result.fail(testOutput.message || testOutput.error || 'Test failed');
        } else {
          result.pass(typeof testOutput === 'object' ? JSON.stringify(testOutput) : String(testOutput));
        }

        // Run afterEach
        if (this.afterEach) this.afterEach();

      } catch (e) {
        result.error(e.message);
      }

      result.duration = Date.now() - startTime;
      this.results.push(result);

      // Log result
      const icon = result.status === 'passed' ? '✓' : result.status === 'failed' ? '✗' : '○';
      console.log(`  ${icon} [${result.testId}] ${result.testName} (${result.duration}ms)`);
      if (result.status !== 'passed') {
        console.log(`    → ${result.message}`);
      }
    });

    // Run afterAll
    if (this.afterAll) {
      try {
        this.afterAll();
      } catch (e) {
        console.error('afterAll failed: ' + e.message);
      }
    }

    return this.getReport();
  }

  /**
   * Get test report
   */
  getReport() {
    const passed = this.results.filter(r => r.status === 'passed').length;
    const failed = this.results.filter(r => r.status === 'failed').length;
    const errors = this.results.filter(r => r.status === 'error').length;
    const skipped = this.results.filter(r => r.status === 'skipped').length;
    const total = this.results.length;

    return {
      module: this.moduleName,
      total,
      passed,
      failed,
      errors,
      skipped,
      passRate: total > 0 ? Math.round((passed / total) * 100) : 0,
      results: this.results.map(r => ({
        testId: r.testId,
        testName: r.testName,
        status: r.status,
        message: r.message,
        duration: r.duration
      }))
    };
  }
}

// ============================================================================
// ASSERTION HELPERS
// ============================================================================

/**
 * Assert that a value is truthy
 */
function assertTrue(value, message = 'Expected true') {
  if (!value) throw new Error(message);
  return true;
}

/**
 * Assert that a value is falsy
 */
function assertFalse(value, message = 'Expected false') {
  if (value) throw new Error(message);
  return true;
}

/**
 * Assert that two values are equal
 */
function assertEquals(expected, actual, message) {
  if (expected !== actual) {
    throw new Error(message || `Expected ${expected} but got ${actual}`);
  }
  return true;
}

/**
 * Assert that a value is not null/undefined
 */
function assertNotNull(value, message = 'Expected non-null value') {
  if (value === null || value === undefined) {
    throw new Error(message);
  }
  return true;
}

/**
 * Assert that an array has expected length
 */
function assertLength(array, expectedLength, message) {
  if (!Array.isArray(array)) {
    throw new Error('Expected an array');
  }
  if (array.length !== expectedLength) {
    throw new Error(message || `Expected length ${expectedLength} but got ${array.length}`);
  }
  return true;
}

/**
 * Assert that an object has a property
 */
function assertHasProperty(obj, prop, message) {
  if (!(prop in obj)) {
    throw new Error(message || `Expected object to have property ${prop}`);
  }
  return true;
}

/**
 * Assert that a function throws an error
 */
function assertThrows(fn, message = 'Expected function to throw') {
  try {
    fn();
    throw new Error(message);
  } catch (e) {
    if (e.message === message) throw e;
    return true;
  }
}

// ============================================================================
// TEST RESULTS LOGGING
// ============================================================================

/**
 * Log test results to a sheet
 */
function logTestResultsToSheet(reports) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Test Results');

  if (!sheet) {
    sheet = ss.insertSheet('Test Results');
  }

  // Clear and set headers
  sheet.clear();
  sheet.appendRow([
    'Run Date', 'Module', 'Test ID', 'Test Name', 'Status', 'Message', 'Duration (ms)'
  ]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');

  const runDate = new Date().toISOString();

  // Add results
  reports.forEach(report => {
    report.results.forEach(result => {
      sheet.appendRow([
        runDate,
        report.module,
        result.testId,
        result.testName,
        result.status,
        result.message,
        result.duration
      ]);
    });
  });

  // Add summary
  sheet.appendRow([]);
  sheet.appendRow(['=== SUMMARY ===']);
  reports.forEach(report => {
    sheet.appendRow([
      '', report.module,
      `${report.passed}/${report.total} passed`,
      `${report.passRate}%`,
      report.failed > 0 ? `${report.failed} failed` : '',
      report.errors > 0 ? `${report.errors} errors` : ''
    ]);
  });

  // Color code results
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  for (let i = 1; i < values.length; i++) {
    const status = values[i][4];
    if (status === 'passed') {
      sheet.getRange(i + 1, 5).setBackground('#d4edda').setFontColor('#155724');
    } else if (status === 'failed') {
      sheet.getRange(i + 1, 5).setBackground('#f8d7da').setFontColor('#721c24');
    } else if (status === 'error') {
      sheet.getRange(i + 1, 5).setBackground('#fff3cd').setFontColor('#856404');
    }
  }

  sheet.autoResizeColumns(1, 7);
  return sheet.getUrl();
}

/**
 * Generate JSON test report
 */
function generateTestReportJSON(reports) {
  const totalTests = reports.reduce((sum, r) => sum + r.total, 0);
  const totalPassed = reports.reduce((sum, r) => sum + r.passed, 0);
  const totalFailed = reports.reduce((sum, r) => sum + r.failed, 0);
  const totalErrors = reports.reduce((sum, r) => sum + r.errors, 0);

  return {
    timestamp: new Date().toISOString(),
    summary: {
      totalModules: reports.length,
      totalTests,
      totalPassed,
      totalFailed,
      totalErrors,
      overallPassRate: totalTests > 0 ? Math.round((totalPassed / totalTests) * 100) : 0
    },
    modules: reports
  };
}

// ============================================================================
// CLEANUP UTILITIES
// ============================================================================

/**
 * Clean up test sheets (delete sheets starting with "Test_")
 */
function cleanupTestSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let deleted = 0;

  sheets.forEach(sheet => {
    if (sheet.getName().startsWith('Test_')) {
      ss.deleteSheet(sheet);
      deleted++;
    }
  });

  return { deleted, message: `Deleted ${deleted} test sheets` };
}

/**
 * Clean up test data from a sheet
 */
function cleanupTestData(sheetName, testPrefix = 'TEST-') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return { deleted: 0, message: 'Sheet not found' };

  const data = sheet.getDataRange().getValues();
  let deleted = 0;

  // Find rows with test data (check first column for test prefix)
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).startsWith(testPrefix)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }

  return { deleted, message: `Deleted ${deleted} test rows from ${sheetName}` };
}
