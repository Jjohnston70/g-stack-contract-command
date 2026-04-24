/**
 * Function Runner for Contract Tracker
 *
 * Provides a sheet-based interface to run functions safely with logging.
 * Supports installable triggers for checkbox-based execution.
 */

// ============================================================================
// FUNCTION CATALOG - Lists all available functions in this module
// ============================================================================

const FUNCTION_CATALOG = [
  // Initialization & Setup
  { name: 'initializeSheets', file: 'Code.gs', category: 'Setup', description: 'Initialize all required sheets (Contracts, Alerts Log, Settings)', requiresConfirm: true, isDestructive: false },
  { name: 'createTriggers', file: 'Code.gs', category: 'Setup', description: 'Create automatic triggers (daily expirations, weekly renewals)', requiresConfirm: true, isDestructive: false },
  { name: 'removeTriggers', file: 'Code.gs', category: 'Setup', description: 'Remove all project triggers', requiresConfirm: true, isDestructive: true },

  // UI Dialogs
  { name: 'showDashboard', file: 'Code.gs', category: 'UI', description: 'Open the interactive contract dashboard', requiresConfirm: false, isDestructive: false },
  { name: 'showSidebar', file: 'Code.gs', category: 'UI', description: 'Open the tools sidebar', requiresConfirm: false, isDestructive: false },
  { name: 'showAddContractDialog', file: 'Code.gs', category: 'UI', description: 'Open dialog to add a new contract', requiresConfirm: false, isDestructive: false },
  { name: 'showAlertSettings', file: 'Code.gs', category: 'UI', description: 'Show current alert configuration', requiresConfirm: false, isDestructive: false },
  { name: 'showHelp', file: 'Code.gs', category: 'UI', description: 'Show help documentation', requiresConfirm: false, isDestructive: false },
  { name: 'showAbout', file: 'Code.gs', category: 'UI', description: 'Show about dialog', requiresConfirm: false, isDestructive: false },

  // Contract Views & Filters
  { name: 'filterActiveContracts', file: 'Code.gs', category: 'Contracts', description: 'Filter to show only active contracts', requiresConfirm: false, isDestructive: false },
  { name: 'filterExpiringSoon', file: 'Code.gs', category: 'Contracts', description: 'Show contracts expiring in next 30 days', requiresConfirm: false, isDestructive: false },
  { name: 'showAllContracts', file: 'Code.gs', category: 'Contracts', description: 'Remove filters and show all contracts', requiresConfirm: false, isDestructive: false },

  // Alerts & Notifications
  { name: 'checkExpirations', file: 'Code.gs', category: 'Alerts', description: 'Check for expiring contracts and send alerts', requiresConfirm: false, isDestructive: false },
  { name: 'sendRenewalReminders', file: 'Code.gs', category: 'Alerts', description: 'Send renewal reminder emails for upcoming decisions', requiresConfirm: true, isDestructive: false },

  // Reports
  { name: 'generateContractSummary', file: 'Code.gs', category: 'Reports', description: 'Generate overall contract summary report', requiresConfirm: false, isDestructive: false },
  { name: 'generateRenewalReport', file: 'Code.gs', category: 'Reports', description: 'Generate 90-day renewal report', requiresConfirm: false, isDestructive: false },
  { name: 'generateVendorAnalysis', file: 'Code.gs', category: 'Reports', description: 'Analyze contracts by vendor', requiresConfirm: false, isDestructive: false },

  // Data Functions
  { name: 'getContractStats', file: 'Code.gs', category: 'Data', description: 'Get contract statistics summary', requiresConfirm: false, isDestructive: false },
  { name: 'getExpiringContracts', file: 'Code.gs', category: 'Data', description: 'Get list of expiring contracts (30 days)', requiresConfirm: false, isDestructive: false },
  { name: 'getRecentContracts', file: 'Code.gs', category: 'Data', description: 'Get most recently added contracts', requiresConfirm: false, isDestructive: false },
  { name: 'getContractTypes', file: 'Code.gs', category: 'Data', description: 'Get list of contract types in use', requiresConfirm: false, isDestructive: false }
];

// ============================================================================
// FUNCTION RUNNER SHEET MANAGEMENT
// ============================================================================

/**
 * Initialize the Function Runner sheet
 */
function initializeFunctionRunnerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Function Runner');

  if (!sheet) {
    sheet = ss.insertSheet('Function Runner');
  } else {
    sheet.clear();
  }

  // Set up headers
  const headers = ['Run', 'Function Name', 'File', 'Category', 'Description', 'Requires Confirm', 'Last Run', 'Last Result'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white');

  // Populate with function catalog
  const data = FUNCTION_CATALOG.map(fn => [
    false, // Checkbox
    fn.name,
    fn.file,
    fn.category,
    fn.description,
    fn.requiresConfirm ? 'Yes' : 'No',
    '', // Last Run
    ''  // Last Result
  ]);

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);

    // Add checkboxes to Run column
    sheet.getRange(2, 1, data.length, 1).insertCheckboxes();

    // Apply conditional formatting for destructive functions
    FUNCTION_CATALOG.forEach((fn, index) => {
      if (fn.isDestructive) {
        sheet.getRange(index + 2, 2, 1, 5).setBackground('#ffcdd2'); // Light red
      }
    });
  }

  // Format columns
  sheet.setColumnWidth(1, 50);   // Run
  sheet.setColumnWidth(2, 200);  // Function Name
  sheet.setColumnWidth(3, 120);  // File
  sheet.setColumnWidth(4, 100);  // Category
  sheet.setColumnWidth(5, 350);  // Description
  sheet.setColumnWidth(6, 120);  // Requires Confirm
  sheet.setColumnWidth(7, 150);  // Last Run
  sheet.setColumnWidth(8, 200);  // Last Result

  // Freeze header row
  sheet.setFrozenRows(1);

  // Add data validation for category filter
  const categories = [...new Set(FUNCTION_CATALOG.map(fn => fn.category))];

  SpreadsheetApp.getUi().alert(
    'Function Runner Initialized',
    `Created Function Runner sheet with ${FUNCTION_CATALOG.length} functions.\n\n` +
    `Categories: ${categories.join(', ')}\n\n` +
    'Use "Setup Function Runner Trigger" to enable checkbox execution.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return sheet;
}

// ============================================================================
// TRIGGER MANAGEMENT
// ============================================================================

/**
 * Set up the installable trigger for Function Runner
 */
function setupFunctionRunnerTrigger() {
  // Remove existing trigger first
  removeFunctionRunnerTrigger();

  // Create new installable trigger
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFunctionRunnerEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    'Trigger Created',
    'Function Runner trigger is now active.\n\n' +
    'Check any box in the "Run" column to execute that function.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Remove the Function Runner trigger
 */
function removeFunctionRunnerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFunctionRunnerEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Handle edits to the Function Runner sheet
 */
function onFunctionRunnerEdit(e) {
  const sheet = e.range.getSheet();

  // Only process Function Runner sheet
  if (sheet.getName() !== 'Function Runner') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Only process checkbox column (column 1) and not header row
  if (col !== 1 || row < 2) return;

  // Only process if checked (true)
  if (e.value !== 'TRUE') return;

  // Get function name
  const functionName = sheet.getRange(row, 2).getValue();
  const requiresConfirm = sheet.getRange(row, 6).getValue() === 'Yes';

  // Confirm if required
  if (requiresConfirm) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Confirm Execution',
      `Are you sure you want to run "${functionName}"?\n\n` +
      'This action may modify data or settings.',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      // Uncheck the box
      e.range.setValue(false);
      return;
    }
  }

  // Execute the function
  _executeFunctionByName(functionName, sheet, row, e.range);
}

/**
 * Execute a function by name and log the result
 */
function _executeFunctionByName(functionName, sheet, row, checkboxRange) {
  const startTime = new Date();

  try {
    // Get the function reference
    const fn = this[functionName];

    if (typeof fn !== 'function') {
      throw new Error(`Function "${functionName}" not found`);
    }

    // Execute the function
    const result = fn();

    // Log success
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;

    sheet.getRange(row, 7).setValue(startTime.toLocaleString());
    sheet.getRange(row, 8).setValue(`Success (${duration.toFixed(2)}s)`);
    sheet.getRange(row, 8).setBackground('#d4edda');

  } catch (error) {
    // Log error
    sheet.getRange(row, 7).setValue(startTime.toLocaleString());
    sheet.getRange(row, 8).setValue(`Error: ${error.message}`);
    sheet.getRange(row, 8).setBackground('#f8d7da');

  } finally {
    // Uncheck the box
    checkboxRange.setValue(false);
  }
}

/**
 * Run a function from the catalog by name (manual execution)
 */
function runFunctionFromRunner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Function Runner');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Please initialize the Function Runner sheet first.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Run Function',
    'Enter the function name to execute:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const functionName = response.getResponseText().trim();

  // Find the function in catalog
  const fnInfo = FUNCTION_CATALOG.find(f => f.name === functionName);

  if (!fnInfo) {
    ui.alert('Function not found', `"${functionName}" is not in the function catalog.`, ui.ButtonSet.OK);
    return;
  }

  // Confirm if required
  if (fnInfo.requiresConfirm) {
    const confirm = ui.alert(
      'Confirm Execution',
      `Are you sure you want to run "${functionName}"?\n\n${fnInfo.description}`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;
  }

  // Execute
  try {
    const fn = this[functionName];
    if (typeof fn !== 'function') {
      throw new Error('Function not found');
    }

    fn();
    ui.alert('Success', `"${functionName}" executed successfully.`, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', `Failed to execute "${functionName}": ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Get function catalog for external use
 */
function getFunctionCatalog() {
  return FUNCTION_CATALOG;
}
