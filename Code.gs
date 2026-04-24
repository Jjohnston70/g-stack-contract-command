/**
 * True North Contract Tracker
 *
 * Contract and agreement tracking system with renewal alerts,
 * document linking, and comprehensive reporting.
 *
 * Features:
 * - Contract list management (CRUD)
 * - Renewal date tracking
 * - Expiration alerts (email notifications)
 * - Document linking (Drive integration)
 * - Contract status workflow
 * - Interactive dashboard
 */

// ============================================================================
// MENU & INITIALIZATION
// ============================================================================

/**
 * Creates menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(COMPANY_CONFIG.MENU_TITLE)
    .addItem('📊 Open Dashboard', 'showDashboard')
    .addItem('🌐 Open Dashboard in Browser', 'openDashboardInBrowser')
    .addItem('📝 Open Sidebar', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Contracts')
      .addItem('Add New Contract', 'showAddContractDialog')
      .addItem('View Active Contracts', 'filterActiveContracts')
      .addItem('View Expiring Soon', 'filterExpiringSoon')
      .addItem('View All Contracts', 'showAllContracts'))
    .addSubMenu(ui.createMenu('📊 Reports')
      .addItem('Contract Summary', 'generateContractSummary')
      .addItem('Renewal Report', 'generateRenewalReport')
      .addItem('Vendor Analysis', 'generateVendorAnalysis'))
    .addSubMenu(ui.createMenu('🔔 Alerts')
      .addItem('Check Expirations Now', 'checkExpirations')
      .addItem('Send Renewal Reminders', 'sendRenewalReminders')
      .addItem('Configure Alert Settings', 'showAlertSettings'))
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('Initialize Sheets', 'initializeSheets')
      .addItem('Initialize Function Runner', 'initializeFunctionRunnerSheet')
      .addSeparator()
      .addItem('Create Triggers', 'createTriggers')
      .addItem('Remove Triggers', 'removeTriggers')
      .addSeparator()
      .addItem('Setup Function Runner Trigger', 'setupFunctionRunnerTrigger')
      .addItem('Remove Function Runner Trigger', 'removeFunctionRunnerTrigger'))
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

/**
 * Web app entry point.
 */
function doGet(e) {
  try {
    const spreadsheetId = e && e.parameter ? e.parameter.sid : '';
    const ss = getContractSpreadsheet_(spreadsheetId);
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.spreadsheetId = ss.getId();
    return template.evaluate().setTitle('Contract Dashboard');
  } catch (error) {
    return HtmlService.createHtmlOutput('<h3>Contract Tracker Web App Setup Required</h3><p>' + error.message + '</p>');
  }
}

/**
 * Opens the deployed web app URL with spreadsheet context.
 */
function openDashboardInBrowser() {
  const ss = getContractSpreadsheet_();
  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert('Deploy as Web App first, then retry.');
    return;
  }
  const fullUrl = webAppUrl + '?sid=' + encodeURIComponent(ss.getId());
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; padding: 12px;">' +
    '<p><strong>Open browser dashboard:</strong></p>' +
    '<p><a href="' + fullUrl + '" target="_blank">' + fullUrl + '</a></p>' +
    '</div>'
  ).setWidth(700).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'Contract Browser Link');
}

/**
 * Get spreadsheet in both bound-sheet and web-app contexts.
 */
function getContractSpreadsheet_(spreadsheetId) {
  if (spreadsheetId) {
    PropertiesService.getScriptProperties().setProperty('CONTRACT_SPREADSHEET_ID', spreadsheetId);
    return SpreadsheetApp.openById(spreadsheetId);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('CONTRACT_SPREADSHEET_ID', active.getId());
    return active;
  }
  const savedId = PropertiesService.getScriptProperties().getProperty('CONTRACT_SPREADSHEET_ID');
  if (savedId) return SpreadsheetApp.openById(savedId);
  throw new Error('No spreadsheet context found.');
}

/**
 * Initialize all required sheets with headers
 */
function initializeSheets() {
  const ss = getContractSpreadsheet_();
  const ui = SpreadsheetApp.getUi();

  // Contracts sheet
  let contractsSheet = ss.getSheetByName('Contracts');
  if (!contractsSheet) {
    contractsSheet = ss.insertSheet('Contracts');
    contractsSheet.appendRow([
      'Contract ID', 'Contract Name', 'Vendor/Party', 'Type', 'Status',
      'Start Date', 'End Date', 'Renewal Date', 'Value', 'Currency',
      'Auto-Renew', 'Notice Period (Days)', 'Owner', 'Department',
      'Document Link', 'Notes', 'Created Date', 'Last Modified'
    ]);
    contractsSheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    contractsSheet.setFrozenRows(1);
  }

  // Alerts Log sheet
  let alertsSheet = ss.getSheetByName('Alerts Log');
  if (!alertsSheet) {
    alertsSheet = ss.insertSheet('Alerts Log');
    alertsSheet.appendRow(['Date', 'Contract ID', 'Contract Name', 'Alert Type', 'Recipient', 'Status', 'Notes']);
    alertsSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    alertsSheet.setFrozenRows(1);
  }

  // Settings sheet
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Setting', 'Value']);
    settingsSheet.appendRow(['Alert Days Before Expiration', '30,14,7']);
    settingsSheet.appendRow(['Alert Email Recipients', Session.getActiveUser().getEmail()]);
    settingsSheet.appendRow(['Auto-Send Alerts', 'TRUE']);
    settingsSheet.appendRow(['Default Notice Period', '30']);
    settingsSheet.appendRow(['Documents Folder ID', '']);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    settingsSheet.setColumnWidth(1, 250);
    settingsSheet.setColumnWidth(2, 400);
  }

  ui.alert('Setup Complete', 'All sheets have been initialized successfully!', ui.ButtonSet.OK);
}

// ============================================================================
// SIDEBAR
// ============================================================================

/**
 * Shows the sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle(COMPANY_CONFIG.MENU_TITLE)
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get sidebar data
 */
function getSidebarData() {
  const stats = getContractStats();
  return {
    companyName: COMPANY_CONFIG.NAME,
    menuTitle: COMPANY_CONFIG.MENU_TITLE,
    stats: stats,
    expiringContracts: getExpiringContracts(30),
    contractTypes: getContractTypes()
  };
}

// ============================================================================
// DASHBOARD
// ============================================================================

/**
 * Shows the interactive dashboard
 */
function showDashboard() {
  const ss = getContractSpreadsheet_();
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.spreadsheetId = ss.getId();
  const html = template.evaluate()
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Contract Dashboard');
  SpreadsheetApp.getUi().showModalDialog(html, 'Contract Tracker Dashboard');
}

/**
 * Get all dashboard data
 */
function getDashboardData(spreadsheetId) {
  getContractSpreadsheet_(spreadsheetId);
  return {
    companyName: COMPANY_CONFIG.NAME,
    lastUpdated: new Date().toLocaleString(),
    stats: getContractStats(),
    expiringContracts: getExpiringContracts(30),
    recentContracts: getRecentContracts(10),
    charts: getChartData(),
    health: getSystemHealth()
  };
}

/**
 * Get contract statistics
 */
function getContractStats() {
  const stats = {
    total: 0,
    active: 0,
    expiringSoon: 0,
    expired: 0,
    pending: 0,
    totalValue: 0,
    byType: {},
    byStatus: {},
    byDepartment: {}
  };

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return stats;

    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const thirtyDaysOut = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      stats.total++;

      const status = (data[i][4] || 'Unknown').toString();
      const contractType = (data[i][3] || 'Other').toString();
      const department = (data[i][13] || 'Unassigned').toString();
      const endDate = data[i][6] ? new Date(data[i][6]) : null;
      const value = parseFloat(data[i][8]) || 0;

      stats.totalValue += value;

      // By status
      stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;

      // By type
      stats.byType[contractType] = (stats.byType[contractType] || 0) + 1;

      // By department
      stats.byDepartment[department] = (stats.byDepartment[department] || 0) + 1;

      // Status counts
      if (status.toLowerCase() === 'active') {
        stats.active++;

        if (endDate && endDate <= thirtyDaysOut && endDate > now) {
          stats.expiringSoon++;
        }
      } else if (status.toLowerCase() === 'expired') {
        stats.expired++;
      } else if (status.toLowerCase() === 'pending' || status.toLowerCase() === 'draft') {
        stats.pending++;
      }

      // Check for expired active contracts
      if (status.toLowerCase() === 'active' && endDate && endDate < now) {
        stats.expired++;
        stats.active--;
      }
    }
  } catch (e) {
    Logger.log('Error getting contract stats: ' + e.message);
  }

  return stats;
}

/**
 * Get contracts expiring within N days
 */
function getExpiringContracts(days) {
  const expiring = [];

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return expiring;

    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const futureDate = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const status = (data[i][4] || '').toString().toLowerCase();
      const endDate = data[i][6] ? new Date(data[i][6]) : null;

      if (status === 'active' && endDate && endDate >= now && endDate <= futureDate) {
        const daysUntil = Math.ceil((endDate - now) / (1000 * 60 * 60 * 24));

        expiring.push({
          id: data[i][0],
          name: data[i][1],
          vendor: data[i][2],
          type: data[i][3],
          endDate: endDate.toLocaleDateString(),
          daysUntil: daysUntil,
          value: data[i][8],
          autoRenew: data[i][10],
          owner: data[i][12],
          docLink: data[i][14]
        });
      }
    }

    expiring.sort((a, b) => a.daysUntil - b.daysUntil);
  } catch (e) {
    Logger.log('Error getting expiring contracts: ' + e.message);
  }

  return expiring;
}

/**
 * Get recent contracts
 */
function getRecentContracts(limit) {
  const recent = [];

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return recent;

    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1).filter(row => row[0]);

    // Sort by created date (descending)
    rows.sort((a, b) => {
      const dateA = a[16] ? new Date(a[16]) : new Date(0);
      const dateB = b[16] ? new Date(b[16]) : new Date(0);
      return dateB - dateA;
    });

    for (let i = 0; i < Math.min(limit, rows.length); i++) {
      const row = rows[i];
      recent.push({
        id: row[0],
        name: row[1],
        vendor: row[2],
        type: row[3],
        status: row[4],
        startDate: row[5] ? new Date(row[5]).toLocaleDateString() : '',
        endDate: row[6] ? new Date(row[6]).toLocaleDateString() : '',
        value: row[8],
        owner: row[12],
        docLink: row[14]
      });
    }
  } catch (e) {
    Logger.log('Error getting recent contracts: ' + e.message);
  }

  return recent;
}

/**
 * Get chart data for visualizations
 */
function getChartData() {
  const stats = getContractStats();

  const statusEntries = Object.entries(stats.byStatus).sort((a, b) => b[1] - a[1]);
  const typeEntries = Object.entries(stats.byType).sort((a, b) => b[1] - a[1]);
  const deptEntries = Object.entries(stats.byDepartment).sort((a, b) => b[1] - a[1]);

  return {
    statusLabels: statusEntries.map(e => e[0]),
    statusCounts: statusEntries.map(e => e[1]),
    typeLabels: typeEntries.map(e => e[0]),
    typeCounts: typeEntries.map(e => e[1]),
    deptLabels: deptEntries.map(e => e[0]),
    deptCounts: deptEntries.map(e => e[1]),
    expirationTimeline: getExpirationTimeline()
  };
}

/**
 * Get expiration timeline for next 6 months
 */
function getExpirationTimeline() {
  const timeline = [];
  const now = new Date();

  for (let i = 0; i < 6; i++) {
    const monthStart = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + i + 1, 0);

    const monthLabel = monthStart.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });

    timeline.push({
      month: monthLabel,
      count: countContractsExpiringInRange(monthStart, monthEnd)
    });
  }

  return timeline;
}

/**
 * Count contracts expiring in a date range
 */
function countContractsExpiringInRange(startDate, endDate) {
  let count = 0;

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return count;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const status = (data[i][4] || '').toString().toLowerCase();
      const endDateCell = data[i][6] ? new Date(data[i][6]) : null;

      if (status === 'active' && endDateCell && endDateCell >= startDate && endDateCell <= endDate) {
        count++;
      }
    }
  } catch (e) {
    Logger.log('Error counting contracts: ' + e.message);
  }

  return count;
}

/**
 * Get system health status
 */
function getSystemHealth() {
  const health = {
    sheets: { status: 'ok', message: 'All sheets present' },
    triggers: { status: 'ok', message: 'Triggers configured' },
    settings: { status: 'ok', message: 'Settings valid' }
  };

  const ss = getContractSpreadsheet_();

  // Check sheets
  const requiredSheets = ['Contracts', 'Alerts Log', 'Settings'];
  const missingSheets = requiredSheets.filter(name => !ss.getSheetByName(name));
  if (missingSheets.length > 0) {
    health.sheets = { status: 'warning', message: `Missing: ${missingSheets.join(', ')}` };
  }

  // Check triggers
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    health.triggers = { status: 'warning', message: 'No triggers configured' };
  }

  // Check settings
  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet) {
    const alertEmail = getSetting('Alert Email Recipients');
    if (!alertEmail) {
      health.settings = { status: 'warning', message: 'Alert email not configured' };
    }
  }

  return health;
}

// ============================================================================
// CONTRACT CRUD OPERATIONS
// ============================================================================

/**
 * Shows add contract dialog
 */
function showAddContractDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddContract')
    .setWidth(600)
    .setHeight(700)
    .setTitle('Add New Contract');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Contract');
}

/**
 * Add a new contract
 */
function addContract(contractData) {
  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) {
      return { success: false, message: 'Contracts sheet not found. Please run initialization.' };
    }

    const contractId = 'CTR-' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd') + '-' +
                       Math.random().toString(36).substring(2, 8).toUpperCase();

    const now = new Date();

    sheet.appendRow([
      contractId,
      contractData.name,
      contractData.vendor,
      contractData.type,
      contractData.status || 'Active',
      contractData.startDate,
      contractData.endDate,
      contractData.renewalDate || contractData.endDate,
      contractData.value || 0,
      contractData.currency || 'USD',
      contractData.autoRenew || false,
      contractData.noticePeriod || 30,
      contractData.owner || Session.getActiveUser().getEmail(),
      contractData.department,
      contractData.docLink || '',
      contractData.notes || '',
      now,
      now
    ]);

    return { success: true, contractId: contractId, message: 'Contract added successfully!' };
  } catch (e) {
    Logger.log('Error adding contract: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Update a contract
 */
function updateContract(contractId, updates) {
  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) {
      return { success: false, message: 'Contracts sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === contractId) {
        const rowNum = i + 1;

        if (updates.name !== undefined) sheet.getRange(rowNum, 2).setValue(updates.name);
        if (updates.vendor !== undefined) sheet.getRange(rowNum, 3).setValue(updates.vendor);
        if (updates.type !== undefined) sheet.getRange(rowNum, 4).setValue(updates.type);
        if (updates.status !== undefined) sheet.getRange(rowNum, 5).setValue(updates.status);
        if (updates.startDate !== undefined) sheet.getRange(rowNum, 6).setValue(updates.startDate);
        if (updates.endDate !== undefined) sheet.getRange(rowNum, 7).setValue(updates.endDate);
        if (updates.renewalDate !== undefined) sheet.getRange(rowNum, 8).setValue(updates.renewalDate);
        if (updates.value !== undefined) sheet.getRange(rowNum, 9).setValue(updates.value);
        if (updates.autoRenew !== undefined) sheet.getRange(rowNum, 11).setValue(updates.autoRenew);
        if (updates.noticePeriod !== undefined) sheet.getRange(rowNum, 12).setValue(updates.noticePeriod);
        if (updates.owner !== undefined) sheet.getRange(rowNum, 13).setValue(updates.owner);
        if (updates.department !== undefined) sheet.getRange(rowNum, 14).setValue(updates.department);
        if (updates.docLink !== undefined) sheet.getRange(rowNum, 15).setValue(updates.docLink);
        if (updates.notes !== undefined) sheet.getRange(rowNum, 16).setValue(updates.notes);

        // Update last modified
        sheet.getRange(rowNum, 18).setValue(new Date());

        return { success: true, message: 'Contract updated successfully!' };
      }
    }

    return { success: false, message: 'Contract not found.' };
  } catch (e) {
    Logger.log('Error updating contract: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Delete a contract
 */
function deleteContract(contractId) {
  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) {
      return { success: false, message: 'Contracts sheet not found.' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === contractId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Contract deleted successfully!' };
      }
    }

    return { success: false, message: 'Contract not found.' };
  } catch (e) {
    Logger.log('Error deleting contract: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Get contract by ID
 */
function getContract(contractId) {
  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === contractId) {
        return {
          id: data[i][0],
          name: data[i][1],
          vendor: data[i][2],
          type: data[i][3],
          status: data[i][4],
          startDate: data[i][5],
          endDate: data[i][6],
          renewalDate: data[i][7],
          value: data[i][8],
          currency: data[i][9],
          autoRenew: data[i][10],
          noticePeriod: data[i][11],
          owner: data[i][12],
          department: data[i][13],
          docLink: data[i][14],
          notes: data[i][15],
          createdDate: data[i][16],
          lastModified: data[i][17]
        };
      }
    }

    return null;
  } catch (e) {
    Logger.log('Error getting contract: ' + e.message);
    return null;
  }
}

/**
 * Get contract types from data
 */
function getContractTypes() {
  const types = new Set();

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return ['Service', 'Subscription', 'License', 'Lease', 'Employment', 'NDA', 'Other'];

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][3]) {
        types.add(data[i][3].toString());
      }
    }
  } catch (e) {
    Logger.log('Error getting contract types: ' + e.message);
  }

  const defaultTypes = ['Service', 'Subscription', 'License', 'Lease', 'Employment', 'NDA', 'Other'];
  return [...new Set([...types, ...defaultTypes])].sort();
}

// ============================================================================
// FILTERING & VIEWS
// ============================================================================

/**
 * Filter to show only active contracts
 */
function filterActiveContracts() {
  const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Please initialize sheets first.');
    return;
  }

  // Remove existing filter
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  const range = sheet.getDataRange();
  const filter = range.createFilter();

  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo('Active')
    .build();

  filter.setColumnFilterCriteria(5, criteria);

  SpreadsheetApp.getUi().alert('Filter Applied', 'Showing active contracts only. Use Data > Remove Filter to see all.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Filter to show contracts expiring soon
 */
function filterExpiringSoon() {
  const expiring = getExpiringContracts(30);
  const ui = SpreadsheetApp.getUi();

  if (expiring.length === 0) {
    ui.alert('No Expiring Contracts', 'No contracts are expiring in the next 30 days.', ui.ButtonSet.OK);
    return;
  }

  const message = expiring.slice(0, 10).map(c =>
    `• ${c.name} (${c.vendor}) - ${c.daysUntil} days`
  ).join('\n');

  ui.alert(
    `Expiring Soon (${expiring.length} contracts)`,
    message + (expiring.length > 10 ? `\n\n...and ${expiring.length - 10} more` : ''),
    ui.ButtonSet.OK
  );
}

/**
 * Show all contracts (remove filter)
 */
function showAllContracts() {
  const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
  if (sheet && sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  SpreadsheetApp.getUi().alert('All contracts are now visible.');
}

// ============================================================================
// ALERTS & NOTIFICATIONS
// ============================================================================

/**
 * Check for expiring contracts and send alerts
 */
function checkExpirations() {
  const settings = getSettings();
  const alertDays = (settings['Alert Days Before Expiration'] || '30,14,7')
    .split(',')
    .map(d => parseInt(d.trim()));

  const expiring = [];
  const now = new Date();

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const status = (data[i][4] || '').toString().toLowerCase();
      const endDate = data[i][6] ? new Date(data[i][6]) : null;

      if (status === 'active' && endDate) {
        const daysUntil = Math.ceil((endDate - now) / (1000 * 60 * 60 * 24));

        if (alertDays.includes(daysUntil) || daysUntil <= 0) {
          expiring.push({
            id: data[i][0],
            name: data[i][1],
            vendor: data[i][2],
            endDate: endDate,
            daysUntil: daysUntil,
            owner: data[i][12],
            value: data[i][8]
          });
        }
      }
    }
  } catch (e) {
    Logger.log('Error checking expirations: ' + e.message);
  }

  if (expiring.length > 0) {
    const autoSend = settings['Auto-Send Alerts'] === 'TRUE';

    if (autoSend) {
      sendExpirationAlerts(expiring);
    }

    const ui = SpreadsheetApp.getUi();
    const message = expiring.map(c =>
      `• ${c.name}: ${c.daysUntil <= 0 ? 'EXPIRED' : c.daysUntil + ' days'}`
    ).join('\n');

    ui.alert(
      `Found ${expiring.length} Contract(s) Requiring Attention`,
      message + (autoSend ? '\n\nAlerts have been sent.' : '\n\nAuto-send is disabled.'),
      ui.ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('No contracts require attention at this time.');
  }
}

/**
 * Send expiration alert emails
 */
function sendExpirationAlerts(contracts) {
  const settings = getSettings();
  const recipients = settings['Alert Email Recipients'];

  if (!recipients) {
    Logger.log('No alert recipients configured');
    return;
  }

  const subject = `${COMPANY_CONFIG.NAME} - Contract Expiration Alert`;

  const expiredContracts = contracts.filter(c => c.daysUntil <= 0);
  const expiringContracts = contracts.filter(c => c.daysUntil > 0);

  let body = `<div style="font-family: Arial, sans-serif; max-width: 600px;">`;
  body += `<h2 style="color: #1a73e8;">Contract Expiration Alert</h2>`;
  body += `<p>The following contracts require your attention:</p>`;

  if (expiredContracts.length > 0) {
    body += `<h3 style="color: #d93025;">⚠️ Expired Contracts (${expiredContracts.length})</h3>`;
    body += `<ul>`;
    expiredContracts.forEach(c => {
      body += `<li><strong>${c.name}</strong> (${c.vendor}) - Expired ${Math.abs(c.daysUntil)} days ago</li>`;
    });
    body += `</ul>`;
  }

  if (expiringContracts.length > 0) {
    body += `<h3 style="color: #f9ab00;">📅 Expiring Soon (${expiringContracts.length})</h3>`;
    body += `<ul>`;
    expiringContracts.forEach(c => {
      body += `<li><strong>${c.name}</strong> (${c.vendor}) - ${c.daysUntil} days remaining</li>`;
    });
    body += `</ul>`;
  }

  body += `<p style="margin-top: 20px;"><a href="${getContractSpreadsheet_().getUrl()}" style="background: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">View Contract Tracker</a></p>`;
  body += `<hr style="margin-top: 30px;">`;
  body += `<p style="color: #666; font-size: 12px;">${COMPANY_CONFIG.NAME} Contract Tracker</p>`;
  body += `</div>`;

  try {
    GmailApp.sendEmail(recipients, subject, '', { htmlBody: body });
    logAlert(contracts, 'Expiration Alert', recipients, 'Sent');
  } catch (e) {
    Logger.log('Error sending alert email: ' + e.message);
    logAlert(contracts, 'Expiration Alert', recipients, 'Failed: ' + e.message);
  }
}

/**
 * Send renewal reminder emails
 */
function sendRenewalReminders() {
  const settings = getSettings();
  const defaultNoticePeriod = parseInt(settings['Default Notice Period']) || 30;

  const needsRenewal = [];
  const now = new Date();

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const status = (data[i][4] || '').toString().toLowerCase();
      const renewalDate = data[i][7] ? new Date(data[i][7]) : null;
      const noticePeriod = data[i][11] || defaultNoticePeriod;
      const autoRenew = data[i][10];

      if (status === 'active' && renewalDate && !autoRenew) {
        const noticeDate = new Date(renewalDate.getTime() - noticePeriod * 24 * 60 * 60 * 1000);
        const daysUntilNotice = Math.ceil((noticeDate - now) / (1000 * 60 * 60 * 24));

        if (daysUntilNotice <= 7 && daysUntilNotice >= 0) {
          needsRenewal.push({
            id: data[i][0],
            name: data[i][1],
            vendor: data[i][2],
            renewalDate: renewalDate,
            noticePeriod: noticePeriod,
            daysUntilNotice: daysUntilNotice,
            owner: data[i][12]
          });
        }
      }
    }
  } catch (e) {
    Logger.log('Error checking renewal reminders: ' + e.message);
  }

  if (needsRenewal.length > 0) {
    sendRenewalAlerts(needsRenewal);
    SpreadsheetApp.getUi().alert(
      'Renewal Reminders Sent',
      `Sent reminders for ${needsRenewal.length} contract(s).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('No renewal reminders needed at this time.');
  }
}

/**
 * Send renewal alert emails
 */
function sendRenewalAlerts(contracts) {
  const settings = getSettings();
  const recipients = settings['Alert Email Recipients'];

  if (!recipients) return;

  const subject = `${COMPANY_CONFIG.NAME} - Contract Renewal Action Required`;

  let body = `<div style="font-family: Arial, sans-serif; max-width: 600px;">`;
  body += `<h2 style="color: #1a73e8;">Renewal Action Required</h2>`;
  body += `<p>The following contracts have upcoming renewal decision deadlines:</p>`;
  body += `<ul>`;

  contracts.forEach(c => {
    body += `<li><strong>${c.name}</strong> (${c.vendor})`;
    body += `<br>Renewal Date: ${c.renewalDate.toLocaleDateString()}`;
    body += `<br>Notice Period: ${c.noticePeriod} days`;
    body += `<br>Days until notice deadline: <strong>${c.daysUntilNotice}</strong></li>`;
  });

  body += `</ul>`;
  body += `<p style="margin-top: 20px;"><a href="${getContractSpreadsheet_().getUrl()}" style="background: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">View Contract Tracker</a></p>`;
  body += `</div>`;

  try {
    GmailApp.sendEmail(recipients, subject, '', { htmlBody: body });
    logAlert(contracts, 'Renewal Reminder', recipients, 'Sent');
  } catch (e) {
    Logger.log('Error sending renewal email: ' + e.message);
  }
}

/**
 * Log an alert to the Alerts Log sheet
 */
function logAlert(contracts, alertType, recipient, status) {
  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Alerts Log');
    if (!sheet) return;

    const now = new Date();
    contracts.forEach(c => {
      sheet.appendRow([now, c.id, c.name, alertType, recipient, status, '']);
    });
  } catch (e) {
    Logger.log('Error logging alert: ' + e.message);
  }
}

/**
 * Show alert settings dialog
 */
function showAlertSettings() {
  const settings = getSettings();
  const ui = SpreadsheetApp.getUi();

  const message = `Current Alert Settings:

• Days before expiration: ${settings['Alert Days Before Expiration'] || '30,14,7'}
• Recipients: ${settings['Alert Email Recipients'] || 'Not set'}
• Auto-send: ${settings['Auto-Send Alerts'] || 'FALSE'}
• Default notice period: ${settings['Default Notice Period'] || '30'} days

Edit settings in the Settings sheet.`;

  ui.alert('Alert Settings', message, ui.ButtonSet.OK);
}

// ============================================================================
// REPORTS
// ============================================================================

/**
 * Generate contract summary report
 */
function generateContractSummary() {
  const stats = getContractStats();
  const ui = SpreadsheetApp.getUi();

  const report = `CONTRACT SUMMARY REPORT
========================
Generated: ${new Date().toLocaleString()}

OVERVIEW:
• Total Contracts: ${stats.total}
• Active: ${stats.active}
• Expiring (30 days): ${stats.expiringSoon}
• Expired: ${stats.expired}
• Pending/Draft: ${stats.pending}

TOTAL VALUE: $${stats.totalValue.toLocaleString()}

BY STATUS:
${Object.entries(stats.byStatus).map(([k, v]) => `• ${k}: ${v}`).join('\n')}

BY TYPE:
${Object.entries(stats.byType).map(([k, v]) => `• ${k}: ${v}`).join('\n')}

BY DEPARTMENT:
${Object.entries(stats.byDepartment).map(([k, v]) => `• ${k}: ${v}`).join('\n')}`;

  ui.alert('Contract Summary', report, ui.ButtonSet.OK);
}

/**
 * Generate renewal report
 */
function generateRenewalReport() {
  const expiring = getExpiringContracts(90);
  const ui = SpreadsheetApp.getUi();

  if (expiring.length === 0) {
    ui.alert('Renewal Report', 'No contracts expiring in the next 90 days.', ui.ButtonSet.OK);
    return;
  }

  let report = `RENEWAL REPORT (Next 90 Days)
==============================
Generated: ${new Date().toLocaleString()}
Contracts Found: ${expiring.length}

`;

  expiring.forEach(c => {
    report += `• ${c.name}
  Vendor: ${c.vendor}
  Expires: ${c.endDate} (${c.daysUntil} days)
  Value: ${c.value ? '$' + parseFloat(c.value).toLocaleString() : 'N/A'}
  Auto-Renew: ${c.autoRenew ? 'Yes' : 'No'}

`;
  });

  ui.alert('Renewal Report', report, ui.ButtonSet.OK);
}

/**
 * Generate vendor analysis
 */
function generateVendorAnalysis() {
  const vendors = {};

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Contracts');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Please initialize sheets first.');
      return;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const vendor = (data[i][2] || 'Unknown').toString();
      const value = parseFloat(data[i][8]) || 0;
      const status = (data[i][4] || '').toString().toLowerCase();

      if (!vendors[vendor]) {
        vendors[vendor] = { count: 0, value: 0, active: 0 };
      }

      vendors[vendor].count++;
      vendors[vendor].value += value;
      if (status === 'active') vendors[vendor].active++;
    }
  } catch (e) {
    Logger.log('Error analyzing vendors: ' + e.message);
  }

  const sorted = Object.entries(vendors)
    .sort((a, b) => b[1].value - a[1].value)
    .slice(0, 15);

  let report = `VENDOR ANALYSIS
================
Generated: ${new Date().toLocaleString()}
Unique Vendors: ${Object.keys(vendors).length}

TOP VENDORS BY VALUE:
`;

  sorted.forEach(([vendor, data], i) => {
    report += `${i + 1}. ${vendor}
   Contracts: ${data.count} (${data.active} active)
   Total Value: $${data.value.toLocaleString()}

`;
  });

  SpreadsheetApp.getUi().alert('Vendor Analysis', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// DOCUMENT MANAGEMENT
// ============================================================================

/**
 * Link a document to a contract
 */
function linkDocument(contractId, documentUrl) {
  return updateContract(contractId, { docLink: documentUrl });
}

/**
 * Open linked document
 */
function openLinkedDocument(contractId) {
  const contract = getContract(contractId);
  if (contract && contract.docLink) {
    const html = HtmlService.createHtmlOutput(
      `<script>window.open('${contract.docLink}', '_blank'); google.script.host.close();</script>`
    );
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening Document...');
  } else {
    SpreadsheetApp.getUi().alert('No document linked to this contract.');
  }
}

// ============================================================================
// SETTINGS & CONFIGURATION
// ============================================================================

/**
 * Get all settings as object
 */
function getSettings() {
  const settings = {};

  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Settings');
    if (!sheet) return settings;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        settings[data[i][0]] = data[i][1];
      }
    }
  } catch (e) {
    Logger.log('Error getting settings: ' + e.message);
  }

  return settings;
}

/**
 * Get a single setting value
 */
function getSetting(key) {
  const settings = getSettings();
  return settings[key];
}

/**
 * Update a setting
 */
function updateSetting(key, value) {
  try {
    const sheet = getContractSpreadsheet_().getSheetByName('Settings');
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return true;
      }
    }

    // Add new setting if not found
    sheet.appendRow([key, value]);
    return true;
  } catch (e) {
    Logger.log('Error updating setting: ' + e.message);
    return false;
  }
}

// ============================================================================
// TRIGGERS
// ============================================================================

/**
 * Create automatic triggers
 */
function createTriggers() {
  // Remove existing triggers first
  removeTriggers();

  // Daily expiration check at 8 AM
  ScriptApp.newTrigger('checkExpirations')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Weekly renewal reminders on Monday at 9 AM
  ScriptApp.newTrigger('sendRenewalReminders')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  SpreadsheetApp.getUi().alert(
    'Triggers Created',
    'Daily expiration checks (8 AM) and weekly renewal reminders (Monday 9 AM) have been configured.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Remove all triggers
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

// ============================================================================
// HELP & ABOUT
// ============================================================================

/**
 * Shows help dialog
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Help - ' + COMPANY_CONFIG.MENU_TITLE,
    `CONTRACT TRACKER HELP
======================

GETTING STARTED:
1. Run Setup > Initialize Sheets
2. Add your first contract via Contracts > Add New
3. Configure alerts in the Settings sheet
4. Create triggers for automated alerts

KEY FEATURES:
• Track contract lifecycle from draft to expiration
• Get automatic email alerts before expiration
• Link documents from Google Drive
• Generate reports for analysis

SHEET STRUCTURE:
• Contracts - All contract records
• Alerts Log - History of sent alerts
• Settings - Configuration options

SUPPORT:
${COMPANY_CONFIG.TRUENORTH_EMAIL}`,
    ui.ButtonSet.OK
  );
}

/**
 * Shows about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About ' + COMPANY_CONFIG.MENU_TITLE,
    `Contract Tracker
Version 1.0.0

Track contracts, renewals, and expirations with automated alerts.

${COMPANY_CONFIG.NAME}
${COMPANY_CONFIG.EMAIL || ''}

Support: ${COMPANY_CONFIG.TRUENORTH_EMAIL}`,
    ui.ButtonSet.OK
  );
}

/**
 * Execute dashboard actions
 */
function executeDashboardAction(action, params, spreadsheetId) {
  getContractSpreadsheet_(spreadsheetId);
  switch(action) {
    case 'refresh':
      return getDashboardData();
    case 'updateStatus':
      return updateContract(params.contractId, { status: params.status });
    case 'addContract':
      return addContract(params);
    case 'deleteContract':
      return deleteContract(params.contractId);
    default:
      return { success: false, message: 'Unknown action' };
  }
}

