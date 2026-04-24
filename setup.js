#!/usr/bin/env node
/**
 * Contract Tracker - Client Setup Script
 * Creates a standalone Google Apps Script project for a client
 *
 * Usage: node setup.js
 *
 * True North Data Strategies
 * jacob@truenorthstrategyops.com
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { execSync } = require('child_process');

// Industry Templates - contract types, tracking fields, and renewal patterns for different business types
const INDUSTRY_TEMPLATES = {
  'consulting': {
    name: 'Consulting / Agency',
    emoji: '💼',
    contractTypes: ['Master Service Agreement', 'Statement of Work', 'Retainer Agreement', 'NDA', 'Subcontractor'],
    trackingFields: ['Client Name', 'Project', 'Value', 'Start Date', 'End Date', 'Auto-Renew'],
    renewalPeriods: ['Monthly', 'Quarterly', 'Annual', 'Project-based']
  },
  'real-estate': {
    name: 'Real Estate',
    emoji: '🏠',
    contractTypes: ['Listing Agreement', 'Purchase Agreement', 'Lease', 'Property Management', 'Referral Agreement'],
    trackingFields: ['Property Address', 'Client', 'Commission', 'List Date', 'Expiry Date', 'Status'],
    renewalPeriods: ['6 Months', '12 Months', '24 Months', 'Month-to-Month']
  },
  'healthcare': {
    name: 'Healthcare / Medical',
    emoji: '🏥',
    contractTypes: ['Provider Agreement', 'Payer Contract', 'Vendor Agreement', 'BAA', 'Employment Contract'],
    trackingFields: ['Provider/Vendor', 'Type', 'Effective Date', 'Term End', 'Compliance Status', 'Value'],
    renewalPeriods: ['Annual', 'Bi-Annual', 'Multi-Year', 'Evergreen']
  },
  'construction': {
    name: 'Construction / Trades',
    emoji: '🔧',
    contractTypes: ['General Contract', 'Subcontract', 'Change Order', 'Warranty', 'Maintenance Agreement'],
    trackingFields: ['Job Name', 'Contractor', 'Contract Value', 'Start Date', 'Completion Date', 'Retainage'],
    renewalPeriods: ['Project Duration', 'Annual', 'Warranty Period', 'Ongoing']
  },
  'technology': {
    name: 'Technology / SaaS',
    emoji: '💻',
    contractTypes: ['SaaS Agreement', 'License Agreement', 'SLA', 'Partner Agreement', 'Reseller Agreement'],
    trackingFields: ['Customer', 'Product/Service', 'MRR/ARR', 'Start Date', 'Renewal Date', 'Tier'],
    renewalPeriods: ['Monthly', 'Annual', 'Multi-Year', 'Auto-Renew']
  },
  'finance': {
    name: 'Finance / Accounting',
    emoji: '💰',
    contractTypes: ['Engagement Letter', 'Audit Contract', 'Advisory Agreement', 'Tax Prep Agreement', 'Confidentiality'],
    trackingFields: ['Client', 'Service Type', 'Fee Structure', 'Engagement Date', 'Deadline', 'Status'],
    renewalPeriods: ['Annual', 'Tax Year', 'Quarterly', 'Project-based']
  },
  'legal': {
    name: 'Legal / Law Firm',
    emoji: '⚖️',
    contractTypes: ['Retainer Agreement', 'Engagement Letter', 'Fee Agreement', 'Referral Agreement', 'Co-Counsel'],
    trackingFields: ['Client/Matter', 'Attorney', 'Fee Type', 'Retainer Amount', 'Start Date', 'Status'],
    renewalPeriods: ['Annual', 'Matter Duration', 'Evergreen', 'Monthly Refresh']
  },
  'insurance': {
    name: 'Insurance',
    emoji: '🛡️',
    contractTypes: ['Policy', 'Binder', 'Agency Agreement', 'Carrier Contract', 'Claims Agreement'],
    trackingFields: ['Insured/Agent', 'Policy Type', 'Premium', 'Effective Date', 'Expiration', 'Carrier'],
    renewalPeriods: ['Annual', 'Semi-Annual', 'Quarterly', 'Monthly']
  },
  'general': {
    name: 'General Business',
    emoji: '🏢',
    contractTypes: ['Service Agreement', 'Vendor Contract', 'Employment', 'NDA', 'Partnership'],
    trackingFields: ['Party Name', 'Type', 'Value', 'Start Date', 'End Date', 'Status'],
    renewalPeriods: ['Monthly', 'Annual', 'Multi-Year', 'Ongoing']
  }
};

// Emoji to industry default mapping
const EMOJI_INDUSTRY_MAP = {
  '💼': 'consulting',
  '🏠': 'real-estate',
  '🏥': 'healthcare',
  '🔧': 'construction',
  '💻': 'technology',
  '💰': 'finance',
  '⚖️': 'legal',
  '🛡️': 'insurance',
  '🏢': 'general',
  '📝': 'general'
};

// Module Configuration
const MODULE_CONFIG = {
  name: 'Contract Tracker',
  description: 'Contract and agreement tracking system',
  projectSuffix: 'Contract-Tracker',
  projectType: 'sheets',
  defaultIcon: '📝',

  // Files to include in deployment
  templateFiles: [
    'config.gs',
    'Code.gs',
    'FunctionRunner.gs',
    'appsscript.json'
  ],

  // File rename mappings (source -> destination pattern)
  fileRenames: {},

  // Legacy string replacements (for backward compatibility)
  legacyReplacements: {},

  // Next steps displayed after setup
  nextSteps: [
    'Open the Google Sheet that was created',
    'Refresh the page to see the custom menu',
    'Click the custom menu in the toolbar',
    'Grant permissions when prompted',
    'Add your contracts to track renewals and expirations'
  ]
};

// =========================================
// Readline Utilities
// =========================================

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function ask(question) {
  return new Promise(resolve => rl.question(question, resolve));
}

// =========================================
// Main Setup Function
// =========================================

async function main() {
  console.log('\n========================================');
  console.log(`   ${MODULE_CONFIG.name} - Client Setup`);
  console.log(`   ${MODULE_CONFIG.description}`);
  console.log('========================================\n');
  console.log('This wizard will create a customized Google Apps Script project');
  console.log('with your company branding and deploy it to Google Workspace.\n');

  // Gather client information
  const answers = await collectCompanyInfo();

  // Display summary
  displaySummary(answers);

  // Confirm
  const confirm = await ask('\nProceed with setup? (y/n): ');
  if (confirm.toLowerCase() !== 'y') {
    console.log('Setup cancelled.');
    rl.close();
    process.exit(0);
  }

  // Create output directory
  const safeName = createSafeName(answers.companyName);
  const outputDir = await createOutputDirectory(safeName);

  // Process template files
  console.log('\nProcessing template files...');
  const replacements = buildReplacementMap(answers, safeName);
  processFiles(outputDir, replacements);

  // Create .clasp.json.template for reference
  createClaspTemplate(outputDir);

  // Close readline before clasp operations
  rl.close();

  // Deploy with clasp
  const projectTitle = `${answers.shortName} ${MODULE_CONFIG.projectSuffix}`;
  const claspJson = deployWithClasp(outputDir, projectTitle);

  // Display completion message
  displayComplete(claspJson, answers);
}

// =========================================
// Company Info Collection
// =========================================

async function collectCompanyInfo() {
  const answers = {};

  // Required: Company Name
  answers.companyName = await ask('Company Name (e.g., "Acme Corp"): ');
  if (!answers.companyName.trim()) {
    console.log('Error: Company name is required.');
    rl.close();
    process.exit(1);
  }
  answers.companyName = answers.companyName.trim();

  // Optional: Short Name (default to company name)
  const shortNameInput = await ask(`Short Name for menus [${answers.companyName}]: `);
  answers.shortName = shortNameInput.trim() || answers.companyName;

  // Optional: Menu Icon
  const iconInput = await ask(`Menu Icon emoji [${MODULE_CONFIG.defaultIcon}]: `);
  answers.menuIcon = iconInput.trim() || MODULE_CONFIG.defaultIcon;

  // Industry Selection
  const suggestedIndustry = EMOJI_INDUSTRY_MAP[answers.menuIcon] || 'general';
  console.log('\n--- Industry Contract Templates ---');
  const industries = Object.keys(INDUSTRY_TEMPLATES);
  industries.forEach((key, i) => {
    const ind = INDUSTRY_TEMPLATES[key];
    const marker = key === suggestedIndustry ? ' (suggested)' : '';
    console.log(`  ${i + 1}. ${ind.emoji} ${ind.name}${marker}`);
  });
  const industryInput = await ask(`\nSelect industry (1-${industries.length}) [${industries.indexOf(suggestedIndustry) + 1}]: `);
  const industryIndex = parseInt(industryInput.trim()) - 1;
  if (industryIndex >= 0 && industryIndex < industries.length) {
    answers.industry = industries[industryIndex];
  } else {
    answers.industry = suggestedIndustry;
  }
  answers.industryData = INDUSTRY_TEMPLATES[answers.industry];

  // Optional: Email
  answers.companyEmail = (await ask('\nCompany Email (optional): ')).trim();

  // Optional: Phone
  answers.companyPhone = (await ask('Company Phone (optional): ')).trim();

  return answers;
}

// =========================================
// Display Functions
// =========================================

function displaySummary(answers) {
  console.log('\n========================================');
  console.log('         CONFIGURATION SUMMARY');
  console.log('========================================\n');

  console.log('COMPANY INFORMATION:');
  console.log(`  Company Name: ${answers.companyName}`);
  console.log(`  Short Name:   ${answers.shortName}`);
  console.log(`  Menu Icon:    ${answers.menuIcon}`);
  console.log(`  Email:        ${answers.companyEmail || '(not set)'}`);
  console.log(`  Phone:        ${answers.companyPhone || '(not set)'}`);

  console.log('\nINDUSTRY TEMPLATE:');
  console.log(`  ${answers.industryData.emoji} ${answers.industryData.name}`);

  console.log('\nCONTRACT TYPES:');
  answers.industryData.contractTypes.forEach(type => {
    console.log(`  📄 ${type}`);
  });

  console.log('\nTRACKING FIELDS:');
  answers.industryData.trackingFields.forEach(field => {
    console.log(`  📋 ${field}`);
  });

  console.log('\nRENEWAL PERIODS:');
  answers.industryData.renewalPeriods.forEach(period => {
    console.log(`  🔄 ${period}`);
  });

  console.log('\n========================================');
}

function displayComplete(claspJson, answers) {
  console.log('\n========================================');
  console.log('           SETUP COMPLETE!');
  console.log('========================================\n');
  console.log(`${MODULE_CONFIG.name} created successfully!\n`);

  if (claspJson && claspJson.scriptId) {
    console.log('Script Editor:');
    console.log(`  https://script.google.com/d/${claspJson.scriptId}/edit\n`);
  }

  console.log('Next Steps:');
  MODULE_CONFIG.nextSteps.forEach((step, i) => {
    let displayStep = step;
    if (step.includes('custom menu')) {
      displayStep = `Click "${answers.menuIcon} ${answers.shortName} Contracts" menu`;
    }
    console.log(`  ${i + 1}. ${displayStep}`);
  });
  console.log('');

  console.log('Support: jacob@truenorthstrategyops.com\n');
}

// =========================================
// File Processing
// =========================================

function createSafeName(companyName) {
  return companyName
    .replace(/[^a-zA-Z0-9]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

async function createOutputDirectory(safeName) {
  const dirName = `${safeName}-${MODULE_CONFIG.projectSuffix}`;
  const outputDir = path.join(__dirname, '..', '..', dirName);

  if (fs.existsSync(outputDir)) {
    const overwrite = await ask(`\nDirectory "${dirName}" exists. Overwrite? (y/n): `);
    if (overwrite.toLowerCase() !== 'y') {
      console.log('Setup cancelled.');
      rl.close();
      process.exit(0);
    }
    fs.rmSync(outputDir, { recursive: true });
  }

  console.log(`\nCreating project in: ${outputDir}`);
  fs.mkdirSync(outputDir, { recursive: true });

  return outputDir;
}

function buildReplacementMap(answers, safeName) {
  const replacements = {
    // Standard placeholders
    '{{COMPANY_NAME}}': answers.companyName,
    '{{SHORT_NAME}}': answers.shortName,
    '{{MENU_ICON}}': answers.menuIcon,
    '{{COMPANY_EMAIL}}': answers.companyEmail || '',
    '{{COMPANY_PHONE}}': answers.companyPhone || '',
    '{{SAFE_NAME}}': safeName,
    '{{TRUENORTH_EMAIL}}': 'jacob@truenorthstrategyops.com',

    // Industry placeholders
    '{{INDUSTRY_KEY}}': answers.industry,
    '{{INDUSTRY_NAME}}': answers.industryData.name,
    '{{CONTRACT_TYPES}}': JSON.stringify(answers.industryData.contractTypes),
    '{{TRACKING_FIELDS}}': JSON.stringify(answers.industryData.trackingFields),
    '{{RENEWAL_PERIODS}}': JSON.stringify(answers.industryData.renewalPeriods)
  };

  return replacements;
}

function processFiles(outputDir, replacements) {
  for (const filename of MODULE_CONFIG.templateFiles) {
    const srcPath = path.join(__dirname, filename);
    let destFilename = MODULE_CONFIG.fileRenames[filename] || filename;

    // Apply replacements to filename
    for (const [search, replace] of Object.entries(replacements)) {
      destFilename = destFilename.split(search).join(replace);
    }

    const destPath = path.join(outputDir, destFilename);

    if (fs.existsSync(srcPath)) {
      let content = fs.readFileSync(srcPath, 'utf8');

      // Apply replacements to content
      for (const [search, replace] of Object.entries(replacements)) {
        content = content.split(search).join(replace);
      }

      fs.writeFileSync(destPath, content);
      console.log(`  ✓ ${destFilename}`);
    } else {
      console.log(`  ⚠ ${filename} not found, skipping`);
    }
  }
}

function createClaspTemplate(outputDir) {
  const template = {
    scriptId: 'WILL_BE_SET_BY_CLASP_CREATE',
    rootDir: './'
  };
  const templatePath = path.join(outputDir, '.clasp.json.template');
  fs.writeFileSync(templatePath, JSON.stringify(template, null, 2));
  console.log('  ✓ .clasp.json.template');
}

// =========================================
// Clasp Deployment
// =========================================

function deployWithClasp(outputDir, projectTitle) {
  console.log('\n--- Creating Google Apps Script Project ---');

  try {
    process.chdir(outputDir);

    // Check if clasp is available
    try {
      execSync('clasp --version', { stdio: 'pipe' });
    } catch (e) {
      console.log('\nNote: clasp CLI not found or not logged in.');
      console.log('Install with: npm install -g @google/clasp');
      console.log('Login with: clasp login');
      console.log('\nFiles have been prepared in the output directory.');
      console.log('You can manually run these commands later:');
      console.log(`  cd "${outputDir}"`);
      console.log(`  clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`);
      console.log('  clasp push --force');
      return null;
    }

    // Create new project
    console.log('Creating new spreadsheet and script...');
    execSync(`clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`, { stdio: 'inherit' });

    // Push files
    console.log('\nPushing files to Google Apps Script...');
    execSync('clasp push --force', { stdio: 'inherit' });

    // Read the created .clasp.json
    const claspJsonPath = path.join(outputDir, '.clasp.json');
    if (fs.existsSync(claspJsonPath)) {
      return JSON.parse(fs.readFileSync(claspJsonPath, 'utf8'));
    }

    return null;

  } catch (error) {
    console.error('\nError during clasp operations:', error.message);
    console.log('\nFiles have been prepared. You can manually run:');
    console.log(`  cd "${outputDir}"`);
    console.log(`  clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`);
    console.log('  clasp push --force');
    return null;
  }
}

// =========================================
// Run
// =========================================

main().catch(err => {
  console.error('Setup failed:', err);
  rl.close();
  process.exit(1);
});
