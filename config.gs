/**
 * Company Configuration
 * Edit these values for each client deployment
 */

const COMPANY_CONFIG = {
  // Company Details
  NAME: '{{COMPANY_NAME}}',
  SHORT_NAME: '{{SHORT_NAME}}',

  // Menu & UI Settings
  MENU_ICON: '{{MENU_ICON}}',
  MENU_TITLE: '{{MENU_ICON}} {{SHORT_NAME}} Contracts',

  // Contact Information
  EMAIL: '{{COMPANY_EMAIL}}',
  PHONE: '{{COMPANY_PHONE}}',

  // Branding
  SIGNATURE: 'Best regards,\n{{COMPANY_NAME}}',

  // Admin Settings
  TRUENORTH_EMAIL: 'jacob@truenorthstrategyops.com'
};

/**
 * Get company name for use in scripts
 */
function getCompanyName() {
  return COMPANY_CONFIG.NAME;
}

/**
 * Get menu title for use in scripts
 */
function getMenuTitle() {
  return COMPANY_CONFIG.MENU_TITLE;
}

/**
 * Get email signature
 */
function getSignature() {
  return COMPANY_CONFIG.SIGNATURE;
}
