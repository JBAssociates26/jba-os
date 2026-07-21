/**
 * JBA OS — Native Checklist Engine
 *
 * Reusable checklist framework for any workflow action.
 * Initial implementation: SELLER_LISTING / PRE_LISTING / PRELIST_CHECKLIST.
 *
 * SHEETS CREATED
 * - Checklist Templates
 * - Checklist Responses
 *
 * FIRST-TIME SETUP
 * Run setupNativeChecklistEngine() once.
 */

const JBA_CHECKLIST = {
  templatesSheet: 'Checklist Templates',
  responsesSheet: 'Checklist Responses',
  itemTypes: ['CHECKBOX', 'TEXT', 'TEXTAREA', 'SELECT', 'DATE', 'TIME', 'NUMBER', 'CURRENCY', 'EMAIL', 'PHONE', 'URL', 'DATETIME', 'MULTISELECT']
};

/**
 * Creates checklist sheets and seeds the initial PRELIST_CHECKLIST template.
 * Safe to run more than once.
 */

/**
 * ONE-TIME V4 UPGRADE
 *
 * Rebuilds Checklist Templates with the flexible V4 schema while preserving
 * existing Pre-Listing and Photography questions. The old template sheet is
 * copied to a timestamped backup before any changes are made.
 *
 * Run upgradeChecklistEngineV4() once from the Apps Script editor.
 */
function upgradeChecklistEngineV4() {
  const ss = getDatabase_();
  const templateName = JBA_CHECKLIST.templatesSheet;
  const oldSheet = ss.getSheetByName(templateName);
  const legacyRows = oldSheet ? sheetObjects_(templateName) : [];

  if (oldSheet) {
    const stamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'America/Detroit',
      'yyyyMMdd_HHmmss'
    );
    const backup = oldSheet.copyTo(ss);
    backup.setName('Checklist Templates Backup ' + stamp);
  }

  if (oldSheet) {
    oldSheet.clear();
  }

  const sheet = oldSheet || ss.insertSheet(templateName);
  const headers = checklistTemplateHeadersV4_();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1f2933')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Preserve existing non-MLS templates by translating the legacy schema.
  const migrated = legacyRows
    .filter(row => row['Action Key'] !== 'MLS_SUBMISSION')
    .map(convertLegacyChecklistRowToV4_);

  if (migrated.length) {
    writeChecklistTemplateRowsV4_(migrated);
  }

  seedMlsSubmissionChecklistTemplateV4_();
  ensureParallelActionStatusSheet_();

  sheet.autoResizeColumns(1, headers.length);

  return {
    success: true,
    migratedRows: migrated.length,
    message:
      'Checklist Templates upgraded to V4. The old sheet was backed up, and the MLS Submission template was rebuilt.'
  };
}

function checklistTemplateHeadersV4_() {
  return [
    'Template ID',
    'Workflow Key',
    'Stage Key',
    'Action Key',
    'Section',
    'Item Order',
    'Item Key',
    'Item Label',
    'Item Type',
    'Required Mode',
    'Conditional Rule',
    'Options',
    'Default Value',
    'Read Only?',
    'Source Field',
    'Help Text',
    'Active?'
  ];
}

function convertLegacyChecklistRowToV4_(row) {
  return {
    'Template ID': row['Template ID'] || '',
    'Workflow Key': row['Workflow Key'] || '',
    'Stage Key': row['Stage Key'] || '',
    'Action Key': row['Action Key'] || '',
    'Section': inferChecklistSectionV4_(row),
    'Item Order': Number(row['Item Order'] || 0),
    'Item Key': row['Item Key'] || '',
    'Item Label': row['Item Label'] || '',
    'Item Type': row['Item Type'] || 'TEXT',
    'Required Mode': row['Required?'] === 'Yes' ? 'YES' : 'NO',
    'Conditional Rule': '',
    'Options': row['Options'] || '',
    'Default Value': '',
    'Read Only?': 'No',
    'Source Field': '',
    'Help Text': row['Help Text'] || '',
    'Active?': row['Active?'] || 'Yes'
  };
}

function inferChecklistSectionV4_(row) {
  const action = row['Action Key'] || '';
  const key = row['Item Key'] || '';

  if (action === 'PHOTO_ORDER') {
    if (/PHOTO_DATE|PHOTO_TIME|PHOTOGRAPHER/.test(key)) return 'Appointment';
    if (/FLOOR|MEASURE|DRONE|VIDEO|TWILIGHT|STAGING|SERVICE/.test(key)) return 'Services';
    if (/ACCESS|ATTEND|LOCKBOX|OCCUPANCY/.test(key)) return 'Access';
    return 'Additional Information';
  }

  if (action === 'PRELIST_CHECKLIST') {
    if (/APPOINTMENT/.test(key)) return 'Appointment';
    if (/CONTACT|LEAD_SOURCE/.test(key)) return 'Client Information';
    if (/CMA|PRESENTATION|PAPERWORK/.test(key)) return 'Preparation';
    if (/BINDER/.test(key)) return 'Binder';
    return 'Additional Information';
  }

  return 'Checklist';
}

function writeChecklistTemplateRowsV4_(rows) {
  if (!rows || !rows.length) return;

  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.templatesSheet);
  const headers = checklistTemplateHeadersV4_();
  const values = rows.map(row => headers.map(header => row[header] ?? ''));

  sheet.getRange(
    sheet.getLastRow() + 1,
    1,
    values.length,
    headers.length
  ).setValues(values);
}

function seedMlsSubmissionChecklistTemplateV4_() {
  const rows = [
    checklistV4Row_('LISTING_DATES', 'Listing Dates', 10, 'COMING_SOON_DATE',
      'Coming Soon Date', 'DATE', 'NO', '', '', '',
      'Leave blank when the property will go directly live.'),
    checklistV4Row_('LISTING_DATES', 'Listing Dates', 20, 'LIVE_DATE',
      'Live Date', 'DATE', 'YES', '', '', '',
      'The anticipated date the listing should become active.'),

    checklistV4Row_('MARKETING', 'Marketing', 30, 'PERSONAL_SIGN',
      'Using Personal Sign?', 'SELECT', 'YES', '', 'Yes|No', '', ''),
    checklistV4Row_('MARKETING', 'Marketing', 40, 'OPEN_HOUSE',
      'Open House?', 'SELECT', 'YES', '', 'Yes|No', 'No', ''),
    checklistV4Row_('MARKETING', 'Marketing', 50, 'OPEN_HOUSE_DATE',
      'Open House Date', 'DATE', 'CONDITIONAL', 'OPEN_HOUSE=Yes', '', '',
      'Required when Open House is Yes.'),
    checklistV4Row_('MARKETING', 'Marketing', 60, 'OPEN_HOUSE_START_TIME',
      'Open House Start Time', 'TIME', 'CONDITIONAL', 'OPEN_HOUSE=Yes', '', '',
      'Required when Open House is Yes.'),
    checklistV4Row_('MARKETING', 'Marketing', 70, 'OPEN_HOUSE_END_TIME',
      'Open House End Time', 'TIME', 'CONDITIONAL', 'OPEN_HOUSE=Yes', '', '',
      'Required when Open House is Yes.'),

    checklistV4Row_('SHOWING', 'Showing Setup', 80, 'SHOWING_INSTRUCTIONS',
      'Showing Instructions', 'SELECT', 'YES', '',
      'Appointment required|Go and show|Call listing agent|Call seller|Other', '', ''),
    checklistV4Row_('SHOWING', 'Showing Setup', 90, 'SHOWING_INSTRUCTIONS_OTHER',
      'Additional Showing Instructions', 'TEXTAREA', 'CONDITIONAL',
      'SHOWING_INSTRUCTIONS=Other', '', '',
      'Required when Showing Instructions is Other.'),
    checklistV4Row_('SHOWING', 'Showing Setup', 100, 'OCCUPANCY_STATUS',
      'Occupancy Status', 'SELECT', 'YES', '',
      'Occupied|Vacant|Owner occupied|Tenant occupied', '', ''),
    checklistV4Row_('SHOWING', 'Showing Setup', 110, 'OCCUPANCY_NOTES',
      'Occupancy / Showing Restrictions', 'TEXTAREA', 'NO', '', '', '',
      'Include notice requirements, pets, alarms, preferred times, or access restrictions.'),

    checklistV4Row_('COMPLIANCE', 'Compliance', 120,
      'INSPECTION_REQUIREMENTS_CONTACTED',
      'Have you called the city and county for inspection requirements?',
      'SELECT', 'YES', '', 'Yes|No|Not required', '', ''),

    checklistV4Row_('NOTES', 'Additional Information', 130, 'MLS_NOTES',
      'MLS Submission Notes', 'TEXTAREA', 'NO', '', '', '', ''),

    checklistV4Row_('CONFIRMATION', 'Final Confirmation', 140, 'READY_FOR_TC',
      'This listing is ready for Transaction Coordinator review and MLS entry.',
      'CHECKBOX', 'YES', '', '', '', '')
  ];

  writeChecklistTemplateRowsV4_(rows);
}

function checklistV4Row_(
  sectionKey,
  section,
  order,
  itemKey,
  label,
  type,
  requiredMode,
  conditionalRule,
  options,
  defaultValue,
  helpText
) {
  return {
    'Template ID': 'SELLER_LISTING_MLS_SUBMISSION_V2',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'LISTING_PREPARATION',
    'Action Key': 'MLS_SUBMISSION',
    'Section': section,
    'Item Order': order,
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Required Mode': requiredMode,
    'Conditional Rule': conditionalRule,
    'Options': options,
    'Default Value': defaultValue,
    'Read Only?': 'No',
    'Source Field': '',
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}

/* =========================================================
   CONTRACT-TO-CLOSE CHECKLISTS
   Accepted Offer (Agent) -> Under Contract (TC) -> Closing (TC)

   These are a starting draft. Since templates are config-driven,
   adjust items directly in the Checklist Templates sheet afterward —
   no code changes needed.

   Run setupContractToCloseChecklists() once. Safe to re-run.
   ========================================================= */

function setupContractToCloseChecklists() {
  seedAcceptedOfferChecklistTemplate_();
  seedContractChecklistTemplate_();
  seedClosingChecklistTemplate_();
}

/**
 * Deactivates every existing template item for a Workflow Key + Action Key
 * so a reseed fully replaces the item set instead of leaving old,
 * since-removed items active alongside the new ones. upsertChecklistTemplateRows_
 * only ever inserts/updates by key — it never removes — so call this first
 * when a checklist's item list has changed shape.
 */
function deactivateChecklistTemplateItems_(workflowKey, actionKey) {
  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.templatesSheet);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const workflowCol = headers.indexOf('Workflow Key');
  const actionCol = headers.indexOf('Action Key');
  const activeCol = headers.indexOf('Active?');

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    if (
      data[rowIndex][workflowCol] === workflowKey &&
      data[rowIndex][actionCol] === actionKey
    ) {
      sheet.getRange(rowIndex + 1, activeCol + 1).setValue('No');
    }
  }
}

function acceptedOfferChecklistRow_(
  section, order, itemKey, label, type,
  requiredMode, options, helpText
) {
  return {
    'Template ID': 'SELLER_LISTING_ACCEPTED_OFFER_V2',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'LIVE',
    'Action Key': 'ACCEPTED_OFFER',
    'Section': section,
    'Item Order': order,
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Required Mode': requiredMode,
    'Conditional Rule': '',
    'Options': options,
    'Default Value': '',
    'Read Only?': 'No',
    'Source Field': '',
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}

/**
 * Kept minimal on purpose: full contract details are captured by the TC
 * at Under Contract (matches how the team's existing checklist works).
 * This is just the agent's handoff confirming an offer was accepted.
 */
function seedAcceptedOfferChecklistTemplate_() {
  deactivateChecklistTemplateItems_('SELLER_LISTING', 'ACCEPTED_OFFER');

  const rows = [
    acceptedOfferChecklistRow_('Offer', 10, 'OFFER_NOTES', 'Notes for the Transaction Coordinator', 'TEXTAREA', 'NO', '', ''),
    acceptedOfferChecklistRow_('Final Confirmation', 20, 'READY_FOR_CONTRACT', 'An offer has been accepted and this file is ready for Transaction Coordinator.', 'CHECKBOX', 'YES', '', '')
  ];

  upsertChecklistTemplateRows_(rows);
}

function contractChecklistRow_(
  section, order, itemKey, label, type,
  requiredMode, conditionalRule, options, readOnly, sourceField, helpText
) {
  return {
    'Template ID': 'SELLER_LISTING_CONTRACT_CHECKLIST_V2',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'UNDER_CONTRACT',
    'Action Key': 'CONTRACT_CHECKLIST',
    'Section': section,
    'Item Order': order,
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Required Mode': requiredMode,
    'Conditional Rule': conditionalRule,
    'Options': options,
    'Default Value': '',
    'Read Only?': readOnly,
    'Source Field': sourceField,
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}

/**
 * Matches the fields on the team's existing Under Contract checklist.
 * Address/City/ZIP/Agent/Client name are already on the transaction and
 * shown in the checklist header, so they aren't repeated here.
 */
function seedContractChecklistTemplate_() {
  deactivateChecklistTemplateItems_('SELLER_LISTING', 'CONTRACT_CHECKLIST');

  const rows = [
    contractChecklistRow_('Seller Details', 10, 'PRIMARY_RESIDENCE', 'Will this be their primary residence?', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),
    contractChecklistRow_('Seller Details', 20, 'MARITAL_STATUS', 'Marital Status', 'TEXT', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Seller Details', 30, 'US_CITIZEN', 'US Citizen', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),
    contractChecklistRow_('Seller Details', 40, 'CLIENT_CURRENT_ADDRESS', 'Client Current Address', 'TEXT', 'NO', '', '', 'No', '', 'If different from the listed property.'),

    contractChecklistRow_('Second Seller (if applicable)', 50, 'SECOND_CONTACT_NAME', 'Second Contact Name', 'TEXT', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Second Seller (if applicable)', 60, 'SECOND_CONTACT_EMAIL', 'Second Contact Email', 'EMAIL', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Second Seller (if applicable)', 70, 'SECOND_CONTACT_PHONE', 'Second Contact Phone', 'PHONE', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Second Seller (if applicable)', 80, 'SECOND_MARITAL_STATUS', 'Second Client Marital Status', 'TEXT', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Second Seller (if applicable)', 90, 'SECOND_US_CITIZEN', 'Second Client US Citizen', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),

    contractChecklistRow_('Buyer Side', 100, 'COOPERATING_AGENT_NAME', 'Cooperating Agent Name', 'TEXT', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Buyer Side', 110, 'COOPERATING_AGENT_EMAIL', 'Cooperating Agent Email', 'EMAIL', 'NO', '', '', 'No', '', ''),

    contractChecklistRow_('Contract Terms', 120, 'TRANSACTION_AMOUNT', 'Transaction Amount', 'CURRENCY', 'YES', '', '', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 130, 'EARNEST_MONEY_AMOUNT', 'Earnest Money Amount', 'CURRENCY', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 140, 'SEND_EARNEST_REQUEST', 'Send Earnest Request', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 150, 'UNDER_CONTRACT_DATE', 'Under Contract Date', 'DATE', 'YES', '', '', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 160, 'FORECASTED_CLOSED_DATE', 'Forecasted Closed Date', 'DATE', 'YES', '', '', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 170, 'CONCESSIONS_AMOUNT', 'Concessions Amount', 'CURRENCY', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 180, 'POSSESSION', 'Possession', 'TEXT', 'NO', '', '', 'No', '', 'e.g. At close'),
    contractChecklistRow_('Contract Terms', 190, 'ADDENDUMS_OUTSIDE_PA', 'Any Addendums Outside of the PA', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 200, 'ADDENDUMS_LIST', 'If Yes, List Each One', 'TEXTAREA', 'CONDITIONAL', 'ADDENDUMS_OUTSIDE_PA=Yes', '', 'No', '', ''),
    contractChecklistRow_('Contract Terms', 210, 'ABO_OR_PENDING', 'ABO or Pending', 'SELECT', 'NO', '', 'ABO|Pending', 'No', '', ''),

    contractChecklistRow_('Inspection', 220, 'INSPECTION_DATE', 'Inspection Date', 'DATE', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Inspection', 230, 'INSPECTION_TIME', 'Inspection Time', 'TIME', 'NO', '', '', 'No', '', ''),

    contractChecklistRow_('Lead Source', 240, 'PRIMARY_LEAD_SOURCE', 'Primary Lead Source', 'TEXT', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Lead Source', 250, 'LEAD_SOURCE', 'Lead Source', 'TEXT', 'NO', '', '', 'No', '', ''),

    contractChecklistRow_('Commission & Fees', 260, 'TOTAL_COMMISSION_PERCENT', 'Total Commission %', 'NUMBER', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 270, 'YOUR_COMMISSION_PERCENT', 'Your Commission Percentage', 'NUMBER', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 280, 'TRANSACTION_FEE', 'Transaction Fee', 'CURRENCY', 'NO', '', '', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 290, 'HOME_WARRANTY', 'Home Warranty', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 300, 'HOME_WARRANTY_PAID_BY', 'Home Warranty Paid By', 'TEXT', 'CONDITIONAL', 'HOME_WARRANTY=Yes', '', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 310, 'HOME_WARRANTY_AMOUNT', 'Home Warranty Amount', 'CURRENCY', 'CONDITIONAL', 'HOME_WARRANTY=Yes', '', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 320, 'REFERRAL', 'Referral', 'SELECT', 'NO', '', 'Yes|No', 'No', '', ''),
    contractChecklistRow_('Commission & Fees', 330, 'REFERRAL_AMOUNT', 'Referral Amount', 'CURRENCY', 'CONDITIONAL', 'REFERRAL=Yes', '', 'No', '', ''),

    contractChecklistRow_('Additional Information', 340, 'CONTRACT_NOTES', 'Notes', 'TEXTAREA', 'NO', '', '', 'No', '', ''),

    contractChecklistRow_('Final Confirmation', 350, 'READY_FOR_CLOSING_PREP', 'This file is ready for closing preparation.', 'CHECKBOX', 'YES', '', '', 'No', '', '')
  ];

  upsertChecklistTemplateRows_(rows);
}

function closingChecklistRow_(
  section, order, itemKey, label, type,
  requiredMode, options, readOnly, sourceField, helpText
) {
  return {
    'Template ID': 'SELLER_LISTING_CLOSING_CHECKLIST_V1',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'CLOSING',
    'Action Key': 'CLOSING_CHECKLIST',
    'Section': section,
    'Item Order': order,
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Required Mode': requiredMode,
    'Conditional Rule': '',
    'Options': options,
    'Default Value': '',
    'Read Only?': readOnly,
    'Source Field': sourceField,
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}

function seedClosingChecklistTemplate_() {
  const rows = [
    closingChecklistRow_('Closing Details', 10, 'CLOSING_DATE', 'Closing Date', 'DATE', 'NO', '', 'Yes', 'Closing Date', ''),

    closingChecklistRow_('Final Steps', 20, 'FINAL_WALKTHROUGH_COMPLETE', 'Final Walkthrough Completed?', 'SELECT', 'YES', 'Yes|No', 'No', '', ''),
    closingChecklistRow_('Final Steps', 30, 'CLOSING_DISCLOSURE_SENT', 'Closing Disclosure Sent?', 'SELECT', 'YES', 'Yes|No', 'No', '', ''),
    closingChecklistRow_('Final Steps', 40, 'KEYS_TRANSFERRED', 'Keys/Access Transferred?', 'SELECT', 'YES', 'Yes|No', 'No', '', ''),
    closingChecklistRow_('Final Steps', 50, 'COMMISSION_CONFIRMED', 'Commission Disbursement Confirmed?', 'SELECT', 'YES', 'Yes|No', 'No', '', ''),
    closingChecklistRow_('Final Steps', 60, 'MLS_UPDATED_CLOSED', 'MLS Status Updated to Closed?', 'SELECT', 'YES', 'Yes|No', 'No', '', ''),

    closingChecklistRow_('Additional Information', 70, 'CLOSING_NOTES', 'Notes', 'TEXTAREA', 'NO', '', 'No', '', ''),

    closingChecklistRow_('Final Confirmation', 80, 'READY_TO_ARCHIVE', 'Move to Post Close', 'CHECKBOX', 'YES', '', 'No', '', '')
  ];

  upsertChecklistTemplateRows_(rows);
}

/**
 * Writes the Under Contract checklist's transaction amount / forecasted
 * closed date onto the transaction row so the Closing checklist can
 * display the closing date read-only via Source Field.
 */
function syncContractDetailsToTransaction_(tx, checklist) {
  const values = {};
  checklist.items.forEach(item => values[item.itemKey] = item.value || '');

  updateTransactionFields_(tx['Transaction ID'], {
    'Contract Price': values['TRANSACTION_AMOUNT'] || '',
    'Closing Date': values['FORECASTED_CLOSED_DATE'] || '',
    'MLS Status': values['ABO_OR_PENDING'] || ''
  });
}

function isChecklistTemplateV4_() {
  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.templatesSheet);
  if (!sheet || sheet.getLastRow() === 0) return false;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  return headers.includes('Required Mode') &&
    headers.includes('Conditional Rule') &&
    headers.includes('Section');
}

function evaluateChecklistCondition_(rule, values) {
  if (!rule) return true;

  const match = String(rule).match(/^\s*([^!=]+?)\s*(=|!=)\s*(.*?)\s*$/);
  if (!match) return true;

  const key = match[1].trim();
  const operator = match[2];
  const expected = match[3].trim();
  const actual = String(values[key] || '').trim();

  return operator === '=' ? actual === expected : actual !== expected;
}

function resolveChecklistSourceValue_(tx, sourceField) {
  if (!sourceField) return '';
  return tx[sourceField] ?? '';
}


function setupNativeChecklistEngine() {
  const ss = getDatabase_();

  const existingTemplateSheet = ss.getSheetByName(JBA_CHECKLIST.templatesSheet);
  if (existingTemplateSheet && existingTemplateSheet.getLastRow() > 0 && !isChecklistTemplateV4_()) {
    throw new Error(
      'Checklist Templates is using the older schema. Run upgradeChecklistEngineV4() once instead.'
    );
  }

  ensureNativeChecklistSheet_(
    ss,
    JBA_CHECKLIST.templatesSheet,
    checklistTemplateHeadersV4_()
  );

  ensureNativeChecklistSheet_(
    ss,
    JBA_CHECKLIST.responsesSheet,
    [
      'Transaction ID',
      'Action Key',
      'Item Key',
      'Value',
      'Completed?',
      'Updated At',
      'Updated By'
    ]
  );

  ensureParallelActionStatusSheet_();

  const currentTemplates = sheetObjects_(JBA_CHECKLIST.templatesSheet);
  if (!currentTemplates.length) {
    seedMlsSubmissionChecklistTemplateV4_();
  }

  return {
    success: true,
    message: 'Native Checklist Engine is ready, including parallel Photography and MLS Submission checklists.'
  };
}

function ensureNativeChecklistSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * Seeds a streamlined pre-listing task.
 * Edit labels, options, order, or Active? directly in Checklist Templates.
 */
function seedPreListChecklistTemplate_() {
  const existing = sheetObjects_(JBA_CHECKLIST.templatesSheet)
    .filter(row =>
      row['Workflow Key'] === 'SELLER_LISTING' &&
      row['Action Key'] === 'PRELIST_CHECKLIST'
    );

  if (existing.length) return;

  const items = [
    checklistTemplateRow_(
      'APPOINTMENT_FORMAT',
      'Appointment format',
      'SELECT',
      'In person at the property|In person at another location|Video appointment|Phone appointment',
      'Yes',
      10,
      'Choose how the appointment will take place.'
    ),
    checklistTemplateRow_(
      'APPOINTMENT_LOCATION',
      'Appointment location if different from the property',
      'TEXT',
      '',
      'No',
      20,
      ''
    ),
    checklistTemplateRow_(
      'LEAD_SOURCE',
      'Lead source',
      'SELECT',
      'Sphere|Past Client|Referral|Database|Open House|Sign Call|Internet Lead|Social Media|Agent Referral|Builder|Other',
      'Yes',
      30,
      ''
    ),
    checklistTemplateRow_(
      'SECOND_CONTACT_NAME',
      'Second seller or decision-maker',
      'TEXT',
      '',
      'No',
      40,
      'Leave blank when there is no second contact.'
    ),
    checklistTemplateRow_(
      'SECOND_CONTACT_PHONE',
      'Second contact phone',
      'TEXT',
      '',
      'No',
      50,
      ''
    ),
    checklistTemplateRow_(
      'SECOND_CONTACT_EMAIL',
      'Second contact email',
      'TEXT',
      '',
      'No',
      60,
      ''
    ),
    checklistTemplateRow_(
      'APPOINTMENT_CONFIRMED',
      'Appointment confirmed with the seller',
      'CHECKBOX',
      '',
      'Yes',
      70,
      ''
    ),
    checklistTemplateRow_(
      'CMA_PREPARED',
      'Comparative market analysis prepared',
      'CHECKBOX',
      '',
      'Yes',
      80,
      ''
    ),
    checklistTemplateRow_(
      'LISTING_PRESENTATION_PREPARED',
      'Listing presentation prepared',
      'CHECKBOX',
      '',
      'Yes',
      90,
      ''
    ),
    checklistTemplateRow_(
      'REQUIRED_PAPERWORK_PREPARED',
      'Required listing paperwork prepared',
      'CHECKBOX',
      '',
      'Yes',
      100,
      'Prepare the documents appropriate for this property type.'
    ),
    checklistTemplateRow_(
      'BINDER_DELIVERY_METHOD',
      'Binder delivery method',
      'SELECT',
      'Not applicable|Printed binder for agent pickup|Printed binder delivered to office|Digital binder only|Both printed and digital',
      'No',
      110,
      'JBA OS already knows whether a binder was requested.'
    ),
    checklistTemplateRow_(
      'BINDER_NEEDED_BY_DATE',
      'Binder needed by date',
      'DATE',
      '',
      'No',
      120,
      ''
    ),
    checklistTemplateRow_(
      'BINDER_INSTRUCTIONS',
      'Binder preparation instructions',
      'TEXTAREA',
      '',
      'No',
      130,
      'Include property-specific, seller-specific, branding, or rush instructions.'
    ),
    checklistTemplateRow_(
      'PROCESSING_FEE',
      'JBA transaction / processing fee',
      'SELECT',
      'Standard ($495)|Different Amount|Waived|Unsure – Please Verify',
      'Yes',
      140,
      ''
    ),
    checklistTemplateRow_(
      'PROCESSING_FEE_DETAILS',
      'Non-standard processing fee explanation',
      'TEXTAREA',
      '',
      'No',
      150,
      'Complete only when the fee is different, waived, or needs review.'
    ),
    checklistTemplateRow_(
      'REFERRAL_DETAILS',
      'Referral details',
      'TEXTAREA',
      '',
      'No',
      160,
      'Include the referring person or company and expected percentage or amount.'
    ),
    checklistTemplateRow_(
      'DATABASE_MATCHES',
      'Database best matches needed',
      'SELECT',
      'No|Yes',
      'Yes',
      170,
      ''
    ),
    checklistTemplateRow_(
      'DATABASE_MATCH_CRITERIA',
      'Database match criteria',
      'TEXTAREA',
      '',
      'No',
      180,
      'Include price range, geography, property type, bedrooms, acreage, or other criteria.'
    ),
    checklistTemplateRow_(
      'NET_SHEET',
      'Seller net sheet needed',
      'SELECT',
      'No|Yes',
      'Yes',
      190,
      ''
    ),
    checklistTemplateRow_(
      'NET_SHEET_DETAILS',
      'Net sheet details',
      'TEXTAREA',
      '',
      'No',
      200,
      'Include estimated sale price, commission, mortgage payoff, costs, and concessions.'
    ),
    checklistTemplateRow_(
      'SEND_DOCS_BEFORE_APPT',
      'Send documents to seller for signature before appointment',
      'SELECT',
      'No|Yes',
      'Yes',
      210,
      ''
    ),
    checklistTemplateRow_(
      'OPEN_HOUSE_DISCUSSION',
      'Open house discussion needed',
      'SELECT',
      'No|Yes|Maybe – Still Verifying',
      'Yes',
      220,
      ''
    ),
    checklistTemplateRow_(
      'MARKETING_APPROVAL',
      'Create flyers or marketing before further approval',
      'SELECT',
      'No – Wait for Agent Approval|Yes',
      'Yes',
      230,
      ''
    ),
    checklistTemplateRow_(
      'SPECIAL_INSTRUCTIONS',
      'Notes and special instructions',
      'TEXTAREA',
      '',
      'No',
      240,
      ''
    ),
    checklistTemplateRow_(
      'READY_FOR_PREPARATION',
      'Information is complete and ready for preparation',
      'CHECKBOX',
      '',
      'Yes',
      250,
      ''
    )
  ];

  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.templatesSheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  sheet.getRange(2, 1, items.length, headers.length)
    .setValues(items.map(item => headers.map(header => item[header] || '')));
}


/**
 * Seeds the native Order Photography task.
 * Existing transaction data is shown as context, so agents only enter
 * photography-specific information.
 */
function seedPhotoOrderChecklistTemplate_() {
  const existing = sheetObjects_(JBA_CHECKLIST.templatesSheet)
    .filter(row =>
      row['Workflow Key'] === 'SELLER_LISTING' &&
      row['Action Key'] === 'PHOTO_ORDER'
    );

  if (existing.length) return;

  const items = [
    photoChecklistTemplateRow_(
      'ALL_DOCS_SIGNED',
      'All documents signed and turned in?',
      'SELECT',
      'Yes|No|Unsure – Please Verify',
      'Yes',
      10,
      ''
    ),
    photoChecklistTemplateRow_(
      'PHOTOGRAPHER',
      'Preferred photographer',
      'SELECT',
      'AP Photography|No preference|Other',
      'Yes',
      20,
      ''
    ),
    photoChecklistTemplateRow_(
      'PHOTOGRAPHER_OTHER',
      'Other photographer name',
      'TEXT',
      '',
      'No',
      30,
      'Complete only when Other is selected.'
    ),
    photoChecklistTemplateRow_(
      'PHOTO_DATE',
      'Date of photos',
      'DATE',
      '',
      'Yes',
      40,
      ''
    ),
    photoChecklistTemplateRow_(
      'PHOTO_TIME',
      'Time of photos',
      'TEXT',
      '',
      'Yes',
      50,
      'Examples: 2:00 PM, Morning, or Afternoon.'
    ),
    photoChecklistTemplateRow_(
      'SERVICE_2D_FLOOR_PLAN',
      '2D floor plan',
      'CHECKBOX',
      '',
      'No',
      60,
      ''
    ),
    photoChecklistTemplateRow_(
      'SERVICE_MEASUREMENTS',
      'Property measurements',
      'CHECKBOX',
      '',
      'No',
      70,
      ''
    ),
    photoChecklistTemplateRow_(
      'SERVICE_DRONE',
      'Drone / aerial photography',
      'CHECKBOX',
      '',
      'No',
      80,
      ''
    ),
    photoChecklistTemplateRow_(
      'SERVICE_VIDEO',
      'Property video',
      'CHECKBOX',
      '',
      'No',
      90,
      ''
    ),
    photoChecklistTemplateRow_(
      'SERVICE_TWILIGHT',
      'Twilight photography',
      'CHECKBOX',
      '',
      'No',
      100,
      ''
    ),
    photoChecklistTemplateRow_(
      'SERVICE_VIRTUAL_STAGING',
      'Virtual staging',
      'CHECKBOX',
      '',
      'No',
      110,
      ''
    ),
    photoChecklistTemplateRow_(
      'OTHER_PHOTO_SERVICES',
      'Other photo services',
      'TEXTAREA',
      '',
      'No',
      120,
      ''
    ),
    photoChecklistTemplateRow_(
      'PHOTOGRAPHER_ACCESS',
      'How will the photographer gain access?',
      'SELECT',
      'Agent Entry|Seller Entry|Lockbox|Vacant / Open Access|Tenant Entry|Other',
      'Yes',
      130,
      ''
    ),
    photoChecklistTemplateRow_(
      'ACCESS_INSTRUCTIONS',
      'Access instructions',
      'TEXTAREA',
      '',
      'No',
      140,
      'Include lockbox location, gate instructions, contact details, pets, alarms, or special access notes.'
    ),
    photoChecklistTemplateRow_(
      'AGENT_ATTENDING',
      'Will the agent be attending?',
      'SELECT',
      'Yes|No|Unsure',
      'Yes',
      150,
      ''
    ),
    photoChecklistTemplateRow_(
      'TAKING_LOCKBOX',
      'Taking a lockbox?',
      'SELECT',
      'Yes|No|Already installed|Not needed',
      'Yes',
      160,
      ''
    ),
    photoChecklistTemplateRow_(
      'OCCUPANCY_STATUS',
      'Occupancy status',
      'SELECT',
      'Occupied|Vacant|Tenant Occupied|New Construction|Other',
      'Yes',
      170,
      ''
    ),
    photoChecklistTemplateRow_(
      'BASEMENT',
      'Basement',
      'SELECT',
      'No Basement|Yes - Finished|Yes - Unfinished|Yes - Partially Finished',
      'Yes',
      180,
      ''
    ),
    photoChecklistTemplateRow_(
      'BASEMENT_FINISHED_SQFT',
      'Basement finished square footage',
      'NUMBER',
      '',
      'No',
      190,
      'Enter only finished basement square footage.'
    ),
    photoChecklistTemplateRow_(
      'PHOTO_NOTES',
      'Photography notes',
      'TEXTAREA',
      '',
      'No',
      200,
      'Include special rooms, acreage, outbuildings, water features, seller concerns, or anything the photographer and TC should know.'
    ),
    photoChecklistTemplateRow_(
      'READY_TO_ORDER_PHOTOS',
      'Photography order is complete and ready to send',
      'CHECKBOX',
      '',
      'Yes',
      210,
      ''
    )
  ];

  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.templatesSheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  sheet.getRange(sheet.getLastRow() + 1, 1, items.length, headers.length)
    .setValues(items.map(item => headers.map(header => item[header] || '')));
}

function photoChecklistTemplateRow_(itemKey, label, type, options, required, order, helpText) {
  return {
    'Template ID': 'SELLER_LISTING_PHOTO_ORDER_V1',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'LISTING_SECURED',
    'Action Key': 'PHOTO_ORDER',
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Options': options,
    'Required?': required,
    'Item Order': order,
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}

function checklistTemplateRow_(itemKey, label, type, options, required, order, helpText) {
  return {
    'Template ID': 'SELLER_LISTING_PRELIST_V1',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'PRE_LISTING',
    'Action Key': 'PRELIST_CHECKLIST',
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Options': options,
    'Required?': required,
    'Item Order': order,
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}


/* =========================================================
   PARALLEL LISTING PREPARATION
   Photography and MLS submission may be completed in either order.
   ========================================================= */

function ensureParallelActionStatusSheet_() {
  const ss = getDatabase_();
  const name = 'Parallel Action Status';
  let sheet = ss.getSheetByName(name);

  if (!sheet) sheet = ss.insertSheet(name);

  const headers = [
    'Transaction ID',
    'Action Key',
    'Status',
    'Completed At',
    'Completed By',
    'Updated At'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}


/**
 * Inserts new checklist-template rows and updates matching rows safely.
 * Matching key: Workflow Key + Action Key + Item Key.
 */
function upsertChecklistTemplateRows_(rows) {
  if (!rows || !rows.length) return;

  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.templatesSheet);

  if (!sheet) {
    throw new Error(
      'Checklist Templates sheet does not exist. Run setupNativeChecklistEngine() first.'
    );
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = {};

  headers.forEach((header, index) => {
    headerIndex[header] = index;
  });

  const requiredHeaders = [
    'Workflow Key',
    'Action Key',
    'Item Key'
  ];

  requiredHeaders.forEach(header => {
    if (headerIndex[header] === undefined) {
      throw new Error(
        'Checklist Templates is missing the required column: ' + header
      );
    }
  });

  const existingRowByKey = {};

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    const key = [
      row[headerIndex['Workflow Key']],
      row[headerIndex['Action Key']],
      row[headerIndex['Item Key']]
    ].join('||');

    existingRowByKey[key] = rowIndex + 1;
  }

  const rowsToAppend = [];

  rows.forEach(item => {
    const key = [
      item['Workflow Key'],
      item['Action Key'],
      item['Item Key']
    ].join('||');

    const values = headers.map(header => item[header] || '');
    const existingSheetRow = existingRowByKey[key];

    if (existingSheetRow) {
      sheet.getRange(existingSheetRow, 1, 1, headers.length)
        .setValues([values]);
    } else {
      rowsToAppend.push(values);
    }
  });

  if (rowsToAppend.length) {
    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      rowsToAppend.length,
      headers.length
    ).setValues(rowsToAppend);
  }
}

function seedMlsSubmissionChecklistTemplate_() {
  const rows = [
    mlsChecklistRow_('COMING_SOON_DATE', 'Coming Soon Date', 'DATE', '', 'No', 10,
      'Leave blank when the property will go directly live.'),
    mlsChecklistRow_('LIVE_DATE', 'Live Date', 'DATE', '', 'Yes', 20,
      'The anticipated date the listing should become active.'),
    mlsChecklistRow_('PERSONAL_SIGN', 'Using Personal Sign?', 'SELECT', 'Yes|No', 'Yes', 30, ''),
    mlsChecklistRow_('OPEN_HOUSE', 'Open House?', 'SELECT', 'Yes|No', 'Yes', 40, ''),
    mlsChecklistRow_('OPEN_HOUSE_DATE_TIME', 'Open House Date and Time', 'TEXT', '', 'No', 50,
      'Required when Open House is Yes. Include the date, start time, and end time.'),
    mlsChecklistRow_('SHOWING_INSTRUCTIONS', 'Showing Instructions', 'SELECT',
      'Appointment required|Go and show|Call listing agent|Call seller|Other', 'Yes', 60, ''),
    mlsChecklistRow_('SHOWING_INSTRUCTIONS_OTHER', 'Additional Showing Instructions', 'TEXTAREA', '', 'No', 70, ''),
    mlsChecklistRow_('OCCUPANCY_STATUS', 'Occupancy Status', 'SELECT',
      'Occupied|Vacant|Owner occupied|Tenant occupied', 'Yes', 80, ''),
    mlsChecklistRow_('OCCUPANCY_NOTES', 'Occupancy / Showing Restrictions', 'TEXTAREA', '', 'No', 90,
      'Include preferred showing windows, notice requirements, pets, alarms, or access restrictions.'),
    mlsChecklistRow_('INSPECTION_REQUIREMENTS_CONTACTED',
      'Have you called the city and county for inspection requirements?',
      'SELECT', 'Yes|No|Not required', 'Yes', 100, ''),
    mlsChecklistRow_('MLS_NOTES', 'MLS Submission Notes', 'TEXTAREA', '', 'No', 110, ''),
    mlsChecklistRow_('READY_FOR_TC',
      'This listing is ready for Transaction Coordinator review and MLS entry.',
      'CHECKBOX', '', 'Yes', 120, '')
  ];

  upsertChecklistTemplateRows_(rows);
}

function mlsChecklistRow_(itemKey, label, type, options, required, order, helpText) {
  return {
    'Template ID': 'SELLER_LISTING_MLS_SUBMISSION_V1',
    'Workflow Key': 'SELLER_LISTING',
    'Stage Key': 'LISTING_PREPARATION',
    'Action Key': 'MLS_SUBMISSION',
    'Item Key': itemKey,
    'Item Label': label,
    'Item Type': type,
    'Options': options,
    'Required?': required,
    'Item Order': order,
    'Help Text': helpText,
    'Active?': 'Yes'
  };
}

function getParallelListingPreparation(transactionId) {
  const auth = getAuthorizedTransaction_(transactionId);
  const tx = auth.tx;

  if (tx['Workflow Key'] !== 'SELLER_LISTING') {
    return { enabled: false, items: [] };
  }

  const statusMap = getParallelActionStatusMap_(transactionId);

  return {
    enabled: true,
    transactionId: transactionId,
    items: [
      {
        actionKey: 'PHOTO_ORDER',
        actionName: 'Order Photography',
        status: statusMap['PHOTO_ORDER'] || 'Not Started'
      },
      {
        actionKey: 'MLS_SUBMISSION',
        actionName: 'Complete MLS Submission',
        status: statusMap['MLS_SUBMISSION'] || 'Not Started'
      }
    ]
  };
}

function getParallelActionStatusMap_(transactionId) {
  ensureParallelActionStatusSheet_();

  const rows = sheetObjects_('Parallel Action Status')
    .filter(row => row['Transaction ID'] === transactionId);

  const map = {};
  rows.forEach(row => {
    map[row['Action Key']] = row['Status'] || 'Not Started';
  });

  return map;
}

function markParallelActionComplete_(transactionId, actionKey, userEmail) {
  ensureParallelActionStatusSheet_();

  const sheet = getDatabase_().getSheetByName('Parallel Action Status');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const txCol = headers.indexOf('Transaction ID');
  const actionCol = headers.indexOf('Action Key');
  const now = new Date();

  for (let r = 1; r < values.length; r++) {
    if (values[r][txCol] === transactionId && values[r][actionCol] === actionKey) {
      const row = {};
      headers.forEach((header, index) => row[header] = values[r][index]);
      row['Status'] = 'Complete';
      row['Completed At'] = now;
      row['Completed By'] = userEmail;
      row['Updated At'] = now;
      sheet.getRange(r + 1, 1, 1, headers.length)
        .setValues([headers.map(header => row[header] || '')]);
      return;
    }
  }

  const row = {
    'Transaction ID': transactionId,
    'Action Key': actionKey,
    'Status': 'Complete',
    'Completed At': now,
    'Completed By': userEmail,
    'Updated At': now
  };

  sheet.appendRow(headers.map(header => row[header] || ''));
}

function getChecklistActionByKey_(tx, actionKey) {
  const configured = sheetObjects_(JBA_OS.sheets.workflowActions)
    .find(row =>
      row['Workflow Key'] === tx['Workflow Key'] &&
      row['Action Key'] === actionKey &&
      row['Active?'] === 'Yes'
    );

  if (configured) return configured;

  const fallback = {
    'Workflow Key': tx['Workflow Key'],
    'Stage Key': tx['Current Stage Key'],
    'Action Key': actionKey,
    'Action Name':
      actionKey === 'PHOTO_ORDER'
        ? 'Order Photography'
        : 'Complete MLS Submission',
    'Assigned Role': 'Agent',
    'Active?': 'Yes'
  };

  return fallback;
}

function canOpenParallelChecklist_(user, tx, actionKey) {
  if (tx['Workflow Key'] !== 'SELLER_LISTING') return false;
  if (!['PHOTO_ORDER', 'MLS_SUBMISSION'].includes(actionKey)) return false;

  if (['Executive Admin', 'Operations Admin'].includes(user.role)) return true;

  const agentEmail = normalizeEmail_(tx['Agent Email']);
  const userEmail = normalizeEmail_(user.email);
  return user.role === 'Agent' && agentEmail === userEmail;
}

function getChecklistTemplateItemsForAction_(workflowKey, actionKey) {
  return sheetObjects_('Checklist Templates')
    .filter(row =>
      row['Workflow Key'] === workflowKey &&
      row['Action Key'] === actionKey &&
      row['Active?'] === 'Yes'
    )
    .sort((a, b) =>
      Number(a['Item Order'] || 0) - Number(b['Item Order'] || 0)
    );
}

function validateMlsConditionalRequirements_(checklist) {
  const values = {};
  checklist.items.forEach(item => values[item.itemKey] = item.value || '');

  if (values['COMING_SOON_DATE'] && values['LIVE_DATE']) {
    const comingSoon = new Date(values['COMING_SOON_DATE']);
    const live = new Date(values['LIVE_DATE']);

    if (!isNaN(comingSoon.getTime()) &&
        !isNaN(live.getTime()) &&
        comingSoon >= live) {
      throw new Error('Coming Soon Date must be before the Live Date.');
    }
  }

  if (values['OPEN_HOUSE'] === 'Yes') {
    const start = values['OPEN_HOUSE_START_TIME'];
    const end = values['OPEN_HOUSE_END_TIME'];

    if (start && end && start >= end) {
      throw new Error('Open House End Time must be after the Start Time.');
    }
  }
}

function sendMlsSubmissionNotificationToTC_(tx, checklist, user) {
  const assignedTc = normalizeEmail_(tx['Assigned TC Email']);
  const operations = normalizeEmail_(tx['Assigned Operations Email']);
  const recipient = assignedTc || operations;
  const agentEmail = normalizeEmail_(tx['Agent Email']);

  if (!recipient) {
    throw new Error(
      'No Transaction Coordinator or Operations email is assigned. Assign a recipient before completing the MLS submission.'
    );
  }

  const values = {};
  checklist.items.forEach(item => values[item.itemKey] = item.value || '');

  const address =
    tx['Property Address'] ||
    tx['Address Line 1'] ||
    'Address pending';

  const rows = checklist.items
    .filter(item => item.itemKey !== 'READY_FOR_TC')
    .map(item =>
      '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #ece8e2;font-weight:700;">' +
          escapeHtml_(item.label) +
        '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #ece8e2;">' +
          escapeHtml_(String(item.value || '—')) +
        '</td>' +
      '</tr>'
    )
    .join('');

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;color:#27231f;max-width:720px;">' +
      '<h2 style="margin-bottom:4px;">MLS Submission Ready for Review</h2>' +
      '<p style="margin-top:0;color:#6f685f;">' +
        escapeHtml_(address) +
      '</p>' +
      '<p>The agent has completed the MLS submission. Please review the information and enter the property into the MLS.</p>' +
      '<table style="border-collapse:collapse;width:100%;margin-top:18px;">' +
        rows +
      '</table>' +
      '<p style="margin-top:18px;"><strong>Submitted by:</strong> ' +
        escapeHtml_(user.displayName || user.email) +
      '</p>' +
    '</div>';

  MailApp.sendEmail({
    to: recipient,
    cc: agentEmail || '',
    subject: 'MLS Submission Ready for Review — ' + address,
    body:
      'The MLS submission is ready for review and MLS entry for ' +
      address + '.',
    htmlBody: htmlBody,
    name: 'JBA OS'
  });

  logActivity_(
    user,
    tx['Transaction ID'],
    'MLS_SUBMISSION_SENT_TO_TC',
    tx['Current Stage Name'],
    tx['Current Stage Name'],
    'Complete MLS Submission',
    'TC Review & Enter MLS',
    'MLS submission emailed to ' + recipient +
      (agentEmail ? ' and CC’d to ' + agentEmail : '')
  );
}


/**
 * Returns the native checklist for the transaction's current action.
 */
function getNativeChecklist(transactionId, requestedActionKey) {
  const auth = getAuthorizedTransaction_(transactionId);
  const tx = auth.tx;
  const actionKey =
    requestedActionKey ||
    tx['Current Action Key'];

  const action = getChecklistActionByKey_(tx, actionKey);

  if (!action) {
    throw new Error('No active action is configured.');
  }

  const isParallel = ['PHOTO_ORDER', 'MLS_SUBMISSION'].includes(actionKey);

  if (
    isParallel
      ? !canOpenParallelChecklist_(auth.user, tx, actionKey)
      : !canUserPerformAction_(auth.user, action)
  ) {
    throw new Error('You are not authorized to complete this checklist.');
  }

  const templates = isParallel
    ? getChecklistTemplateItemsForAction_(tx['Workflow Key'], actionKey)
    : getChecklistTemplateItems_(
        tx['Workflow Key'],
        tx['Current Stage Key'],
        actionKey
      );

  if (!templates.length) {
    throw new Error('No native checklist is configured for this task.');
  }

  const responses = getChecklistResponseMap_(
    transactionId,
    action['Action Key']
  );

  const preliminaryItems = templates.map(template => {
    const response = responses[template['Item Key']] || {};
    const readOnly = template['Read Only?'] === 'Yes';
    const responseValue = response.value;
    const defaultValue = template['Default Value'] || '';
    const sourceValue = readOnly
      ? resolveChecklistSourceValue_(tx, template['Source Field'])
      : '';

    return {
      itemKey: template['Item Key'],
      label: template['Item Label'],
      type: template['Item Type'],
      options: parseChecklistOptions_(template['Options']),
      requiredMode:
        template['Required Mode'] ||
        (template['Required?'] === 'Yes' ? 'YES' : 'NO'),
      conditionalRule: template['Conditional Rule'] || '',
      section: template['Section'] || 'Checklist',
      readOnly: readOnly,
      sourceField: template['Source Field'] || '',
      order: Number(template['Item Order'] || 0),
      helpText: template['Help Text'] || '',
      value:
        responseValue !== undefined && responseValue !== ''
          ? responseValue
          : (sourceValue !== '' ? sourceValue : defaultValue),
      completed: response.completed === true
    };
  });

  const currentValues = {};
  preliminaryItems.forEach(item => currentValues[item.itemKey] = item.value || '');

  const items = preliminaryItems.map(item => {
    const visible = evaluateChecklistCondition_(
      item.conditionalRule,
      currentValues
    );

    return Object.assign({}, item, {
      visible: visible,
      required:
        item.requiredMode === 'YES' ||
        (item.requiredMode === 'CONDITIONAL' && visible)
    });
  });

  const progress = calculateChecklistProgress_(items.filter(item => item.visible));

  return {
    success: true,
    transactionId: tx['Transaction ID'],
    actionKey: action['Action Key'],
    actionName: action['Action Name'],
    stageName: tx['Current Stage Name'],
    propertyAddress: tx['Property Address'] || tx['Address Line 1'] || '',
    sellerName: [tx['Client First Name'], tx['Client Last Name']]
      .filter(Boolean)
      .join(' '),
    listingAgent: tx['Listing Agent'] || '',
    appointment: [tx['Appointment Date'], tx['Appointment Time']]
      .filter(Boolean)
      .join(' at ') || 'Not scheduled',
    binderNeeded: tx['Binder Needed?'] || 'Not specified',
    urgency: tx['Urgency'] || 'Normal',
    progress: progress,
    items: items
  };
}

/**
 * Saves partial progress without advancing the workflow.
 */
function saveNativeChecklistProgress(transactionId, actionKey, submittedItems) {
  const auth = getAuthorizedTransaction_(transactionId);
  const action = getChecklistActionByKey_(auth.tx, actionKey);
  const isParallel = ['PHOTO_ORDER', 'MLS_SUBMISSION'].includes(actionKey);

  if (
    isParallel
      ? !canOpenParallelChecklist_(auth.user, auth.tx, actionKey)
      : !canUserPerformAction_(auth.user, action)
  ) {
    throw new Error('You are not authorized to complete this checklist.');
  }

  const templates = isParallel
    ? getChecklistTemplateItemsForAction_(auth.tx['Workflow Key'], actionKey)
    : getChecklistTemplateItems_(
        auth.tx['Workflow Key'],
        auth.tx['Current Stage Key'],
        actionKey
      );

  const submittedMap = {};
  (submittedItems || []).forEach(item => {
    submittedMap[String(item.itemKey || '')] = item;
  });

  templates.forEach(template => {
    const submitted = submittedMap[template['Item Key']];
    if (!submitted) return;

    upsertChecklistResponse_(
      transactionId,
      actionKey,
      template,
      submitted.value,
      auth.user.email
    );
  });

  const result = getNativeChecklist(transactionId, actionKey);

  logActivity_(
    auth.user,
    transactionId,
    'CHECKLIST_PROGRESS_SAVED',
    auth.tx['Current Stage Name'],
    auth.tx['Current Stage Name'],
    action['Action Name'],
    action['Action Name'],
    `${result.progress.completedRequired} of ${result.progress.totalRequired} required items complete`
  );

  return result;
}

/**
 * Validates required items, saves, and advances to the next workflow stage.
 */
function completeNativeChecklist(transactionId, actionKey, submittedItems) {
  const auth = getAuthorizedTransaction_(transactionId);
  const action = getChecklistActionByKey_(auth.tx, actionKey);
  const isParallel = ['PHOTO_ORDER', 'MLS_SUBMISSION'].includes(actionKey);

  if (
    isParallel
      ? !canOpenParallelChecklist_(auth.user, auth.tx, actionKey)
      : !canUserPerformAction_(auth.user, action)
  ) {
    throw new Error('You are not authorized to complete this checklist.');
  }

  if (
    isParallel &&
    isParallelWorkflowActionComplete_(transactionId, actionKey)
  ) {
    return reconcileParallelWorkflowForTransaction(transactionId);
  }

  saveNativeChecklistProgress(transactionId, actionKey, submittedItems);

  const checklist = getNativeChecklist(transactionId, actionKey);
  const missing = checklist.items.filter(item =>
    item.visible !== false &&
    item.required &&
    !isChecklistItemComplete_(item.type, item.value)
  );

  if (missing.length) {
    throw new Error(
      'Complete all required items before finishing: ' +
      missing.map(item => item.label).join(', ')
    );
  }

  if (actionKey === 'MLS_SUBMISSION') {
    validateMlsConditionalRequirements_(checklist);
    sendMlsSubmissionNotificationToTC_(auth.tx, checklist, auth.user);
    createTasksForAction_(auth.user, auth.tx, actionKey);
  }

  if (actionKey === 'PHOTO_ORDER') {
    sendPhotoOrderNotificationToTC_(auth.tx, checklist, auth.user);
    createTasksForAction_(auth.user, auth.tx, actionKey);
  }

  if (actionKey === 'CONTRACT_CHECKLIST') {
    syncContractDetailsToTransaction_(auth.tx, checklist);
  }

  if (isParallel) {
    return completeParallelWorkflowAction_(
      auth.user,
      auth.tx,
      action,
      checklist
    );
  }

  return advanceTransaction_(
    auth.user,
    auth.tx,
    action,
    `Native checklist completed: ${checklist.progress.totalRequired} required items`
  );
}


/**
 * Sends the completed photography order to the assigned TC.
 * Falls back to Operations when no TC is assigned so the order is not lost.
 */
function sendPhotoOrderNotificationToTC_(tx, checklist, user) {
  const assignedTc = normalizeEmail_(tx['Assigned TC Email']);
  const operations = normalizeEmail_(tx['Assigned Operations Email']);
  const recipient = assignedTc || operations;

  if (!recipient) {
    throw new Error(
      'No Transaction Coordinator or Operations email is assigned. Assign a recipient before completing the photography order.'
    );
  }

  const values = {};
  checklist.items.forEach(item => {
    values[item.itemKey] = item.value || '';
  });

  const services = [
    ['2D floor plan', values['SERVICE_2D_FLOOR_PLAN']],
    ['Measurements', values['SERVICE_MEASUREMENTS']],
    ['Drone / aerial', values['SERVICE_DRONE']],
    ['Video', values['SERVICE_VIDEO']],
    ['Twilight', values['SERVICE_TWILIGHT']],
    ['Virtual staging', values['SERVICE_VIRTUAL_STAGING']]
  ]
    .filter(service => service[1] === 'Yes')
    .map(service => service[0]);

  if (values['OTHER_PHOTO_SERVICES']) {
    services.push(values['OTHER_PHOTO_SERVICES']);
  }

  const address = buildPhotoOrderAddress_(tx);
  const seller = [tx['Client First Name'], tx['Client Last Name']]
    .filter(Boolean)
    .join(' ');

  const subject = `Photography Order: ${address || tx['Transaction ID']}`;

  const plainBody = [
    'A photography order was submitted in JBA OS.',
    '',
    `Transaction: ${tx['Transaction ID']}`,
    `Property: ${address || 'Not provided'}`,
    `Seller: ${seller || 'Not provided'}`,
    `Agent: ${tx['Listing Agent'] || 'Not provided'}`,
    `Agent Email: ${tx['Agent Email'] || 'Not provided'}`,
    '',
    `All documents signed: ${values['ALL_DOCS_SIGNED'] || 'Not provided'}`,
    `Photographer: ${formatPhotoOrderPhotographer_(values)}`,
    `Photo date: ${values['PHOTO_DATE'] || 'Not provided'}`,
    `Photo time: ${values['PHOTO_TIME'] || 'Not provided'}`,
    `Additional services: ${services.length ? services.join(', ') : 'None selected'}`,
    `Access: ${values['PHOTOGRAPHER_ACCESS'] || 'Not provided'}`,
    `Access instructions: ${values['ACCESS_INSTRUCTIONS'] || 'None'}`,
    `Agent attending: ${values['AGENT_ATTENDING'] || 'Not provided'}`,
    `Taking lockbox: ${values['TAKING_LOCKBOX'] || 'Not provided'}`,
    `Occupancy: ${values['OCCUPANCY_STATUS'] || 'Not provided'}`,
    `Basement: ${values['BASEMENT'] || 'Not provided'}`,
    `Finished basement sqft: ${values['BASEMENT_FINISHED_SQFT'] || 'Not provided'}`,
    `Notes: ${values['PHOTO_NOTES'] || 'None'}`,
    '',
    `Submitted by: ${user.displayName || user.email}`,
    `Recipient assignment: ${assignedTc ? 'Transaction Coordinator' : 'Operations fallback'}`
  ].join('\n');

  const htmlBody = buildPhotoOrderEmailHtml_({
    tx: tx,
    address: address,
    seller: seller,
    values: values,
    services: services,
    submittedBy: user.displayName || user.email,
    usedOperationsFallback: !assignedTc
  });

  MailApp.sendEmail({
    to: recipient,
    cc: normalizeEmail_(tx['Agent Email']) || '',
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: 'JBA OS'
  });

  appendObject_(JBA_OS.sheets.notifications, {
    'Notification ID': Utilities.getUuid(),
    'Created At': new Date(),
    'Status': 'Sent',
    'Transaction ID': tx['Transaction ID'],
    'Recipient': recipient,
    'Channel': 'Email',
    'Subject': subject,
    'Message': plainBody,
    'Scheduled For': new Date(),
    'Sent At': new Date(),
    'Error': ''
  });

  logActivity_(
    user,
    tx['Transaction ID'],
    'PHOTO_ORDER_SENT_TO_TC',
    tx['Current Stage Name'],
    tx['Current Stage Name'],
    tx['Current Action Name'],
    tx['Current Action Name'],
    `Photography order emailed to ${recipient}`
  );
}

function formatPhotoOrderPhotographer_(values) {
  if (values['PHOTOGRAPHER'] === 'Other' && values['PHOTOGRAPHER_OTHER']) {
    return values['PHOTOGRAPHER_OTHER'];
  }

  return values['PHOTOGRAPHER'] || 'Not provided';
}

function buildPhotoOrderAddress_(tx) {
  const street = [
    tx['Address Line 1'] || tx['Property Address'],
    tx['Address Line 2']
  ].filter(Boolean).join(', ');

  const cityStateZip = [
    tx['City'],
    tx['State'],
    tx['Postal Code']
  ].filter(Boolean).join(' ');

  return [street, cityStateZip].filter(Boolean).join(', ');
}

function buildPhotoOrderEmailHtml_(data) {
  const v = data.values;

  const rows = [
    ['Transaction', data.tx['Transaction ID']],
    ['Property', data.address || 'Not provided'],
    ['Seller', data.seller || 'Not provided'],
    ['Listing Agent', data.tx['Listing Agent'] || 'Not provided'],
    ['Agent Email', data.tx['Agent Email'] || 'Not provided'],
    ['All Docs Signed', v['ALL_DOCS_SIGNED'] || 'Not provided'],
    ['Photographer', formatPhotoOrderPhotographer_(v)],
    ['Photo Date', v['PHOTO_DATE'] || 'Not provided'],
    ['Photo Time', v['PHOTO_TIME'] || 'Not provided'],
    ['Additional Services', data.services.length ? data.services.join(', ') : 'None selected'],
    ['Access Method', v['PHOTOGRAPHER_ACCESS'] || 'Not provided'],
    ['Access Instructions', v['ACCESS_INSTRUCTIONS'] || 'None'],
    ['Agent Attending', v['AGENT_ATTENDING'] || 'Not provided'],
    ['Taking Lockbox', v['TAKING_LOCKBOX'] || 'Not provided'],
    ['Occupancy', v['OCCUPANCY_STATUS'] || 'Not provided'],
    ['Basement', v['BASEMENT'] || 'Not provided'],
    ['Finished Basement Sqft', v['BASEMENT_FINISHED_SQFT'] || 'Not provided'],
    ['Notes', v['PHOTO_NOTES'] || 'None']
  ];

  return `
    <div style="font-family:Arial,Helvetica,sans-serif;background:#f4f5f7;padding:24px;color:#20242a;">
      <div style="max-width:720px;margin:0 auto;background:#ffffff;border:1px solid #e2e5e9;border-radius:14px;overflow:hidden;">
        <div style="background:#17191d;color:#ffffff;padding:22px 26px;">
          <div style="font-size:12px;color:#c7c9cc;">JBA OS · Photography Order</div>
          <div style="font-size:24px;font-weight:700;margin-top:5px;">${escapeHtml_(data.address || 'Address pending')}</div>
        </div>
        <div style="padding:24px 26px;">
          ${data.usedOperationsFallback
            ? '<div style="padding:12px 14px;background:#fff4e5;border-left:4px solid #b79a5b;margin-bottom:18px;">No TC was assigned, so this order was sent to Operations.</div>'
            : ''}
          <table style="width:100%;border-collapse:collapse;">
            ${rows.map(row => `
              <tr>
                <td style="width:190px;padding:10px 8px;border-bottom:1px solid #e2e5e9;color:#6c737d;vertical-align:top;">${escapeHtml_(row[0])}</td>
                <td style="padding:10px 8px;border-bottom:1px solid #e2e5e9;font-weight:600;vertical-align:top;">${escapeHtml_(row[1])}</td>
              </tr>
            `).join('')}
          </table>
          <div style="margin-top:18px;color:#6c737d;font-size:12px;">
            Submitted by ${escapeHtml_(data.submittedBy)} through JBA OS.
          </div>
        </div>
      </div>
    </div>
  `;
}

function validateChecklistAction_(user, action, actionKey) {
  if (!action) {
    throw new Error('No active action is configured.');
  }

  if (action['Action Key'] !== actionKey) {
    throw new Error('This checklist is no longer the current task.');
  }

  if (!canUserPerformAction_(user, action)) {
    throw new Error('You are not authorized to complete this checklist.');
  }
}

function getChecklistTemplateItems_(workflowKey, stageKey, actionKey) {
  return sheetObjects_(JBA_CHECKLIST.templatesSheet)
    .filter(row =>
      row['Workflow Key'] === workflowKey &&
      row['Stage Key'] === stageKey &&
      row['Action Key'] === actionKey &&
      row['Active?'] === 'Yes'
    )
    .sort((a, b) =>
      Number(a['Item Order'] || 0) - Number(b['Item Order'] || 0)
    );
}

function getChecklistResponseMap_(transactionId, actionKey) {
  const map = {};

  sheetObjects_(JBA_CHECKLIST.responsesSheet)
    .filter(row =>
      String(row['Transaction ID'] || '') === String(transactionId || '') &&
      row['Action Key'] === actionKey
    )
    .forEach(row => {
      map[row['Item Key']] = {
        value: row['Value'] || '',
        completed: row['Completed?'] === 'Yes'
      };
    });

  return map;
}

function upsertChecklistResponse_(transactionId, actionKey, template, rawValue, userEmail) {
  const sheet = getDatabase_().getSheetByName(JBA_CHECKLIST.responsesSheet);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0];

  const idCol = headers.indexOf('Transaction ID');
  const actionCol = headers.indexOf('Action Key');
  const itemCol = headers.indexOf('Item Key');

  const cleanValue = normalizeChecklistValue_(template['Item Type'], rawValue);
  const completed = isChecklistItemComplete_(template['Item Type'], cleanValue)
    ? 'Yes'
    : 'No';

  const rowObject = {
    'Transaction ID': transactionId,
    'Action Key': actionKey,
    'Item Key': template['Item Key'],
    'Value': cleanValue,
    'Completed?': completed,
    'Updated At': new Date(),
    'Updated By': userEmail
  };

  const rowIndex = values.slice(1).findIndex(row =>
    row[idCol] === transactionId &&
    row[actionCol] === actionKey &&
    row[itemCol] === template['Item Key']
  );

  if (rowIndex < 0) {
    sheet.appendRow(headers.map(header => rowObject[header] || ''));
    return;
  }

  headers.forEach((header, columnIndex) => {
    sheet.getRange(rowIndex + 2, columnIndex + 1)
      .setValue(rowObject[header] || '');
  });
}

function normalizeChecklistValue_(type, value) {
  if (type === 'CHECKBOX') {
    return value === true || value === 'true' || value === 'Yes'
      ? 'Yes'
      : '';
  }

  return String(value || '').trim();
}

function isChecklistItemComplete_(type, value) {
  if (type === 'CHECKBOX') {
    return value === true || value === 'true' || value === 'Yes';
  }

  return String(value || '').trim() !== '';
}

function calculateChecklistProgress_(items) {
  const requiredItems = items.filter(item => item.required);
  const completedRequired = requiredItems.filter(item =>
    isChecklistItemComplete_(item.type, item.value)
  ).length;

  return {
    totalRequired: requiredItems.length,
    completedRequired: completedRequired,
    percent: requiredItems.length
      ? Math.round((completedRequired / requiredItems.length) * 100)
      : 100
  };
}

function parseChecklistOptions_(value) {
  return String(value || '')
    .split('|')
    .map(option => option.trim())
    .filter(Boolean);
}

/**
 * Optional diagnostic test.
 */
function testNativeChecklistEngine() {
  const transactions = sheetObjects_(JBA_OS.sheets.transactions);
  const tx = transactions.find(row =>
    row['Current Action Key'] === 'PRELIST_CHECKLIST'
  );

  if (!tx) {
    throw new Error('No transaction is currently on PRELIST_CHECKLIST.');
  }

  console.log(JSON.stringify(
    getNativeChecklist(tx['Transaction ID']),
    null,
    2
  ));
}
