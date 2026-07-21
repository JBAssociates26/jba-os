/**
 * JBA OS VERSION 1
 * Jim Bunn & Associates
 *
 * CORE MODULES
 * 1. Users & Roles
 * 2. Transactions
 * 3. Workflow Stages
 * 4. Workflow Actions
 * 5. Forms Registry
 * 6. Activity Log
 * 7. Notifications Queue
 * 8. Secure Web Portal
 *
 * SECURITY
 * - Server-side authorization on every request
 * - Agents see only their own transactions
 * - Marketing/TC users see only assigned transactions
 * - Admins see all transactions
 * - No direct spreadsheet access for non-admins
 */

const JBA_OS = {
  name: 'JBA OS',
  version: '1.0',
  databaseName: 'JBA OS Database',
  timezone: 'America/Detroit',
  sheets: {
    users: 'Users',
    agents: 'Agents',
    transactions: 'Transactions',
    workflowStages: 'Workflow Stages',
    workflowActions: 'Workflow Actions',
    forms: 'Forms',
    activity: 'Activity Log',
    notifications: 'Notifications Queue',
    settings: 'Settings'
  },
  roles: [
    'Executive Admin',
    'Operations Admin',
    'Agent',
    'Marketing',
    'Transaction Coordinator',
    'Read Only'
  ],
  transactionStatuses: [
    'Active',
    'Paused',
    'Closed',
    'Lost',
    'Cancelled'
  ]
};

/* =========================================================
   INITIAL SETUP
   ========================================================= */

function setupJBAOSVersion1() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('JBA_OS_DATABASE_ID');
  let ss;

  if (dbId) {
    ss = SpreadsheetApp.openById(dbId);
  } else {
    ss = SpreadsheetApp.create(JBA_OS.databaseName);
    props.setProperty('JBA_OS_DATABASE_ID', ss.getId());
  }

  setupSettings_(ss);
  setupUsers_(ss);
  setupAgents_(ss);
  setupTransactions_(ss);
  setupWorkflowStages_(ss);
  setupWorkflowActions_(ss);
  setupForms_(ss);
  setupActivityLog_(ss);
  setupNotifications_(ss);

  seedDefaultWorkflow_();
  seedDefaultUsers_();

  console.log('JBA OS database: ' + ss.getUrl());
  console.log('Next: review Users, Forms, Workflow Stages, and Workflow Actions.');
}

function setupSettings_(ss) {
  const headers = ['Setting', 'Value', 'Description'];
  const rows = [
    ['Setting', 'Value', 'Description'],
    ['Workspace Domain', 'jimbunnandassociates.com', 'Primary JBA Google Workspace domain'],
    ['Require Workspace Domain for Agents', 'Yes', 'Require JBA email for Agent role'],
    ['Default Marketing Email', 'nexusmarketingedge@gmail.com', 'Default marketing assignee'],
    ['Default Operations Email', 'hello@jimbunnandassociates.com', 'Default operations assignee'],
    ['Portal Title', 'JBA OS', 'Portal display title'],
    ['Company Name', 'Jim Bunn & Associates', 'Company name shown in portal']
  ];
  ensureSheet_(ss, JBA_OS.sheets.settings, headers, rows);
}

function setupUsers_(ss) {
  const headers = [
    'Email',
    'Display Name',
    'Role',
    'Agent Email Match',
    'Active?',
    'Can View Financials?',
    'Can Change Stage?',
    'Can Reassign?',
    'Notes'
  ];
  ensureSheet_(ss, JBA_OS.sheets.users, headers, [headers]);

  const sheet = ss.getSheetByName(JBA_OS.sheets.users);
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(JBA_OS.roles, true)
    .setAllowInvalid(false)
    .build();

  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange('C2:C1000').setDataValidation(roleRule);
  sheet.getRange('E2:H1000').setDataValidation(yesNoRule);
}

function setupAgents_(ss) {
  const headers = [
    'Agent ID',
    'Display Name',
    'First Name',
    'Last Name',
    'Email',
    'Phone',
    'Status',
    'Team',
    'Start Date',
    'Notes'
  ];
  ensureSheet_(ss, JBA_OS.sheets.agents, headers, [headers]);
}

function setupTransactions_(ss) {
  const headers = [
    'Transaction ID',
    'Created At',
    'Updated At',
    'Status',
    'Workflow Key',
    'Current Stage Key',
    'Current Stage Name',
    'Current Action Key',
    'Current Action Name',
    'Listing Agent',
    'Agent Email',
    'Assigned Operations Email',
    'Assigned Marketing Email',
    'Assigned TC Email',
    'Client First Name',
    'Client Last Name',
    'Client Email',
    'Client Phone',
    'Property Address',
    'Address Line 1',
    'Address Line 2',
    'City',
    'State',
    'Postal Code',
    'County',
    'Property Type',
    'Appointment Date',
    'Appointment Time',
    'Anticipated List Price',
    'Actual List Price',
    'Contract Price',
    'Closing Date',
    'Needs Review?',
    'Review Reasons',
    'Urgency',
    'Binder Needed?',
    'Binder Status',
    'Notes',
    'Last Action By',
    'Source System',
    'Source Record ID'
  ];
  ensureSheet_(ss, JBA_OS.sheets.transactions, headers, [headers]);

  const sheet = ss.getSheetByName(JBA_OS.sheets.transactions);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(JBA_OS.transactionStatuses, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D10000').setDataValidation(statusRule);
}

function setupWorkflowStages_(ss) {
  const headers = [
    'Workflow Key',
    'Stage Key',
    'Stage Name',
    'Stage Order',
    'Active?',
    'Description'
  ];
  ensureSheet_(ss, JBA_OS.sheets.workflowStages, headers, [headers]);
}

function setupWorkflowActions_(ss) {
  const headers = [
    'Workflow Key',
    'Stage Key',
    'Action Key',
    'Action Name',
    'Action Order',
    'Action Type',
    'Form Key',
    'Assigned Role',
    'Required?',
    'Next Stage Key',
    'Active?',
    'Description'
  ];
  ensureSheet_(ss, JBA_OS.sheets.workflowActions, headers, [headers]);
}

function setupForms_(ss) {
  const headers = [
    'Form Key',
    'Form Name',
    'Google Form ID',
    'Active?',
    'Transaction ID Question',
    'Agent Question',
    'Client First Name Question',
    'Client Last Name Question',
    'Property Address Question',
    'City Question',
    'Postal Code Question',
    'County Question',
    'Property Type Question',
    'Notes'
  ];
  ensureSheet_(ss, JBA_OS.sheets.forms, headers, [headers]);
}

function setupActivityLog_(ss) {
  const headers = [
    'Timestamp',
    'User Email',
    'User Role',
    'Transaction ID',
    'Action',
    'Previous Stage',
    'New Stage',
    'Previous Action',
    'New Action',
    'Details'
  ];
  ensureSheet_(ss, JBA_OS.sheets.activity, headers, [headers]);
}

function setupNotifications_(ss) {
  const headers = [
    'Notification ID',
    'Created At',
    'Status',
    'Transaction ID',
    'Recipient',
    'Channel',
    'Subject',
    'Message',
    'Scheduled For',
    'Sent At',
    'Error'
  ];
  ensureSheet_(ss, JBA_OS.sheets.notifications, headers, [headers]);
}

function ensureSheet_(ss, name, headers, initialRows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, initialRows.length, initialRows[0].length)
      .setValues(initialRows);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function seedDefaultUsers_() {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(JBA_OS.sheets.users);
  if (sheet.getLastRow() > 1) return;

  sheet.getRange(2, 1, 3, 9).setValues([
    [
      'jim@jimbunnandassociates.com',
      'Jim Bunn',
      'Executive Admin',
      '',
      'Yes',
      'Yes',
      'Yes',
      'Yes',
      'Full access'
    ],
    [
      'hello@jimbunnandassociates.com',
      'JBA Operations',
      'Operations Admin',
      '',
      'Yes',
      'Yes',
      'Yes',
      'Yes',
      'Operations access'
    ],
    [
      'nexusmarketingedge@gmail.com',
      'TJ / Nexus Marketing',
      'Marketing',
      '',
      'Yes',
      'No',
      'No',
      'No',
      'Assigned marketing transactions only'
    ]
  ]);
}

function seedDefaultWorkflow_() {
  const ss = getDatabase_();
  const stageSheet = ss.getSheetByName(JBA_OS.sheets.workflowStages);
  const actionSheet = ss.getSheetByName(JBA_OS.sheets.workflowActions);
  const formSheet = ss.getSheetByName(JBA_OS.sheets.forms);

  if (stageSheet.getLastRow() === 1) {
    const stages = [
      ['SELLER_LISTING', 'PRE_LISTING', 'Pre-Listing', 10, 'Yes', 'Prepare for listing appointment'],
      ['SELLER_LISTING', 'APPOINTMENT', 'Listing Appointment', 20, 'Yes', 'Appointment scheduled or completed'],
      ['SELLER_LISTING', 'LISTING_SECURED', 'Listing Secured', 30, 'Yes', 'Listing agreement signed'],
      ['SELLER_LISTING', 'PHOTOS', 'Photography', 40, 'Yes', 'Photography ordered and completed'],
      ['SELLER_LISTING', 'MLS_PREP', 'MLS Preparation', 50, 'Yes', 'Listing information and assets prepared'],
      ['SELLER_LISTING', 'READY_LIVE', 'Ready to Go Live', 60, 'Yes', 'Final review before activation'],
      ['SELLER_LISTING', 'LIVE', 'Live', 70, 'Yes', 'Listing active'],
      ['SELLER_LISTING', 'UNDER_CONTRACT', 'Under Contract', 80, 'Yes', 'Accepted offer and contract workflow'],
      ['SELLER_LISTING', 'CLOSING', 'Closing', 90, 'Yes', 'Closing preparation'],
      ['SELLER_LISTING', 'CLOSED', 'Closed', 100, 'Yes', 'Transaction completed']
    ];
    stageSheet.getRange(2, 1, stages.length, 6).setValues(stages);
  }

  if (actionSheet.getLastRow() === 1) {
    const actions = [
      ['SELLER_LISTING', 'PRE_LISTING', 'PRELIST_CHECKLIST', 'Complete Pre-Listing Checklist', 10, 'FORM', 'PRELIST', 'Agent', 'Yes', 'APPOINTMENT', 'Yes', 'Complete before listing appointment'],
      ['SELLER_LISTING', 'APPOINTMENT', 'RECORD_OUTCOME', 'Record Listing Appointment Outcome', 10, 'INTERNAL', '', 'Agent', 'Yes', 'LISTING_SECURED', 'Yes', 'Record whether listing was secured'],
      ['SELLER_LISTING', 'LISTING_SECURED', 'PHOTO_ORDER', 'Order Photography', 10, 'FORM', 'PHOTO_ORDER', 'Agent', 'Yes', 'PHOTOS', 'Yes', 'Submit photography order'],
      ['SELLER_LISTING', 'PHOTOS', 'CONFIRM_PHOTOS', 'Confirm Photos Complete', 10, 'INTERNAL', '', 'Marketing', 'Yes', 'MLS_PREP', 'Yes', 'Confirm media delivery'],
      ['SELLER_LISTING', 'MLS_PREP', 'MLS_SUBMISSION', 'Complete MLS Submission', 10, 'FORM', 'MLS_SUBMISSION', 'Agent', 'Yes', 'READY_LIVE', 'Yes', 'Submit all listing data'],
      ['SELLER_LISTING', 'READY_LIVE', 'FINAL_REVIEW', 'Final Listing Review', 10, 'INTERNAL', '', 'Operations Admin', 'Yes', 'LIVE', 'Yes', 'Approve listing activation'],
      ['SELLER_LISTING', 'LIVE', 'MARKETING_LAUNCH', 'Launch Listing Marketing', 10, 'INTERNAL', '', 'Marketing', 'Yes', 'LIVE', 'Yes', 'Complete marketing launch'],
      ['SELLER_LISTING', 'LIVE', 'ACCEPTED_OFFER', 'Record Accepted Offer', 20, 'FORM', 'UNDER_CONTRACT', 'Agent', 'Yes', 'UNDER_CONTRACT', 'Yes', 'Start transaction coordination'],
      ['SELLER_LISTING', 'UNDER_CONTRACT', 'CONTRACT_CHECKLIST', 'Complete Under Contract Checklist', 10, 'FORM', 'UNDER_CONTRACT', 'Transaction Coordinator', 'Yes', 'CLOSING', 'Yes', 'Track contract-to-close'],
      ['SELLER_LISTING', 'CLOSING', 'CLOSING_CHECKLIST', 'Complete Closing Checklist', 10, 'FORM', 'CLOSING', 'Transaction Coordinator', 'Yes', 'CLOSED', 'Yes', 'Finalize closing'],
      ['SELLER_LISTING', 'CLOSED', 'ARCHIVE', 'Archive Transaction', 10, 'INTERNAL', '', 'Operations Admin', 'No', 'CLOSED', 'Yes', 'Close out record']
    ];
    actionSheet.getRange(2, 1, actions.length, 12).setValues(actions);
  }

  if (formSheet.getLastRow() === 1) {
    const forms = [
      ['PRELIST', 'Checklist #1 — Pre-Listing', '', 'Yes', 'Transaction ID', 'Agent', 'First Name', 'Last Name', 'Property Address', 'City', 'Postal Code', 'County', 'Type of Property', 'Paste existing Checklist #1 Form ID'],
      ['PHOTO_ORDER', 'Checklist #2 — Photo Order', '', 'Yes', 'Transaction ID', 'Agent', 'First Name', 'Last Name', 'Property Address', 'City', 'Postal Code', 'County', 'Type of Property', 'Paste existing Checklist #2 Form ID'],
      ['MLS_SUBMISSION', 'Checklist #3 — MLS Submission', '', 'No', 'Transaction ID', 'Agent', 'First Name', 'Last Name', 'Property Address', 'City', 'Postal Code', 'County', 'Type of Property', 'Create later'],
      ['UNDER_CONTRACT', 'Checklist #4 — Under Contract', '', 'No', 'Transaction ID', 'Agent', 'First Name', 'Last Name', 'Property Address', 'City', 'Postal Code', 'County', 'Type of Property', 'Create later'],
      ['CLOSING', 'Checklist #5 — Closing', '', 'No', 'Transaction ID', 'Agent', 'First Name', 'Last Name', 'Property Address', 'City', 'Postal Code', 'County', 'Type of Property', 'Create later']
    ];
    formSheet.getRange(2, 1, forms.length, 14).setValues(forms);
  }
}

/* =========================================================
   AUTHORIZATION
   ========================================================= */

function requireUser_() {
  const email = normalizeEmail_(Session.getActiveUser().getEmail());

  if (!email) {
    throw new Error(
      'Your Google account could not be verified. Sign in with an authorized account.'
    );
  }

  const user = getUserByEmail_(email);
  if (!user) {
    throw new Error('This account is not authorized for JBA OS.');
  }
  if (user.active !== 'Yes') {
    throw new Error('This JBA OS account is inactive.');
  }

  if (user.role === 'Agent') {
    const requireDomain =
      getSetting_('Require Workspace Domain for Agents') === 'Yes';
    const domain = String(getSetting_('Workspace Domain') || '').toLowerCase();
    if (requireDomain && !email.endsWith('@' + domain)) {
      throw new Error('Agents must use their JBA Workspace account.');
    }
  }

  return user;
}

function getUserByEmail_(email) {
  const rows = sheetObjects_(JBA_OS.sheets.users);
  const match = rows.find(row =>
    normalizeEmail_(row['Email']) === normalizeEmail_(email)
  );
  if (!match) return null;

  return {
    email: normalizeEmail_(match['Email']),
    displayName: match['Display Name'],
    role: match['Role'],
    agentEmailMatch: normalizeEmail_(
      match['Agent Email Match'] || match['Email']
    ),
    active: match['Active?'],
    canViewFinancials: match['Can View Financials?'] === 'Yes',
    canChangeStage: match['Can Change Stage?'] === 'Yes',
    canReassign: match['Can Reassign?'] === 'Yes'
  };
}

function canAccessTransaction_(user, tx) {
  if (['Executive Admin', 'Operations Admin'].includes(user.role)) return true;

  if (user.role === 'Agent') {
    return normalizeEmail_(tx['Agent Email']) === user.agentEmailMatch;
  }

  if (user.role === 'Marketing') {
    return normalizeEmail_(tx['Assigned Marketing Email']) === user.email;
  }

  if (user.role === 'Transaction Coordinator') {
    return normalizeEmail_(tx['Assigned TC Email']) === user.email;
  }

  if (user.role === 'Read Only') {
    return false;
  }

  return false;
}

function getAuthorizedTransaction_(transactionId) {
  const user = requireUser_();
  const tx = getTransaction_(transactionId);
  if (!tx) throw new Error('Transaction not found.');

  if (!canAccessTransaction_(user, tx)) {
    logActivity_(user, transactionId, 'ACCESS_DENIED', '', '', '', '', 'Unauthorized transaction access attempt');
    throw new Error('You do not have access to this transaction.');
  }

  return { user, tx };
}

/* =========================================================
   WEB APP
   ========================================================= */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('JBA OS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getPortalData() {
  const user = requireUser_();
  const transactions = getAuthorizedTransactions_(user);

  logActivity_(user, '', 'PORTAL_OPENED', '', '', '', '', `${transactions.length} transactions returned`);

  return {
    companyName: getSetting_('Company Name') || 'Jim Bunn & Associates',
    portalTitle: getSetting_('Portal Title') || 'JBA OS',
    user: {
      email: user.email,
      displayName: user.displayName || user.email,
      role: user.role,
      canChangeStage: user.canChangeStage,
      canReassign: user.canReassign
    },
    stats: calculateStats_(transactions),
    transactions: transactions.map(tx => toClientTransaction_(tx, user))
  };
}

function getAuthorizedTransactions_(user) {
  return sheetObjects_(JBA_OS.sheets.transactions)
    .filter(tx => canAccessTransaction_(user, tx));
}

function toClientTransaction_(tx, user) {
  const next = getCurrentActionForTransaction_(tx);

  return {
    transactionId: tx['Transaction ID'],
    status: tx['Status'],
    workflowKey: tx['Workflow Key'],
    currentStageKey: tx['Current Stage Key'],
    currentStageName: tx['Current Stage Name'],
    currentActionKey: next ? next['Action Key'] : '',
    currentActionName: next ? next['Action Name'] : 'No action configured',
    actionType: next ? next['Action Type'] : '',
    formKey: next ? next['Form Key'] : '',
    assignedRole: next ? next['Assigned Role'] : '',
    listingAgent: tx['Listing Agent'],
    agentEmail: tx['Agent Email'],
    clientName: [tx['Client First Name'], tx['Client Last Name']].filter(Boolean).join(' '),
    propertyAddress: tx['Property Address'],
    city: tx['City'],
    appointmentDate: tx['Appointment Date'],
    appointmentTime: tx['Appointment Time'],
    urgency: tx['Urgency'],
    binderStatus: tx['Binder Status'],
    needsReview: user.role === 'Agent' ? false : tx['Needs Review?'] === 'Yes',
    reviewReasons: user.canViewFinancials ? tx['Review Reasons'] : '',
    canLaunchAction: Boolean(next),
    canCompleteInternalAction:
      Boolean(next) &&
      next['Action Type'] === 'INTERNAL' &&
      canUserPerformAction_(user, next),
    canOpenForm:
      Boolean(next) &&
      next['Action Type'] === 'FORM' &&
      canUserPerformAction_(user, next)
  };
}

function launchCurrentAction(transactionId) {
  const auth = getAuthorizedTransaction_(transactionId);
  const action = getCurrentActionForTransaction_(auth.tx);

  if (!action) throw new Error('No active action is configured.');
  if (!canUserPerformAction_(auth.user, action)) {
    throw new Error('You are not authorized to perform this action.');
  }
  if (action['Action Type'] !== 'FORM') {
    throw new Error('The current action is not a form.');
  }

  const form = getFormConfig_(action['Form Key']);
  if (!form || form['Active?'] !== 'Yes') {
    throw new Error('The required form is not active or configured.');
  }

  const formId = String(form['Google Form ID'] || '').trim();
  if (!formId) {
    throw new Error(`No Google Form ID is configured for ${form['Form Name']}.`);
  }

  const url = buildPrefilledFormUrl_(formId, form, auth.tx);

  logActivity_(
    auth.user,
    transactionId,
    'FORM_OPENED',
    auth.tx['Current Stage Name'],
    auth.tx['Current Stage Name'],
    auth.tx['Current Action Name'],
    action['Action Name'],
    form['Form Name']
  );

  return url;
}

function completeInternalAction(transactionId, notes) {
  const auth = getAuthorizedTransaction_(transactionId);
  const action = getCurrentActionForTransaction_(auth.tx);

  if (!action) throw new Error('No active action is configured.');
  if (action['Action Type'] !== 'INTERNAL') {
    throw new Error('The current action must be completed through its form.');
  }
  if (!canUserPerformAction_(auth.user, action)) {
    throw new Error('You are not authorized to complete this action.');
  }

  return advanceTransaction_(auth.user, auth.tx, action, notes || '');
}

function canUserPerformAction_(user, action) {
  if (['Executive Admin', 'Operations Admin'].includes(user.role)) return true;
  return user.role === action['Assigned Role'];
}

/* =========================================================
   WORKFLOW ENGINE
   ========================================================= */

function getCurrentActionForTransaction_(tx) {
  const workflowKey = tx['Workflow Key'];
  const stageKey = tx['Current Stage Key'];

  const actions = sheetObjects_(JBA_OS.sheets.workflowActions)
    .filter(row =>
      row['Workflow Key'] === workflowKey &&
      row['Stage Key'] === stageKey &&
      row['Active?'] === 'Yes'
    )
    .sort((a, b) =>
      Number(a['Action Order'] || 0) - Number(b['Action Order'] || 0)
    );

  if (!actions.length) return null;

  const currentActionKey = tx['Current Action Key'];
  if (currentActionKey) {
    const exact = actions.find(row => row['Action Key'] === currentActionKey);
    if (exact) return exact;
  }

  return actions[0];
}

function advanceTransaction_(user, tx, action, details) {
  const nextStageKey = action['Next Stage Key'] || tx['Current Stage Key'];
  const nextStage = getStage_(tx['Workflow Key'], nextStageKey);

  if (!nextStage) {
    throw new Error(`Next stage is not configured: ${nextStageKey}`);
  }

  const nextAction = getFirstActionForStage_(
    tx['Workflow Key'],
    nextStageKey
  );

  updateTransactionFields_(tx['Transaction ID'], {
    'Updated At': new Date(),
    'Current Stage Key': nextStageKey,
    'Current Stage Name': nextStage['Stage Name'],
    'Current Action Key': nextAction ? nextAction['Action Key'] : '',
    'Current Action Name': nextAction ? nextAction['Action Name'] : '',
    'Last Action By': user.email,
    'Status': nextStageKey === 'CLOSED' ? 'Closed' : tx['Status']
  });

  logActivity_(
    user,
    tx['Transaction ID'],
    'ACTION_COMPLETED',
    tx['Current Stage Name'],
    nextStage['Stage Name'],
    action['Action Name'],
    nextAction ? nextAction['Action Name'] : '',
    details
  );

  queueDefaultNotifications_(
    tx['Transaction ID'],
    nextStage['Stage Name'],
    nextAction ? nextAction['Action Name'] : ''
  );

  createTasksForAction_(user, tx, action['Action Key']);

  return {
    success: true,
    transactionId: tx['Transaction ID'],
    stageName: nextStage['Stage Name'],
    actionName: nextAction ? nextAction['Action Name'] : ''
  };
}

function completeWorkflowFormAction(transactionId, formKey, submissionNotes) {
  const tx = getTransaction_(transactionId);
  if (!tx) throw new Error('Transaction not found.');

  const action = getCurrentActionForTransaction_(tx);
  if (!action) throw new Error('No action configured.');
  if (action['Form Key'] !== formKey) {
    throw new Error('Submitted form does not match the current action.');
  }

  const systemUser = {
    email: 'system@jba-os',
    role: 'System'
  };

  return advanceTransaction_(
    systemUser,
    tx,
    action,
    submissionNotes || 'Form submitted'
  );
}

function getStage_(workflowKey, stageKey) {
  return sheetObjects_(JBA_OS.sheets.workflowStages).find(row =>
    row['Workflow Key'] === workflowKey &&
    row['Stage Key'] === stageKey &&
    row['Active?'] === 'Yes'
  );
}

function getFirstActionForStage_(workflowKey, stageKey) {
  return sheetObjects_(JBA_OS.sheets.workflowActions)
    .filter(row =>
      row['Workflow Key'] === workflowKey &&
      row['Stage Key'] === stageKey &&
      row['Active?'] === 'Yes'
    )
    .sort((a, b) =>
      Number(a['Action Order'] || 0) - Number(b['Action Order'] || 0)
    )[0] || null;
}

/* =========================================================
   TRANSACTION CREATION / UPDATES
   ========================================================= */

function createSellerListingTransaction(payload) {
  const user = requireUser_();

  if (!['Executive Admin', 'Operations Admin', 'Agent'].includes(user.role)) {
    throw new Error('You are not authorized to create transactions.');
  }

  const transactionId =
    payload.transactionId ||
    createTransactionId_();

  const stage = getStage_('SELLER_LISTING', 'PRE_LISTING');
  const action = getFirstActionForStage_('SELLER_LISTING', 'PRE_LISTING');

  const row = {
    'Transaction ID': transactionId,
    'Created At': new Date(),
    'Updated At': new Date(),
    'Status': 'Active',
    'Workflow Key': 'SELLER_LISTING',
    'Current Stage Key': stage['Stage Key'],
    'Current Stage Name': stage['Stage Name'],
    'Current Action Key': action ? action['Action Key'] : '',
    'Current Action Name': action ? action['Action Name'] : '',
    'Listing Agent': payload.listingAgent || user.displayName,
    'Agent Email': normalizeEmail_(payload.agentEmail || user.agentEmailMatch || user.email),
    'Assigned Operations Email': normalizeEmail_(payload.assignedOperationsEmail || getSetting_('Default Operations Email')),
    'Assigned Marketing Email': normalizeEmail_(payload.assignedMarketingEmail || getSetting_('Default Marketing Email')),
    'Assigned TC Email': normalizeEmail_(payload.assignedTCEmail),
    'Client First Name': payload.clientFirstName || '',
    'Client Last Name': payload.clientLastName || '',
    'Client Email': payload.clientEmail || '',
    'Client Phone': payload.clientPhone || '',
    'Property Address': payload.propertyAddress || '',
    'Address Line 1': payload.addressLine1 || payload.propertyAddress || '',
    'Address Line 2': payload.addressLine2 || '',
    'City': payload.city || '',
    'State': payload.state || 'MI',
    'Postal Code': payload.postalCode || '',
    'County': payload.county || '',
    'Property Type': payload.propertyType || '',
    'Appointment Date': payload.appointmentDate || '',
    'Appointment Time': payload.appointmentTime || '',
    'Anticipated List Price': payload.anticipatedListPrice || '',
    'Actual List Price': '',
    'Contract Price': '',
    'Closing Date': '',
    'Needs Review?': payload.needsReview ? 'Yes' : 'No',
    'Review Reasons': payload.reviewReasons || '',
    'Urgency': payload.urgency || '',
    'Binder Needed?': payload.binderNeeded || '',
    'Binder Status': payload.binderStatus || '',
    'Notes': payload.notes || '',
    'Last Action By': user.email,
    'Source System': payload.sourceSystem || 'JBA OS',
    'Source Record ID': payload.sourceRecordId || ''
  };

  appendObject_(JBA_OS.sheets.transactions, row);

  logActivity_(
    user,
    transactionId,
    'TRANSACTION_CREATED',
    '',
    stage['Stage Name'],
    '',
    action ? action['Action Name'] : '',
    row['Property Address']
  );

  return {
    success: true,
    transactionId: transactionId
  };
}

function updateTransactionAssignment(transactionId, assignment) {
  const auth = getAuthorizedTransaction_(transactionId);
  if (!auth.user.canReassign) {
    throw new Error('You do not have permission to reassign transactions.');
  }

  const updates = {};
  if ('agentEmail' in assignment) updates['Agent Email'] = normalizeEmail_(assignment.agentEmail);
  if ('listingAgent' in assignment) updates['Listing Agent'] = assignment.listingAgent;
  if ('operationsEmail' in assignment) updates['Assigned Operations Email'] = normalizeEmail_(assignment.operationsEmail);
  if ('marketingEmail' in assignment) updates['Assigned Marketing Email'] = normalizeEmail_(assignment.marketingEmail);
  if ('tcEmail' in assignment) updates['Assigned TC Email'] = normalizeEmail_(assignment.tcEmail);
  updates['Updated At'] = new Date();
  updates['Last Action By'] = auth.user.email;

  updateTransactionFields_(transactionId, updates);

  logActivity_(
    auth.user,
    transactionId,
    'TRANSACTION_REASSIGNED',
    auth.tx['Current Stage Name'],
    auth.tx['Current Stage Name'],
    auth.tx['Current Action Name'],
    auth.tx['Current Action Name'],
    JSON.stringify(assignment)
  );

  return { success: true };
}

/* =========================================================
   FORM REGISTRY / PREFILL
   ========================================================= */

function getFormConfig_(formKey) {
  return sheetObjects_(JBA_OS.sheets.forms).find(row =>
    row['Form Key'] === formKey
  );
}

function buildPrefilledFormUrl_(formId, config, tx) {
  const form = FormApp.openById(formId);
  let response = form.createResponse();

  const mappings = [
    [config['Transaction ID Question'], tx['Transaction ID']],
    [config['Agent Question'], tx['Listing Agent']],
    [config['Client First Name Question'], tx['Client First Name']],
    [config['Client Last Name Question'], tx['Client Last Name']],
    [config['Property Address Question'], tx['Address Line 1'] || tx['Property Address']],
    [config['City Question'], tx['City']],
    [config['Postal Code Question'], tx['Postal Code']],
    [config['County Question'], tx['County']],
    [config['Property Type Question'], tx['Property Type']]
  ];

  mappings.forEach(([title, value]) => {
    response = addFormResponse_(response, form, title, value);
  });

  return response.toPrefilledUrl();
}

function addFormResponse_(response, form, title, value) {
  const cleanTitle = String(title || '').trim();
  const cleanValue = String(value || '').trim();
  if (!cleanTitle || !cleanValue) return response;

  const item = form.getItems().find(i =>
    String(i.getTitle() || '').trim() === cleanTitle
  );
  if (!item) return response;

  try {
    switch (item.getType()) {
      case FormApp.ItemType.TEXT:
        return response.withItemResponse(
          item.asTextItem().createResponse(cleanValue)
        );
      case FormApp.ItemType.PARAGRAPH_TEXT:
        return response.withItemResponse(
          item.asParagraphTextItem().createResponse(cleanValue)
        );
      case FormApp.ItemType.LIST:
        return response.withItemResponse(
          item.asListItem().createResponse(cleanValue)
        );
      case FormApp.ItemType.MULTIPLE_CHOICE:
        return response.withItemResponse(
          item.asMultipleChoiceItem().createResponse(cleanValue)
        );
      default:
        return response;
    }
  } catch (error) {
    console.warn(`Could not prefill "${cleanTitle}": ${error.message}`);
    return response;
  }
}

/* =========================================================
   NOTIFICATIONS
   ========================================================= */

function queueDefaultNotifications_(transactionId, stageName, actionName) {
  const tx = getTransaction_(transactionId);
  if (!tx) return;

  const recipients = [
    tx['Agent Email'],
    tx['Assigned Operations Email']
  ].filter(Boolean);

  recipients.forEach(recipient => {
    appendObject_(JBA_OS.sheets.notifications, {
      'Notification ID': Utilities.getUuid(),
      'Created At': new Date(),
      'Status': 'Queued',
      'Transaction ID': transactionId,
      'Recipient': recipient,
      'Channel': 'Email',
      'Subject': `JBA OS: ${stageName}`,
      'Message': `${tx['Property Address']} moved to ${stageName}. Next action: ${actionName || 'None'}.`,
      'Scheduled For': new Date(),
      'Sent At': '',
      'Error': ''
    });
  });
}

function processNotificationQueue() {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(JBA_OS.sheets.notifications);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const statusCol = headers.indexOf('Status');
  const recipientCol = headers.indexOf('Recipient');
  const subjectCol = headers.indexOf('Subject');
  const messageCol = headers.indexOf('Message');
  const sentAtCol = headers.indexOf('Sent At');
  const errorCol = headers.indexOf('Error');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][statusCol]) !== 'Queued') continue;

    try {
      MailApp.sendEmail(
        String(data[i][recipientCol]),
        String(data[i][subjectCol]),
        String(data[i][messageCol])
      );
      sheet.getRange(i + 1, statusCol + 1).setValue('Sent');
      sheet.getRange(i + 1, sentAtCol + 1).setValue(new Date());
    } catch (error) {
      sheet.getRange(i + 1, statusCol + 1).setValue('Error');
      sheet.getRange(i + 1, errorCol + 1).setValue(error.message);
    }
  }
}

/* =========================================================
   HELPERS
   ========================================================= */

function getDatabase_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty('JBA_OS_DATABASE_ID');
  if (!id) throw new Error('Run setupJBAOSVersion1() first.');
  return SpreadsheetApp.openById(id);
}

function sheetObjects_(sheetName) {
  const sheet = getDatabase_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0];

  return values.slice(1)
    .filter(row => row.some(value => String(value || '').trim()))
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => obj[header] = row[index]);
      return obj;
    });
}

function getTransaction_(transactionId) {
  return sheetObjects_(JBA_OS.sheets.transactions).find(row =>
    String(row['Transaction ID'] || '').trim() ===
    String(transactionId || '').trim()
  );
}

function appendObject_(sheetName, object) {
  const sheet = getDatabase_().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  sheet.appendRow(headers.map(header => object[header] ?? ''));
}

function updateTransactionFields_(transactionId, updates) {
  const sheet = getDatabase_().getSheetByName(JBA_OS.sheets.transactions);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0];
  const idCol = headers.indexOf('Transaction ID');

  const rowIndex = values.slice(1).findIndex(row =>
    String(row[idCol] || '').trim() === String(transactionId || '').trim()
  );

  if (rowIndex < 0) throw new Error('Transaction not found.');

  Object.keys(updates).forEach(key => {
    const col = headers.indexOf(key);
    if (col >= 0) {
      sheet.getRange(rowIndex + 2, col + 1).setValue(updates[key]);
    }
  });
}

function getSetting_(key) {
  const match = sheetObjects_(JBA_OS.sheets.settings).find(row =>
    row['Setting'] === key
  );
  return match ? match['Value'] : '';
}

function calculateStats_(transactions) {
  return {
    total: transactions.length,
    preListing: transactions.filter(tx =>
      ['PRE_LISTING', 'APPOINTMENT'].includes(tx['Current Stage Key'])
    ).length,
    photos: transactions.filter(tx =>
      ['LISTING_SECURED', 'PHOTOS'].includes(tx['Current Stage Key'])
    ).length,
    live: transactions.filter(tx =>
      tx['Current Stage Key'] === 'LIVE'
    ).length,
    underContract: transactions.filter(tx =>
      tx['Current Stage Key'] === 'UNDER_CONTRACT'
    ).length,
    needsReview: transactions.filter(tx =>
      tx['Needs Review?'] === 'Yes'
    ).length
  };
}

function logActivity_(user, transactionId, action, previousStage, newStage, previousAction, newAction, details) {
  appendObject_(JBA_OS.sheets.activity, {
    'Timestamp': new Date(),
    'User Email': user ? user.email : '',
    'User Role': user ? user.role : '',
    'Transaction ID': transactionId || '',
    'Action': action || '',
    'Previous Stage': previousStage || '',
    'New Stage': newStage || '',
    'Previous Action': previousAction || '',
    'New Action': newAction || '',
    'Details': details || ''
  });
}

function createTransactionId_() {
  const date = Utilities.formatDate(new Date(), JBA_OS.timezone, 'yyyyMMdd');
  return `JBA-${date}-${Utilities.getUuid().slice(0, 8).toUpperCase()}`;
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function testJBAOSAccess() {
  const user = requireUser_();
  const count = getAuthorizedTransactions_(user).length;
  console.log(`Authorized user: ${user.email}`);
  console.log(`Role: ${user.role}`);
  console.log(`Accessible transactions: ${count}`);
}
