/**
 * JBA OS — Task Engine
 *
 * Trigger-driven task creation: when a workflow action completes, create
 * role-assigned tasks configured in the Task Templates sheet.
 *
 * Add this as a new Apps Script file named TaskEngine.gs.
 * Run setupTaskEngine() once to create the Task Templates and Tasks sheets.
 */

const JBA_TASK = {
  templatesSheet: 'Task Templates',
  tasksSheet: 'Tasks',
  roleEmailFields: {
    'Agent': 'Agent Email',
    'Marketing': 'Assigned Marketing Email',
    'Transaction Coordinator': 'Assigned TC Email',
    'Operations Admin': 'Assigned Operations Email'
  }
};

function setupTaskEngine() {
  const ss = getDatabase_();
  setupTaskTemplates_(ss);
  setupTasksSheet_(ss);
  seedDefaultTaskTemplates_();
}

function setupTaskTemplates_(ss) {
  const headers = [
    'Template ID',
    'Workflow Key',
    'Trigger Action Key',
    'Task Order',
    'Role',
    'Task Name',
    'Description',
    'Active?'
  ];
  ensureSheet_(ss, JBA_TASK.templatesSheet, headers, [headers]);
}

function setupTasksSheet_(ss) {
  const headers = [
    'Task ID',
    'Transaction ID',
    'Template ID',
    'Trigger Action Key',
    'Role',
    'Task Name',
    'Status',
    'Assigned Email',
    'Created At',
    'Completed At',
    'Completed By'
  ];
  ensureSheet_(ss, JBA_TASK.tasksSheet, headers, [headers]);

  const sheet = ss.getSheetByName(JBA_TASK.tasksSheet);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Open', 'Complete'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('G2:G10000').setDataValidation(statusRule);
}

function seedDefaultTaskTemplates_() {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(JBA_TASK.templatesSheet);

  if (sheet.getLastRow() === 1) {
    const rows = [
      ['MLS_SUBMISSION_TC_ENTER', 'SELLER_LISTING', 'MLS_SUBMISSION', 10, 'Transaction Coordinator', 'Enter into MLS', '', 'Yes'],
      ['MLS_SUBMISSION_MKT_SOCIAL', 'SELLER_LISTING', 'MLS_SUBMISSION', 20, 'Marketing', 'Social media graphics', '', 'Yes'],
      ['MLS_SUBMISSION_MKT_EMAIL', 'SELLER_LISTING', 'MLS_SUBMISSION', 30, 'Marketing', 'Just Listed email', '', 'Yes'],
      ['MLS_SUBMISSION_MKT_ZILLOW', 'SELLER_LISTING', 'MLS_SUBMISSION', 40, 'Marketing', 'Zillow verification', '', 'Yes'],
      ['MLS_SUBMISSION_AGENT_SIGN', 'SELLER_LISTING', 'MLS_SUBMISSION', 50, 'Agent', 'Install sign', '', 'Yes']
    ];
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function getTaskTemplatesForAction_(workflowKey, actionKey) {
  return sheetObjects_(JBA_TASK.templatesSheet)
    .filter(row =>
      row['Workflow Key'] === workflowKey &&
      row['Trigger Action Key'] === actionKey &&
      row['Active?'] === 'Yes'
    )
    .sort((a, b) => Number(a['Task Order'] || 0) - Number(b['Task Order'] || 0));
}

function resolveTaskAssigneeEmail_(tx, role) {
  const field = JBA_TASK.roleEmailFields[role];
  return field ? normalizeEmail_(tx[field]) : '';
}

/**
 * Creates any not-yet-created tasks configured for this action. Safe to call
 * more than once for the same transaction/action — already-created tasks
 * (matched by Template ID) are skipped.
 */
function createTasksForAction_(user, tx, actionKey) {
  const templates = getTaskTemplatesForAction_(tx['Workflow Key'], actionKey);
  if (!templates.length) return;

  const transactionId = tx['Transaction ID'];
  const existingTemplateIds = sheetObjects_(JBA_TASK.tasksSheet)
    .filter(row =>
      row['Transaction ID'] === transactionId &&
      row['Trigger Action Key'] === actionKey
    )
    .map(row => row['Template ID']);

  const created = [];

  templates.forEach(template => {
    if (existingTemplateIds.includes(template['Template ID'])) return;

    const assignedEmail = resolveTaskAssigneeEmail_(tx, template['Role']);

    appendObject_(JBA_TASK.tasksSheet, {
      'Task ID': Utilities.getUuid(),
      'Transaction ID': transactionId,
      'Template ID': template['Template ID'],
      'Trigger Action Key': actionKey,
      'Role': template['Role'],
      'Task Name': template['Task Name'],
      'Status': 'Open',
      'Assigned Email': assignedEmail,
      'Created At': new Date(),
      'Completed At': '',
      'Completed By': ''
    });

    if (assignedEmail) {
      appendObject_(JBA_OS.sheets.notifications, {
        'Notification ID': Utilities.getUuid(),
        'Created At': new Date(),
        'Status': 'Queued',
        'Transaction ID': transactionId,
        'Recipient': assignedEmail,
        'Channel': 'Email',
        'Subject': `JBA OS Task: ${template['Task Name']}`,
        'Message': `${tx['Property Address']}: ${template['Task Name']} (${template['Role']}).`,
        'Scheduled For': new Date(),
        'Sent At': '',
        'Error': ''
      });
    }

    created.push(template['Task Name']);
  });

  if (created.length) {
    logActivity_(
      user,
      transactionId,
      'TASKS_CREATED',
      tx['Current Stage Name'],
      tx['Current Stage Name'],
      actionKey,
      '',
      `${created.length} task(s) created: ${created.join(', ')}`
    );
  }
}

/**
 * Returns open tasks assigned to the current user's role (all open tasks
 * for Executive Admin / Operations Admin), joined with the transaction's
 * property address.
 */
function getMyTasks() {
  const user = requireUser_();
  const isAdmin = ['Executive Admin', 'Operations Admin'].includes(user.role);

  const tasks = sheetObjects_(JBA_TASK.tasksSheet)
    .filter(row => row['Status'] === 'Open')
    .filter(row => isAdmin || row['Role'] === user.role);

  return tasks
    .map(task => {
      const tx = getTransaction_(task['Transaction ID']) || {};
      return {
        taskId: task['Task ID'],
        transactionId: task['Transaction ID'],
        propertyAddress: tx['Property Address'] || '',
        role: task['Role'],
        taskName: task['Task Name'],
        createdAt: task['Created At']
      };
    })
    .sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
}

function completeTask(taskId) {
  const user = requireUser_();
  const sheet = getDatabase_().getSheetByName(JBA_TASK.tasksSheet);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  const idColumn = headers.indexOf('Task ID');
  const rowIndex = values.slice(1).findIndex(
    row => String(row[idColumn]) === String(taskId)
  );

  if (rowIndex < 0) {
    throw new Error('Task not found: ' + taskId);
  }

  const row = values[rowIndex + 1];
  const task = {};
  headers.forEach((header, index) => task[header] = row[index]);

  const isAdmin = ['Executive Admin', 'Operations Admin'].includes(user.role);
  if (!isAdmin && user.role !== task['Role']) {
    throw new Error('You are not authorized to complete this task.');
  }

  const tx = getTransaction_(task['Transaction ID']);
  if (!tx || !canAccessTransaction_(user, tx)) {
    throw new Error('You do not have access to this transaction.');
  }

  const updates = {
    'Status': 'Complete',
    'Completed At': new Date(),
    'Completed By': user.email
  };

  Object.keys(updates).forEach(header => {
    const columnIndex = headers.indexOf(header);
    if (columnIndex >= 0) {
      sheet.getRange(rowIndex + 2, columnIndex + 1)
        .setValue(updates[header]);
    }
  });

  logActivity_(
    user,
    task['Transaction ID'],
    'TASK_COMPLETED',
    '',
    '',
    '',
    '',
    task['Task Name'] + ' (' + task['Role'] + ')'
  );

  return { success: true };
}

function testTaskEngine() {
  const ss = getDatabase_();
  setupTaskTemplates_(ss);
  setupTasksSheet_(ss);
  seedDefaultTaskTemplates_();

  const templates = getTaskTemplatesForAction_('SELLER_LISTING', 'MLS_SUBMISSION');
  console.log(`Found ${templates.length} MLS_SUBMISSION templates.`);

  const transactions = sheetObjects_(JBA_OS.sheets.transactions);
  if (!transactions.length) {
    console.log('No transactions found — create one before running this test.');
    return;
  }

  const tx = transactions[0];
  const systemUser = { email: 'system@jba-os', role: 'System' };

  createTasksForAction_(systemUser, tx, 'MLS_SUBMISSION');
  createTasksForAction_(systemUser, tx, 'MLS_SUBMISSION');

  const created = sheetObjects_(JBA_TASK.tasksSheet)
    .filter(row => row['Transaction ID'] === tx['Transaction ID']);

  console.log(`Tasks for ${tx['Transaction ID']}: ${created.length} (should equal template count, not double).`);
}
