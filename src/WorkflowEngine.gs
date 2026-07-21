/**
 * JBA OS — Workflow Engine
 *
 * Central routing for sequential and parallel workflow actions.
 *
 * Add this as a new Apps Script file named WorkflowEngine.gs.
 */

const JBA_WORKFLOW_ENGINE = {
  parallelGroups: {
    LISTING_PREPARATION: {
      workflowKey: 'SELLER_LISTING',
      actionKeys: ['PHOTO_ORDER', 'MLS_SUBMISSION']
    }
  }
};

/**
 * Returns true when a parallel action is already complete.
 */
function isParallelWorkflowActionComplete_(transactionId, actionKey) {
  const statusMap = getParallelActionStatusMap_(transactionId);
  return statusMap[actionKey] === 'Complete';
}

/**
 * Marks a parallel action complete, selects the next unfinished action,
 * or advances the transaction when the entire parallel group is complete.
 */
function completeParallelWorkflowAction_(user, tx, action, checklist) {
  const transactionId = tx['Transaction ID'];
  const actionKey = action['Action Key'];

  markParallelActionComplete_(
    transactionId,
    actionKey,
    user.email
  );

  return routeParallelWorkflow_(user, tx, action, checklist);
}

/**
 * Re-evaluates the parallel group without sending emails or resubmitting a
 * checklist. This is useful for repairing transactions completed before the
 * centralized workflow routing was installed.
 */
function reconcileParallelWorkflowForTransaction(transactionId) {
  const auth = getAuthorizedTransaction_(transactionId);
  const tx = auth.tx;
  const currentActionKey = tx['Current Action Key'] || 'MLS_SUBMISSION';
  const action = getChecklistActionByKey_(tx, currentActionKey);

  return routeParallelWorkflow_(
    auth.user,
    tx,
    action,
    null
  );
}

/**
 * The only function that decides what happens after a parallel action.
 */
function routeParallelWorkflow_(user, tx, completedAction, checklist) {
  const transactionId = tx['Transaction ID'];
  const group = getParallelWorkflowGroup_(tx);

  if (!group) {
    throw new Error(
      'No parallel workflow group is configured for this transaction.'
    );
  }

  const statusMap = getParallelActionStatusMap_(transactionId);
  const remainingActionKeys = group.actionKeys.filter(
    actionKey => statusMap[actionKey] !== 'Complete'
  );

  if (remainingActionKeys.length === 0) {
    logActivity_(
      user,
      transactionId,
      'PARALLEL_GROUP_COMPLETED',
      tx['Current Stage Name'],
      tx['Current Stage Name'],
      completedAction ? completedAction['Action Name'] : '',
      '',
      'All listing-preparation workstreams are complete. Advancing workflow.'
    );

    const actionForAdvance =
      completedAction ||
      getChecklistActionByKey_(tx, tx['Current Action Key']);

    return advanceTransaction_(
      user,
      tx,
      actionForAdvance,
      checklist
        ? 'Parallel listing preparation completed: ' +
          checklist.progress.totalRequired +
          ' required checklist items'
        : 'Parallel listing preparation reconciled and completed'
    );
  }

  const nextActionKey = chooseNextParallelAction_(
    tx,
    remainingActionKeys
  );

  const nextAction = setTransactionCurrentAction_(
    transactionId,
    tx,
    nextActionKey
  );

  logActivity_(
    user,
    transactionId,
    'PARALLEL_ACTION_ROUTED',
    tx['Current Stage Name'],
    tx['Current Stage Name'],
    completedAction ? completedAction['Action Name'] : '',
    nextAction['Action Name'],
    'Parallel action completed. Routed to the remaining listing-preparation task.'
  );

  return {
    success: true,
    transactionId: transactionId,
    stageName: tx['Current Stage Name'],
    actionKey: nextActionKey,
    actionName: nextAction['Action Name'],
    parallel: true,
    groupComplete: false,
    preparation: getParallelListingPreparation(transactionId)
  };
}

function getParallelWorkflowGroup_(tx) {
  if (tx['Workflow Key'] !== 'SELLER_LISTING') return null;

  return JBA_WORKFLOW_ENGINE.parallelGroups.LISTING_PREPARATION;
}

function chooseNextParallelAction_(tx, remainingActionKeys) {
  const currentActionKey = tx['Current Action Key'];

  const otherAction = remainingActionKeys.find(
    actionKey => actionKey !== currentActionKey
  );

  return otherAction || remainingActionKeys[0];
}

/**
 * Updates the transaction's current action while keeping it in the same stage.
 */
function setTransactionCurrentAction_(transactionId, tx, actionKey) {
  const action = getChecklistActionByKey_(tx, actionKey);
  const sheet = getDatabase_().getSheetByName(JBA_OS.sheets.transactions);

  if (!sheet) {
    throw new Error('Transactions sheet was not found.');
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error('No transactions were found.');
  }

  const headers = values[0];
  const idColumn = headers.indexOf('Transaction ID');

  if (idColumn < 0) {
    throw new Error(
      'Transactions sheet is missing the Transaction ID column.'
    );
  }

  const rowIndex = values.slice(1).findIndex(
    row => String(row[idColumn]) === String(transactionId)
  );

  if (rowIndex < 0) {
    throw new Error('Transaction not found: ' + transactionId);
  }

  const updates = {
    'Current Action Key': actionKey,
    'Current Action Name': action['Action Name'] || actionKey,
    'Current Assigned Role': action['Assigned Role'] || '',
    'Updated At': new Date(),
    'Updated By': Session.getActiveUser().getEmail() || ''
  };

  Object.keys(updates).forEach(header => {
    const columnIndex = headers.indexOf(header);

    // Some installations may not yet have every optional audit column.
    if (columnIndex >= 0) {
      sheet.getRange(rowIndex + 2, columnIndex + 1)
        .setValue(updates[header]);
    }
  });

  return action;
}
