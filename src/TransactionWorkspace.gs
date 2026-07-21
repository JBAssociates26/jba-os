/**
 * JBA OS — Transaction Workspace
 *
 * Provides the complete workspace data for one transaction:
 * - Property and seller information
 * - Assigned team members
 * - Current workflow stage and task
 * - Workflow progress timeline
 * - Activity history
 * - Role-based permissions
 */


/**
 * Returns all authorized workspace data for a transaction.
 *
 * @param {string} transactionId
 * @return {Object}
 */
function getTransactionWorkspace(transactionId) {
  const auth = getAuthorizedTransaction_(transactionId);
  const user = auth.user;
  const tx = auth.tx;

  const currentAction = getCurrentActionForTransaction_(tx);
  const timeline = buildTransactionTimeline_(tx);
  const activity = getTransactionActivity_(transactionId);

  logActivity_(
    user,
    transactionId,
    'TRANSACTION_WORKSPACE_OPENED',
    tx['Current Stage Name'],
    tx['Current Stage Name'],
    tx['Current Action Name'],
    currentAction ? currentAction['Action Name'] : '',
    'Transaction workspace opened'
  );

  return {
    success: true,

    transactionId: tx['Transaction ID'],
    status: tx['Status'],

    property: {
      fullAddress: buildWorkspaceAddress_(tx),
      addressLine1:
        tx['Address Line 1'] ||
        tx['Property Address'] ||
        '',
      addressLine2: tx['Address Line 2'] || '',
      city: tx['City'] || '',
      state: tx['State'] || '',
      postalCode: tx['Postal Code'] || '',
      county: tx['County'] || '',
      propertyType: tx['Property Type'] || ''
    },

    seller: {
      fullName: [
        tx['Client First Name'],
        tx['Client Last Name']
      ].filter(Boolean).join(' '),

      firstName: tx['Client First Name'] || '',
      lastName: tx['Client Last Name'] || '',
      email: tx['Client Email'] || '',
      phone: tx['Client Phone'] || ''
    },

    appointment: {
      date: tx['Appointment Date'] || '',
      time: tx['Appointment Time'] || '',
      display: buildWorkspaceAppointment_(tx)
    },

    financials: {
      canView: Boolean(user.canViewFinancials),

      anticipatedListPrice:
        user.canViewFinancials
          ? formatWorkspaceCurrency_(
              tx['Anticipated List Price']
            )
          : '',

      actualListPrice:
        user.canViewFinancials
          ? formatWorkspaceCurrency_(
              tx['Actual List Price']
            )
          : '',

      contractPrice:
        user.canViewFinancials
          ? formatWorkspaceCurrency_(
              tx['Contract Price']
            )
          : '',

      closingDate:
        user.canViewFinancials
          ? tx['Closing Date'] || ''
          : ''
    },

    team: {
      listingAgent: tx['Listing Agent'] || '',
      agentEmail: tx['Agent Email'] || '',
      operationsEmail:
        tx['Assigned Operations Email'] || '',
      marketingEmail:
        tx['Assigned Marketing Email'] || '',
      transactionCoordinatorEmail:
        tx['Assigned TC Email'] || ''
    },

    workflow: {
      workflowKey: tx['Workflow Key'] || '',
      currentStageKey:
        tx['Current Stage Key'] || '',
      currentStageName:
        tx['Current Stage Name'] || '',

      currentActionKey:
        currentAction
          ? currentAction['Action Key']
          : '',

      currentActionName:
        currentAction
          ? currentAction['Action Name']
          : 'No action configured',

      actionType:
        currentAction
          ? currentAction['Action Type']
          : '',

      assignedRole:
        currentAction
          ? currentAction['Assigned Role']
          : '',

      formKey:
        currentAction
          ? currentAction['Form Key']
          : '',

      progressPercent:
        calculateTransactionProgress_(timeline),

      timeline: timeline
    },

    listingDetails: {
      urgency: tx['Urgency'] || 'Normal',
      binderNeeded:
        tx['Binder Needed?'] || '',
      binderStatus:
        tx['Binder Status'] || '',
      needsReview:
        tx['Needs Review?'] === 'Yes',
      reviewReasons:
        user.canViewFinancials
          ? tx['Review Reasons'] || ''
          : '',
      notes: tx['Notes'] || '',
      createdAt: tx['Created At'] || '',
      updatedAt: tx['Updated At'] || '',
      lastActionBy:
        tx['Last Action By'] || '',
      archiveReason:
        tx['Archive Reason'] || ''
    },

    permissions: {
      canChangeStage:
        Boolean(user.canChangeStage),

      canReassign:
        Boolean(user.canReassign),

      canViewFinancials:
        Boolean(user.canViewFinancials),

      canOpenForm:
        Boolean(currentAction) &&
        currentAction['Action Type'] === 'FORM' &&
        canUserPerformAction_(
          user,
          currentAction
        ),

      canCompleteInternalAction:
        Boolean(currentAction) &&
        currentAction['Action Type'] === 'INTERNAL' &&
        canUserPerformAction_(
          user,
          currentAction
        ),

      canEditNotes:
        canEditTransactionNotes_(user, tx)
    },

    activity: activity
  };
}


/**
 * Builds the workflow timeline for a transaction.
 *
 * @param {Object} tx
 * @return {Array}
 */
function buildTransactionTimeline_(tx) {
  const stages = sheetObjects_(
    JBA_OS.sheets.workflowStages
  )
    .filter(function(stage) {
      return (
        stage['Workflow Key'] ===
          tx['Workflow Key'] &&
        stage['Active?'] === 'Yes'
      );
    })
    .sort(function(a, b) {
      return (
        Number(a['Stage Order'] || 0) -
        Number(b['Stage Order'] || 0)
      );
    });

  const currentIndex = stages.findIndex(
    function(stage) {
      return (
        stage['Stage Key'] ===
        tx['Current Stage Key']
      );
    }
  );

  return stages.map(function(stage, index) {
    let timelineStatus = 'upcoming';

    if (index < currentIndex) {
      timelineStatus = 'complete';
    }

    if (index === currentIndex) {
      timelineStatus = 'current';
    }

    return {
      stageKey: stage['Stage Key'],
      stageName: stage['Stage Name'],
      stageOrder:
        Number(stage['Stage Order'] || 0),
      description:
        stage['Description'] || '',
      status: timelineStatus
    };
  });
}


/**
 * Calculates progress from the workflow timeline.
 *
 * @param {Array} timeline
 * @return {number}
 */
function calculateTransactionProgress_(timeline) {
  if (!timeline || !timeline.length) {
    return 0;
  }

  const currentIndex = timeline.findIndex(
    function(stage) {
      return stage.status === 'current';
    }
  );

  if (currentIndex < 0) {
    return 0;
  }

  if (timeline.length === 1) {
    return 100;
  }

  return Math.round(
    (currentIndex / (timeline.length - 1)) * 100
  );
}


/**
 * Returns recent activity for one transaction.
 *
 * @param {string} transactionId
 * @return {Array}
 */
function getTransactionActivity_(transactionId) {
  return sheetObjects_(
    JBA_OS.sheets.activity
  )
    .filter(function(row) {
      return (
        String(
          row['Transaction ID'] || ''
        ).trim() ===
        String(transactionId || '').trim()
      );
    })
    .reverse()
    .slice(0, 50)
    .map(function(row) {
      return {
        timestamp: row['Timestamp'] || '',
        userEmail: row['User Email'] || '',
        userRole: row['User Role'] || '',
        action: row['Action'] || '',
        previousStage:
          row['Previous Stage'] || '',
        newStage: row['New Stage'] || '',
        previousAction:
          row['Previous Action'] || '',
        newAction:
          row['New Action'] || '',
        details: row['Details'] || ''
      };
    });
}


/**
 * Updates transaction notes from the workspace.
 *
 * @param {string} transactionId
 * @param {string} notes
 * @return {Object}
 */
function updateTransactionNotes(
  transactionId,
  notes
) {
  const auth =
    getAuthorizedTransaction_(transactionId);

  if (
    !canEditTransactionNotes_(
      auth.user,
      auth.tx
    )
  ) {
    throw new Error(
      'You do not have permission to edit these notes.'
    );
  }

  const cleanNotes =
    String(notes || '').trim();

  updateTransactionFields_(
    transactionId,
    {
      'Notes': cleanNotes,
      'Updated At': new Date(),
      'Last Action By': auth.user.email
    }
  );

  logActivity_(
    auth.user,
    transactionId,
    'TRANSACTION_NOTES_UPDATED',
    auth.tx['Current Stage Name'],
    auth.tx['Current Stage Name'],
    auth.tx['Current Action Name'],
    auth.tx['Current Action Name'],
    'Transaction notes updated'
  );

  return {
    success: true,
    notes: cleanNotes
  };
}


/**
 * Determines whether the user can edit notes.
 *
 * @param {Object} user
 * @param {Object} tx
 * @return {boolean}
 */
function canEditTransactionNotes_(user, tx) {
  if (
    [
      'Executive Admin',
      'Operations Admin'
    ].includes(user.role)
  ) {
    return true;
  }

  if (user.role === 'Agent') {
    return (
      normalizeEmail_(
        tx['Agent Email']
      ) === user.agentEmailMatch
    );
  }

  if (user.role === 'Marketing') {
    return (
      normalizeEmail_(
        tx['Assigned Marketing Email']
      ) === user.email
    );
  }

  if (
    user.role ===
    'Transaction Coordinator'
  ) {
    return (
      normalizeEmail_(
        tx['Assigned TC Email']
      ) === user.email
    );
  }

  return false;
}


/**
 * Builds the full property address.
 *
 * @param {Object} tx
 * @return {string}
 */
function buildWorkspaceAddress_(tx) {
  const address = [
    tx['Address Line 1'] ||
      tx['Property Address'],
    tx['Address Line 2']
  ].filter(Boolean).join(', ');

  const cityStateZip = [
    tx['City'],
    tx['State'],
    tx['Postal Code']
  ].filter(Boolean).join(' ');

  return [
    address,
    cityStateZip
  ].filter(Boolean).join(', ');
}


/**
 * Builds the appointment display.
 *
 * @param {Object} tx
 * @return {string}
 */
function buildWorkspaceAppointment_(tx) {
  const date =
    tx['Appointment Date'] || '';

  const time =
    tx['Appointment Time'] || '';

  if (!date && !time) {
    return 'Not scheduled';
  }

  return [date, time]
    .filter(Boolean)
    .join(' at ');
}


/**
 * Formats a workspace currency value.
 *
 * @param {*} value
 * @return {string}
 */
function formatWorkspaceCurrency_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return 'Not provided';
  }

  const cleanValue = String(value)
    .replace(/[$,\s]/g, '');

  const numberValue =
    Number(cleanValue);

  if (!isFinite(numberValue)) {
    return String(value);
  }

  return '$' +
    numberValue.toLocaleString(
      'en-US',
      {
        maximumFractionDigits: 0
      }
    );
}


/**
 * Manual test.
 *
 * Replace the sample ID with a real Transaction ID.
 */
function testTransactionWorkspace() {
  const transactionId =
    'REPLACE-WITH-TRANSACTION-ID';

  const result =
    getTransactionWorkspace(
      transactionId
    );

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );
}
