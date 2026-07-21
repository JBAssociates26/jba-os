/**
 * JBA OS — NEW BUYER TRANSACTION
 *
 * Add this as a new Apps Script file named:
 * NewBuyerWizard.gs
 *
 * The CRM owns the buyer relationship up through an accepted offer;
 * this is the intake point where JBA OS takes over, starting at Pending.
 *
 * This module uses the existing JBA OS functions:
 * - requireUser_()
 * - sheetObjects_()
 * - getSetting_()
 * - normalizeCurrencyInput_() (NewListingWizard.gs)
 * - createBuyerTransaction() (Code.gs)
 */

function getNewBuyerWizardOptions() {
  const user = requireUser_();

  if (!['Executive Admin', 'Operations Admin', 'Agent'].includes(user.role)) {
    throw new Error('You are not authorized to create transactions.');
  }

  const users = sheetObjects_(JBA_OS.sheets.users)
    .filter(row => row['Active?'] === 'Yes');

  const agents = users
    .filter(row => row['Role'] === 'Agent')
    .map(row => ({
      name: row['Display Name'] || row['Email'],
      email: String(row['Agent Email Match'] || row['Email'] || '')
        .trim()
        .toLowerCase()
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  // Admin users may create a transaction for themselves during testing.
  if (
    !agents.some(agent => agent.email === user.email) &&
    ['Executive Admin', 'Operations Admin'].includes(user.role)
  ) {
    agents.unshift({
      name: user.displayName || user.email,
      email: user.email
    });
  }

  const marketingUsers = users
    .filter(row => row['Role'] === 'Marketing')
    .map(row => ({
      name: row['Display Name'] || row['Email'],
      email: String(row['Email'] || '').trim().toLowerCase()
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const tcUsers = users
    .filter(row => row['Role'] === 'Transaction Coordinator')
    .map(row => ({
      name: row['Display Name'] || row['Email'],
      email: String(row['Email'] || '').trim().toLowerCase()
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  return {
    currentUser: {
      email: user.email,
      displayName: user.displayName || user.email,
      role: user.role
    },
    agents,
    marketingUsers,
    tcUsers,
    defaults: {
      state: 'MI',
      operationsEmail:
        String(getSetting_('Default Operations Email') || '')
          .trim()
          .toLowerCase(),
      marketingEmail:
        String(getSetting_('Default Marketing Email') || '')
          .trim()
          .toLowerCase()
    },
    propertyTypes: [
      'Single Family',
      'Condominium',
      'Townhouse',
      'Multi-Family',
      'Vacant Land',
      'Commercial',
      'Farm',
      'Manufactured Home',
      'Other'
    ],
    counties: [
      'Livingston',
      'Oakland',
      'Genesee',
      'Washtenaw',
      'Ingham',
      'Shiawassee',
      'Jackson',
      'Wayne',
      'Other'
    ]
  };
}

function createNewBuyerTransaction(payload) {
  const user = requireUser_();

  if (!['Executive Admin', 'Operations Admin', 'Agent'].includes(user.role)) {
    throw new Error('You are not authorized to create transactions.');
  }

  payload = payload || {};

  const required = {
    buyerFirstName: 'Buyer first name',
    buyerLastName: 'Buyer last name',
    propertyAddress: 'Property address',
    city: 'City',
    postalCode: 'ZIP code',
    agentEmail: "Buyer's agent",
    contractPrice: 'Contract price'
  };

  Object.keys(required).forEach(key => {
    if (!String(payload[key] || '').trim()) {
      throw new Error(required[key] + ' is required.');
    }
  });

  if (!/^\d{5}(-\d{4})?$/.test(String(payload.postalCode).trim())) {
    throw new Error('Enter a valid ZIP code.');
  }

  if (
    payload.buyerEmail &&
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(payload.buyerEmail).trim())
  ) {
    throw new Error('Enter a valid buyer email address.');
  }

  // Agents may only create transactions assigned to themselves.
  if (
    user.role === 'Agent' &&
    String(payload.agentEmail).trim().toLowerCase() !== user.agentEmailMatch
  ) {
    throw new Error('Agents may only create transactions assigned to themselves.');
  }

  const clean = {
    agentName: String(payload.agentName || '').trim(),
    agentEmail: String(payload.agentEmail || '').trim().toLowerCase(),

    assignedOperationsEmail:
      String(
        payload.assignedOperationsEmail ||
        getSetting_('Default Operations Email') ||
        ''
      ).trim().toLowerCase(),

    assignedMarketingEmail:
      String(
        payload.assignedMarketingEmail ||
        getSetting_('Default Marketing Email') ||
        ''
      ).trim().toLowerCase(),

    assignedTCEmail:
      String(payload.assignedTCEmail || '').trim().toLowerCase(),

    buyerFirstName: String(payload.buyerFirstName || '').trim(),
    buyerLastName: String(payload.buyerLastName || '').trim(),
    buyerEmail: String(payload.buyerEmail || '').trim().toLowerCase(),
    buyerPhone: String(payload.buyerPhone || '').trim(),

    propertyAddress: String(payload.propertyAddress || '').trim(),
    addressLine1: String(payload.propertyAddress || '').trim(),
    addressLine2: String(payload.addressLine2 || '').trim(),
    city: String(payload.city || '').trim(),
    state: String(payload.state || 'MI').trim().toUpperCase(),
    postalCode: String(payload.postalCode || '').trim(),
    county: String(payload.county || '').trim(),
    propertyType: String(payload.propertyType || '').trim(),

    contractPrice: normalizeCurrencyInput_(payload.contractPrice),
    closingDate: String(payload.closingDate || '').trim(),

    notes: String(payload.notes || '').trim(),
    sourceSystem: 'JBA OS New Buyer Transaction'
  };

  const result = createBuyerTransaction(clean);

  sendNewBuyerTransactionNotifications_(result.transactionId, clean);

  return {
    success: true,
    transactionId: result.transactionId,
    message:
      'Buyer transaction created successfully. Transaction ID: ' +
      result.transactionId
  };
}

/**
 * Notifies the agent (and TC, if assigned) that a new buyer transaction
 * was created. Reuses the shared branded-email helpers from
 * NewListingNotifications.gs (buildNewListingEmailShell_ etc. are
 * generic despite the file name, not seller-specific) instead of
 * duplicating that whole system for a simpler single-email case.
 */
function sendNewBuyerTransactionNotifications_(transactionId, buyer) {
  if (!buyer.agentEmail) return { sent: false };

  const portalUrl = ScriptApp.getService().getUrl() || '';
  const address = buildNewListingAddress_({
    propertyAddress: buyer.propertyAddress,
    addressLine2: buyer.addressLine2,
    city: buyer.city,
    state: buyer.state,
    postalCode: buyer.postalCode
  });
  const buyerName = [buyer.buyerFirstName, buyer.buyerLastName].filter(Boolean).join(' ');

  const htmlBody = buildNewListingEmailShell_(
    'New Buyer Transaction Created',
    [
      '<p>A new buyer transaction has entered the JBA OS Pending workflow.</p>',

      buildNewListingDetailsTable_([
        ['Property', address],
        ['Buyer', buyerName || 'Not provided'],
        ['Transaction ID', transactionId],
        ['Contract Price', formatNewListingCurrency_(buyer.contractPrice)],
        ['Closing Date', buyer.closingDate || 'Not provided'],
        ['Current Stage', 'Pending'],
        ['Current Action', 'Complete the Pending Checklist']
      ]),

      buildNewListingActionBox_(
        'Next Step',
        'Open JBA OS and complete the Pending Checklist.'
      ),

      buildNewListingPortalButton_(portalUrl, 'Open JBA OS')
    ].join('')
  );

  const plainBody = [
    'A new buyer transaction has been created in JBA OS.',
    '',
    'Property: ' + address,
    'Buyer: ' + (buyerName || 'Not provided'),
    'Transaction ID: ' + transactionId,
    'Contract Price: ' + formatNewListingCurrency_(buyer.contractPrice),
    '',
    'Current Stage: Pending',
    'Next Step: Complete the Pending Checklist',
    '',
    portalUrl
  ].join('\n');

  MailApp.sendEmail({
    to: buyer.agentEmail,
    cc: buyer.assignedTCEmail || '',
    subject: 'New Buyer Transaction Created — ' + address,
    body: plainBody,
    htmlBody: htmlBody,
    name: 'Jim Bunn & Associates'
  });

  return { sent: true };
}
