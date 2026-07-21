/**
 * JBA OS — NEW LISTING WIZARD
 *
 * Add this as a new Apps Script file named:
 * NewListingWizard.gs
 *
 * This module uses the existing JBA OS functions:
 * - requireUser_()
 * - sheetObjects_()
 * - getSetting_()
 * - createSellerListingTransaction()
 */

function getNewListingWizardOptions() {
  const user = requireUser_();

  if (!['Executive Admin', 'Operations Admin', 'Agent'].includes(user.role)) {
    throw new Error('You are not authorized to create listings.');
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

  // Admin users may create a listing for themselves during testing.
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

function createNewListingFromWizard(payload) {
  const user = requireUser_();

  if (!['Executive Admin', 'Operations Admin', 'Agent'].includes(user.role)) {
    throw new Error('You are not authorized to create listings.');
  }

  payload = payload || {};

  const required = {
    propertyAddress: 'Property address',
    city: 'City',
    postalCode: 'ZIP code',
    clientFirstName: 'Seller first name',
    clientLastName: 'Seller last name',
    agentEmail: 'Listing agent'
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
    payload.clientEmail &&
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(payload.clientEmail).trim())
  ) {
    throw new Error('Enter a valid seller email address.');
  }

  // Agents may only create listings assigned to themselves.
  if (
    user.role === 'Agent' &&
    String(payload.agentEmail).trim().toLowerCase() !== user.agentEmailMatch
  ) {
    throw new Error('Agents may only create listings assigned to themselves.');
  }

  const clean = {
    listingAgent: String(payload.listingAgent || '').trim(),
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

    clientFirstName: String(payload.clientFirstName || '').trim(),
    clientLastName: String(payload.clientLastName || '').trim(),
    clientEmail: String(payload.clientEmail || '').trim().toLowerCase(),
    clientPhone: String(payload.clientPhone || '').trim(),

    propertyAddress: String(payload.propertyAddress || '').trim(),
    addressLine1: String(payload.propertyAddress || '').trim(),
    addressLine2: String(payload.addressLine2 || '').trim(),
    city: String(payload.city || '').trim(),
    state: String(payload.state || 'MI').trim().toUpperCase(),
    postalCode: String(payload.postalCode || '').trim(),
    county: String(payload.county || '').trim(),
    propertyType: String(payload.propertyType || '').trim(),

    appointmentDate: String(payload.appointmentDate || '').trim(),
    appointmentTime: String(payload.appointmentTime || '').trim(),
    anticipatedListPrice: normalizeCurrencyInput_(payload.anticipatedListPrice),

    urgency: String(payload.urgency || '').trim(),
    binderNeeded: String(payload.binderNeeded || '').trim(),
    binderStatus:
      String(payload.binderNeeded || '').trim() === 'Yes'
        ? 'Requested'
        : 'Not Needed',

    notes: String(payload.notes || '').trim(),
    sourceSystem: 'JBA OS New Listing Wizard'
  };

const result = createSellerListingTransaction(clean);
var notificationResult = sendNewListingNotifications_(
  result.transactionId,
  clean
);

  return {
    success: true,
    transactionId: result.transactionId,
    message:
      'Listing created successfully. Transaction ID: ' +
      result.transactionId
  };
}

function normalizeCurrencyInput_(value) {
  const text = String(value || '').replace(/[$,\s]/g, '');
  if (!text) return '';

  const amount = Number(text);
  if (!Number.isFinite(amount) || amount < 0) {
    throw new Error('Expected list price must be a valid positive number.');
  }

  return amount;
}
