/**
 * JBA OS — New Listing Email Notifications
 *
 * Sends:
 * 1. Confirmation and next-step email to the listing agent.
 * 2. New-listing marketing alert to the assigned marketing coordinator.
 */


/**
 * Main notification function.
 *
 * @param {string} transactionId
 * @param {Object} listing
 * @return {Object}
 */
function sendNewListingNotifications_(transactionId, listing) {
  var results = {
    agentSent: false,
    marketingSent: false,
    errors: []
  };

  var portalUrl = ScriptApp.getService().getUrl() || '';
  var propertyAddress = buildNewListingAddress_(listing);
  var sellerName = buildNewListingSellerName_(listing);
  var appointment = formatNewListingAppointment_(
    listing.appointmentDate,
    listing.appointmentTime
  );

  if (listing.agentEmail) {
    try {
      sendNewListingAgentEmail_({
        transactionId: transactionId,
        listing: listing,
        propertyAddress: propertyAddress,
        sellerName: sellerName,
        appointment: appointment,
        portalUrl: portalUrl
      });

      results.agentSent = true;
    } catch (error) {
      results.errors.push(
        'Agent email failed: ' + error.message
      );
    }
  }

  if (listing.assignedMarketingEmail) {
    try {
      sendNewListingMarketingEmail_({
        transactionId: transactionId,
        listing: listing,
        propertyAddress: propertyAddress,
        sellerName: sellerName,
        appointment: appointment,
        portalUrl: portalUrl
      });

      results.marketingSent = true;
    } catch (error) {
      results.errors.push(
        'Marketing email failed: ' + error.message
      );
    }
  }

  return results;
}


/**
 * Sends the listing agent confirmation email.
 *
 * @param {Object} data
 */
function sendNewListingAgentEmail_(data) {
  var listing = data.listing;

  var subject =
    'Your New Listing Is Ready in JBA OS — ' +
    data.propertyAddress;

  var plainBody = [
    'A new listing has been created in JBA OS.',
    '',
    'Property: ' + data.propertyAddress,
    'Seller: ' + (data.sellerName || 'Not provided'),
    'Transaction ID: ' + data.transactionId,
    'Appointment: ' + data.appointment,
    'Expected List Price: ' +
      formatNewListingCurrency_(
        listing.anticipatedListPrice
      ),
    '',
    'Current Stage: Pre-Listing',
    'Your Next Step: Complete the Pre-Listing Checklist',
    '',
    data.portalUrl
  ].join('\n');

  var htmlBody = buildNewListingEmailShell_(
    'Your New Listing Is Ready',
    [
      '<p>Hi ' +
        escapeNewListingHtml_(
          listing.listingAgent || 'there'
        ) +
        ',</p>',

      '<p>Your new listing has been created in JBA OS.</p>',

      buildNewListingDetailsTable_([
        ['Property', data.propertyAddress],
        ['Seller', data.sellerName || 'Not provided'],
        ['Transaction ID', data.transactionId],
        ['Appointment', data.appointment],
        [
          'Expected List Price',
          formatNewListingCurrency_(
            listing.anticipatedListPrice
          )
        ],
        ['Current Stage', 'Pre-Listing'],
        [
          'Current Action',
          'Complete the Pre-Listing Checklist'
        ]
      ]),

      buildNewListingActionBox_(
        'Your Next Step',
        'Open JBA OS and complete the Pre-Listing Checklist.'
      ),

      buildNewListingPortalButton_(
        data.portalUrl,
        'Open JBA OS'
      )
    ].join('')
  );

  MailApp.sendEmail({
    to: listing.agentEmail,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: 'Jim Bunn & Associates'
  });
}


/**
 * Sends the assigned marketing coordinator an alert.
 *
 * @param {Object} data
 */
function sendNewListingMarketingEmail_(data) {
  var listing = data.listing;

  var subject =
    'New JBA Listing Created — ' +
    data.propertyAddress;

  var plainBody = [
    'A new listing has been created in JBA OS.',
    '',
    'Property: ' + data.propertyAddress,
    'Seller: ' + (data.sellerName || 'Not provided'),
    'Listing Agent: ' +
      (
        listing.listingAgent ||
        listing.agentEmail ||
        'Not provided'
      ),
    'Agent Email: ' +
      (listing.agentEmail || 'Not provided'),
    'Transaction ID: ' + data.transactionId,
    'Appointment: ' + data.appointment,
    'Expected List Price: ' +
      formatNewListingCurrency_(
        listing.anticipatedListPrice
      ),
    'Property Type: ' +
      (listing.propertyType || 'Not provided'),
    'County: ' +
      (listing.county || 'Not provided'),
    'Listing Binder Needed: ' +
      (listing.binderNeeded || 'Not specified'),
    'Urgency: ' +
      (listing.urgency || 'Normal'),
    '',
    'Marketing Next Steps:',
    '1. Review the listing details.',
    '2. Prepare the pre-listing folder and materials.',
    '3. Confirm whether a listing binder is needed.',
    '4. Monitor JBA OS for the completed checklist.',
    '5. Wait for photography and launch instructions.',
    '',
    'Notes: ' +
      (listing.notes || 'No notes provided'),
    '',
    data.portalUrl
  ].join('\n');

  var marketingContent = [
    '<p>Hi Marketing Team,</p>',

    '<p>A new listing has entered the JBA pre-listing workflow.</p>',

    buildNewListingDetailsTable_([
      ['Property', data.propertyAddress],
      ['Seller', data.sellerName || 'Not provided'],
      [
        'Listing Agent',
        listing.listingAgent ||
          listing.agentEmail ||
          'Not provided'
      ],
      ['Agent Email', listing.agentEmail || 'Not provided'],
      ['Transaction ID', data.transactionId],
      ['Appointment', data.appointment],
      [
        'Expected List Price',
        formatNewListingCurrency_(
          listing.anticipatedListPrice
        )
      ],
      ['Property Type', listing.propertyType || 'Not provided'],
      ['County', listing.county || 'Not provided'],
      [
        'Listing Binder Needed',
        listing.binderNeeded || 'Not specified'
      ],
      ['Urgency', listing.urgency || 'Normal']
    ]),

    buildNewListingMarketingSteps_()
  ];

  if (listing.notes) {
    marketingContent.push(
      '<h3 style="margin-top:24px;">Agent Notes</h3>' +
      '<div style="' +
        'background:#f7f7f7;' +
        'padding:14px;' +
        'border-radius:8px;' +
        'white-space:pre-wrap;' +
      '">' +
        escapeNewListingHtml_(listing.notes) +
      '</div>'
    );
  }

  marketingContent.push(
    buildNewListingPortalButton_(
      data.portalUrl,
      'Open JBA OS'
    )
  );

  var htmlBody = buildNewListingEmailShell_(
    'New Listing Created',
    marketingContent.join('')
  );

  MailApp.sendEmail({
    to: listing.assignedMarketingEmail,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: 'Jim Bunn & Associates'
  });
}


/**
 * Shared branded email layout.
 *
 * @param {string} title
 * @param {string} content
 * @return {string}
 */
function buildNewListingEmailShell_(title, content) {
  return [
    '<!DOCTYPE html>',
    '<html>',
    '<body style="',
      'margin:0;',
      'padding:0;',
      'background:#f3f4f6;',
      'font-family:Arial,Helvetica,sans-serif;',
      'color:#20242a;',
    '">',

      '<div style="',
        'max-width:680px;',
        'margin:0 auto;',
        'padding:24px 12px;',
      '">',

        '<div style="',
          'background:#17191d;',
          'color:#ffffff;',
          'padding:24px;',
          'border-radius:12px 12px 0 0;',
        '">',

          '<div style="font-size:20px;font-weight:bold;">',
            'Jim Bunn &amp; Associates',
          '</div>',

          '<div style="',
            'color:#c7c9cc;',
            'font-size:13px;',
            'margin-top:5px;',
          '">',
            'JBA OS',
          '</div>',

        '</div>',

        '<div style="',
          'background:#ffffff;',
          'padding:28px;',
          'border-radius:0 0 12px 12px;',
        '">',

          '<h2 style="margin:0 0 20px;font-size:24px;">',
            escapeNewListingHtml_(title),
          '</h2>',

          content,

          '<div style="',
            'margin-top:30px;',
            'padding-top:18px;',
            'border-top:1px solid #e2e5e9;',
            'color:#777777;',
            'font-size:12px;',
          '">',
            'This notification was generated automatically by JBA OS.',
          '</div>',

        '</div>',
      '</div>',
    '</body>',
    '</html>'
  ].join('');
}


/**
 * Creates a details table for the email.
 *
 * @param {Array} rows
 * @return {string}
 */
function buildNewListingDetailsTable_(rows) {
  var htmlRows = rows.map(function(row) {
    return [
      '<tr>',

        '<td style="',
          'padding:10px 12px;',
          'border-bottom:1px solid #eceef1;',
          'color:#6c737d;',
          'font-size:13px;',
          'width:38%;',
          'vertical-align:top;',
        '">',
          escapeNewListingHtml_(row[0]),
        '</td>',

        '<td style="',
          'padding:10px 12px;',
          'border-bottom:1px solid #eceef1;',
          'font-size:14px;',
          'font-weight:600;',
          'vertical-align:top;',
        '">',
          escapeNewListingHtml_(
            row[1] || 'Not provided'
          ),
        '</td>',

      '</tr>'
    ].join('');
  }).join('');

  return [
    '<table role="presentation" ',
      'width="100%" ',
      'cellpadding="0" ',
      'cellspacing="0" ',
      'style="',
        'width:100%;',
        'border:1px solid #e2e5e9;',
        'border-collapse:collapse;',
        'margin:20px 0;',
      '">',
      htmlRows,
    '</table>'
  ].join('');
}


/**
 * Creates the agent next-step box.
 */
function buildNewListingActionBox_(title, message) {
  return [
    '<div style="',
      'background:#f5f1e8;',
      'border-left:4px solid #b79a5b;',
      'padding:16px;',
      'margin:22px 0;',
    '">',

      '<strong>',
        escapeNewListingHtml_(title),
      '</strong>',

      '<p style="margin:8px 0 0;">',
        escapeNewListingHtml_(message),
      '</p>',

    '</div>'
  ].join('');
}


/**
 * Creates the marketing steps section.
 */
function buildNewListingMarketingSteps_() {
  return [
    '<div style="',
      'background:#f5f1e8;',
      'border-left:4px solid #b79a5b;',
      'padding:16px;',
      'margin:22px 0;',
    '">',

      '<strong>Marketing Next Steps</strong>',

      '<ol style="margin-bottom:0;padding-left:22px;">',
        '<li>Review the listing details.</li>',
        '<li>Prepare the pre-listing folder and materials.</li>',
        '<li>Confirm whether a listing binder is needed.</li>',
        '<li>Monitor JBA OS for the completed checklist.</li>',
        '<li>Wait for photography and launch instructions.</li>',
      '</ol>',

    '</div>'
  ].join('');
}


/**
 * Creates the JBA OS button.
 */
function buildNewListingPortalButton_(url, label) {
  if (!url) {
    return '';
  }

  return [
    '<div style="margin-top:24px;">',
      '<a href="',
        escapeNewListingHtml_(url),
      '" style="',
        'display:inline-block;',
        'background:#b79a5b;',
        'color:#20242a;',
        'text-decoration:none;',
        'padding:12px 18px;',
        'border-radius:8px;',
        'font-weight:bold;',
      '">',
        escapeNewListingHtml_(label),
      '</a>',
    '</div>'
  ].join('');
}


/**
 * Builds the complete property address.
 */
function buildNewListingAddress_(listing) {
  var firstLine = [
    listing.propertyAddress,
    listing.addressLine2
  ].filter(Boolean).join(', ');

  var cityStateZip = [
    listing.city,
    listing.state,
    listing.postalCode
  ].filter(Boolean).join(' ');

  return [
    firstLine,
    cityStateZip
  ].filter(Boolean).join(', ');
}


/**
 * Builds the primary seller's name.
 */
function buildNewListingSellerName_(listing) {
  return [
    listing.clientFirstName,
    listing.clientLastName
  ].filter(Boolean).join(' ');
}


/**
 * Formats appointment date and time.
 */
function formatNewListingAppointment_(dateValue, timeValue) {
  if (!dateValue && !timeValue) {
    return 'Not scheduled';
  }

  var formattedDate = dateValue || '';
  var formattedTime = timeValue || '';

  if (dateValue) {
    var dateParts = String(dateValue).split('-');

    if (dateParts.length === 3) {
      formattedDate =
        dateParts[1] +
        '/' +
        dateParts[2] +
        '/' +
        dateParts[0];
    }
  }

  if (timeValue) {
    var timeParts = String(timeValue).split(':');
    var hour = Number(timeParts[0]);
    var minute = timeParts[1] || '00';
    var suffix = hour >= 12 ? 'PM' : 'AM';

    hour = hour % 12 || 12;

    formattedTime =
      hour +
      ':' +
      minute +
      ' ' +
      suffix;
  }

  return [
    formattedDate,
    formattedTime
  ].filter(Boolean).join(' at ');
}


/**
 * Formats list price.
 */
function formatNewListingCurrency_(value) {
  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return 'Not provided';
  }

  var number = Number(value);

  if (!isFinite(number)) {
    return String(value);
  }

  return '$' + number.toLocaleString('en-US', {
    maximumFractionDigits: 0
  });
}


/**
 * Escapes values before inserting into HTML.
 */
function escapeNewListingHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


/**
 * Run this once manually to authorize MailApp permissions.
 */
function authorizeJBAEmailPermissions() {
  var quota = MailApp.getRemainingDailyQuota();

  Logger.log(
    'Email permission authorized. Remaining quota: ' +
    quota
  );
}
