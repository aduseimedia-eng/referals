// Google Apps Script - Deploy this in your Google Sheet
// Instructions:
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire script
// 4. Click "Deploy" > "New deployment"
// 5. Select type: "Web app"
// 6. Execute as: "Me"
// 7. Who has access: "Anyone"
// 8. Click "Deploy" and copy the Web App URL
// 9. Paste that URL in your index.html where it says "YOUR_SCRIPT_URL_HERE"

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Parse the incoming data
    const data = JSON.parse(e.postData.contents);

    // Check if this is a bulk sync or single entry
    if (data.action === 'sync') {
      return syncAllData(sheet, data.referrals);
    } else if (data.action === 'add') {
      return addSingleReferral(sheet, data.referral);
    } else if (data.action === 'update') {
      return updateReferral(sheet, data.referral, data.index);
    } else if (data.action === 'delete') {
      return deleteReferral(sheet, data.index);
    } else if (data.action === 'emailReport') {
      return emailReferralReport(data.email, data.referrals);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Invalid action'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function normalizeRewardStatus(value) {
  const status = String(value || '').trim().toLowerCase();
  if (status === 'yes' || status === 'paid') return 'Paid';
  if (status === 'approved') return 'Approved';
  return 'Pending';
}

function normalizeYesNo(value) {
  return String(value || '').trim().toLowerCase() === 'yes' ? 'Yes' : 'No';
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Referral Tracker')
    .addItem('Set up columns', 'setupReferralSheetColumns')
    .addToUi();
}

function setupReferralSheetColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders(sheet);
  return `Referral Tracker columns are ready on "${sheet.getName()}".`;
}

function reportHeaders() {
  return ['No.', 'Referrer Name/Recommending Agent', 'Membership Number', 'Referred Member Name', 'Admission Year', 'Referred Email', 'Admission Fee Paid', 'Admission Fee Date', 'Reward Status', 'Reward Date', 'Date'];
}

function csvCell(value) {
  return `"${String(value == null ? '' : value).replace(/"/g, '""')}"`;
}

function buildReferralReportCsv(referrals) {
  const rows = referrals.map((referral, index) => [
    index + 1,
    referral.referrerName,
    referral.membershipNumber || '',
    referral.referredName,
    referral.admissionYear || '',
    referral.referredEmail,
    normalizeYesNo(referral.admissionFeePaid),
    referral.admissionFeeDate || '',
    normalizeRewardStatus(referral.rewardGiven),
    referral.rewardDate || '',
    referral.dateAdded || ''
  ]);

  return [reportHeaders()].concat(rows)
    .map(row => row.map(csvCell).join(','))
    .join('\n');
}

function sheetHeaders() {
  return reportHeaders();
}

function getHeaderIndex(headers, names) {
  for (let i = 0; i < names.length; i++) {
    const index = headers.indexOf(names[i]);
    if (index !== -1) return index;
  }
  return -1;
}

function getRowValue(row, headers, names, fallbackIndex) {
  const index = getHeaderIndex(headers, names);
  if (index !== -1) return row[index];
  return fallbackIndex >= 0 ? row[fallbackIndex] : '';
}

function sheetRowToReferral(row, headers) {
  const fallbackAdmissionYear = row[5] && !String(row[5]).includes('@') ? row[5] : '';
  const fallbackEmail = row[6] && String(row[6]).includes('@') ? row[6] : row[5];

  return {
    referrerName: getRowValue(row, headers, ['Referrer Name/Recommending Agent'], 1),
    membershipNumber: getRowValue(row, headers, ['Membership Number'], 2),
    referredName: getRowValue(row, headers, ['Referred Member Name'], 3),
    referredMembershipNo: getRowValue(row, headers, ['Referred Membership No.'], -1),
    admissionYear: getRowValue(row, headers, ['Admission Year'], -1) || fallbackAdmissionYear,
    referredEmail: getRowValue(row, headers, ['Referred Email'], -1) || fallbackEmail,
    admissionFeePaid: normalizeYesNo(getRowValue(row, headers, ['Admission Fee Paid'], -1)),
    admissionFeeDate: getRowValue(row, headers, ['Admission Fee Date'], -1),
    rewardGiven: normalizeRewardStatus(getRowValue(row, headers, ['Reward Status', 'Reward given (Yes/No)'], 7)),
    rewardDate: getRowValue(row, headers, ['Reward Date'], -1),
    dateAdded: getRowValue(row, headers, ['Date'], 8),
    rewardAmount: getRowValue(row, headers, ['Reward Amount'], -1)
  };
}

function referralToSheetRow(referral, index) {
  return [
    index + 1,
    referral.referrerName,
    referral.membershipNumber || '',
    referral.referredName,
    referral.admissionYear || '',
    referral.referredEmail,
    normalizeYesNo(referral.admissionFeePaid),
    referral.admissionFeeDate || '',
    normalizeRewardStatus(referral.rewardGiven),
    referral.rewardDate || '',
    referral.dateAdded || new Date().toLocaleDateString()
  ];
}

function emailReferralReport(email, referrals) {
  if (!email || !String(email).match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'A valid email address is required'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (!referrals || referrals.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'No referrals available to email'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const csv = buildReferralReportCsv(referrals);
  const fileName = `IoD_Referrals_${today}.csv`;

  MailApp.sendEmail({
    to: String(email).trim(),
    subject: `IoD Referral Report - ${today}`,
    body: 'Please find attached the latest IoD referral CSV report.',
    attachments: [
      Utilities.newBlob(csv, 'text/csv', fileName)
    ]
  });

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'CSV report emailed successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function ensureHeaders(sheet) {
  const headers = sheetHeaders();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const lastRow = sheet.getLastRow();
  const currentWidth = Math.max(sheet.getLastColumn(), headers.length);
  const currentHeaders = sheet.getRange(1, 1, 1, currentWidth).getValues()[0];
  const needsMigration = headers.some((header, index) => currentHeaders[index] !== header);

  if (needsMigration) {
    const existingRows = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, currentWidth).getValues()
      : [];
    const referrals = existingRows.map(row => sheetRowToReferral(row, currentHeaders));
    sheet.getRange(1, 1, lastRow, currentWidth).clearContent();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (referrals.length > 0) {
      sheet.getRange(2, 1, referrals.length, headers.length)
        .setValues(referrals.map((referral, index) => referralToSheetRow(referral, index)));
    }
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.showColumns(1, headers.length);
  sheet.autoResizeColumns(1, headers.length);
}

function syncAllData(sheet, referrals) {
  ensureHeaders(sheet);

  // Get existing data
  const lastRow = sheet.getLastRow();
  const existingEmails = [];

  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, sheetHeaders().length).getValues();
    existingData.forEach(row => {
      existingEmails.push(row[5]); // Email is in column 6 (index 5)
    });
  }

  // Add only new referrals that don't exist
  let addedCount = 0;
  referrals.forEach((referral, index) => {
    // Check if this email already exists in the sheet
    if (!existingEmails.includes(referral.referredEmail)) {
      const rowNumber = sheet.getLastRow() + 1;
      sheet.appendRow(referralToSheetRow(referral, rowNumber - 2));
      addedCount++;
    }
  });

  // Renumber all rows
  const totalRows = sheet.getLastRow();
  if (totalRows > 1) {
    for (let i = 2; i <= totalRows; i++) {
      sheet.getRange(i, 1).setValue(i - 1);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: addedCount > 0 ? `${addedCount} new referral(s) added` : 'All data already synced',
    count: addedCount
  })).setMimeType(ContentService.MimeType.JSON);
}

function addSingleReferral(sheet, referral) {
  ensureHeaders(sheet);
  const rowNumber = sheet.getLastRow() + 1;
  sheet.appendRow(referralToSheetRow(referral, rowNumber - 2));

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Referral added successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function updateReferral(sheet, referral, index) {
  ensureHeaders(sheet);
  const row = index + 2; // +2 because of header row and 0-based index

  if (row > sheet.getLastRow()) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Row not found'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  sheet.getRange(row, 2).setValue(referral.referrerName);
  sheet.getRange(row, 3).setValue(referral.membershipNumber);
  sheet.getRange(row, 4).setValue(referral.referredName);
  sheet.getRange(row, 5).setValue(referral.admissionYear);
  sheet.getRange(row, 6).setValue(referral.referredEmail);
  sheet.getRange(row, 7).setValue(normalizeYesNo(referral.admissionFeePaid));
  sheet.getRange(row, 8).setValue(referral.admissionFeeDate || '');
  sheet.getRange(row, 9).setValue(normalizeRewardStatus(referral.rewardGiven));
  sheet.getRange(row, 10).setValue(referral.rewardDate || '');
  // Date column (11) remains unchanged during updates

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Referral updated successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function deleteReferral(sheet, index) {
  const row = index + 2; // +2 because of header row and 0-based index

  if (row > sheet.getLastRow()) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Row not found'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  sheet.deleteRow(row);

  // Renumber all rows
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    sheet.getRange(i, 1).setValue(i - 1);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Referral deleted successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ensureHeaders(sheet);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      // No data, return empty array
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: []
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Get all data except header
    const data = sheet.getRange(2, 1, lastRow - 1, sheetHeaders().length).getValues();

    // Convert to array of objects
    const referrals = data.map(row => sheetRowToReferral(row, sheetHeaders()));

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: referrals
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString(),
      data: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
