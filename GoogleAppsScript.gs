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
  const headers = ['No.', 'Referrer Name/Recommending Agent', 'Membership Number', 'Referred Member Name', 'Referred Membership No.', 'Admission Year', 'Referred Email', 'Reward Status', 'Date', 'Reward Amount', 'Reward Date', 'Admission Fee Paid', 'Admission Fee Date'];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  headers.forEach((header, index) => {
    if (currentHeaders[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });

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
    const existingData = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    existingData.forEach(row => {
      existingEmails.push(row[6]); // Email is in column 7 (index 6)
    });
  }

  // Add only new referrals that don't exist
  let addedCount = 0;
  referrals.forEach((referral, index) => {
    // Check if this email already exists in the sheet
    if (!existingEmails.includes(referral.referredEmail)) {
      const rowNumber = sheet.getLastRow() + 1;
      sheet.appendRow([
        rowNumber - 1,
        referral.referrerName,
        referral.membershipNumber,
        referral.referredName,
        referral.referredMembershipNo,
        referral.admissionYear,
        referral.referredEmail,
        normalizeRewardStatus(referral.rewardGiven),
        referral.dateAdded || new Date().toLocaleDateString(),
        referral.rewardAmount || '',
        referral.rewardDate || '',
        normalizeYesNo(referral.admissionFeePaid),
        referral.admissionFeeDate || ''
      ]);
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
  sheet.appendRow([
    rowNumber - 1,
    referral.referrerName,
    referral.membershipNumber,
    referral.referredName,
    referral.referredMembershipNo,
    referral.admissionYear,
    referral.referredEmail,
    normalizeRewardStatus(referral.rewardGiven),
    referral.dateAdded || new Date().toLocaleDateString(),
    referral.rewardAmount || '',
    referral.rewardDate || '',
    normalizeYesNo(referral.admissionFeePaid),
    referral.admissionFeeDate || ''
  ]);

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
  sheet.getRange(row, 5).setValue(referral.referredMembershipNo);
  sheet.getRange(row, 6).setValue(referral.admissionYear);
  sheet.getRange(row, 7).setValue(referral.referredEmail);
  sheet.getRange(row, 8).setValue(normalizeRewardStatus(referral.rewardGiven));
  // Date column (9) remains unchanged during updates
  sheet.getRange(row, 10).setValue(referral.rewardAmount || '');
  sheet.getRange(row, 11).setValue(referral.rewardDate || '');
  sheet.getRange(row, 12).setValue(normalizeYesNo(referral.admissionFeePaid));
  sheet.getRange(row, 13).setValue(referral.admissionFeeDate || '');

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
    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

    // Convert to array of objects
    const referrals = data.map(row => ({
      referrerName: row[1],
      membershipNumber: row[2],
      referredName: row[3],
      referredMembershipNo: row[4],
      admissionYear: row[5],
      referredEmail: row[6],
      rewardGiven: normalizeRewardStatus(row[7]),
      dateAdded: row[8],
      rewardAmount: row[9],
      rewardDate: row[10],
      admissionFeePaid: normalizeYesNo(row[11]),
      admissionFeeDate: row[12]
    }));

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
