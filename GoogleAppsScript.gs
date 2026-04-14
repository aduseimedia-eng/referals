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

function syncAllData(sheet, referrals) {
  // Set up headers if not present
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['No.', 'Referrer Name/Recommending Agent', 'Membership Number', 'Referred Member Name', 'Referred Membership No.', 'Admission Year', 'Referred Email', 'Reward given (Yes/No)', 'Date']);
  }
  
  // Get existing data
  const lastRow = sheet.getLastRow();
  const existingEmails = [];
  
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
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
        referral.rewardGiven,
        new Date().toLocaleDateString()
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
  // Set up headers if not present
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['No.', 'Referrer Name/Recommending Agent', 'Membership Number', 'Referred Member Name', 'Referred Membership No.', 'Admission Year', 'Referred Email', 'Reward given (Yes/No)', 'Date']);
  }
  const rowNumber = sheet.getLastRow() + 1;
  sheet.appendRow([
    rowNumber - 1,
    referral.referrerName,
    referral.membershipNumber,
    referral.referredName,
    referral.referredMembershipNo,
    referral.admissionYear,
    referral.referredEmail,
    referral.rewardGiven,
    new Date().toLocaleDateString()
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Referral added successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

function updateReferral(sheet, referral, index) {
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
  sheet.getRange(row, 8).setValue(referral.rewardGiven);
  // Date column (9) remains unchanged during updates
  
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
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      // No data, return empty array
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get all data except header
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    
    // Convert to array of objects
    const referrals = data.map(row => ({
      referrerName: row[1],
      membershipNumber: row[2],
      referredName: row[3],
      referredMembershipNo: row[4],
      admissionYear: row[5],
      referredEmail: row[6],
      rewardGiven: row[7],
      dateAdded: row[8]
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
