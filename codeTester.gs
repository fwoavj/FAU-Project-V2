// ============================================================================
// MAIN SUBMISSION AND UI FUNCTIONS
// ============================================================================

/**
 * Main entry point for form submission
 */
function savePersonnelData(formData) {
  try {
    Logger.log('📝 Starting personnel data submission');
    Logger.log('Form data received: ' + JSON.stringify(formData, null, 2));

    // Validate form data
    const validation = validateFormData(formData);
    
    if (!validation.isValid) {
      const errorMessages = validation.errors.map(e => `${e.field}: ${e.message}`).join('\n');
      return {
        success: false,
        message: '❌ Validation failed:\n\n' + errorMessages
      };
    }
    
    if (validation.warnings.length > 0) {
      const warningMessages = validation.warnings.map(w => `${w.field}: ${w.message}`).join('\n');
      Logger.log('⚠️ Warnings:\n' + warningMessages);
    }
    
    // Check for duplicates
    const dupCheck = checkDuplicatePrincipal(formData.fullName, formData.dateOfBirth);
    if (dupCheck.isDuplicate) {
      return {
        success: false,
        message: '❌ Duplicate detected:\n\n' + dupCheck.message
      };
    }
    
    // Process file uploads
    Logger.log('Processing file uploads...');
    formData = processAllFiles(formData);
    
    // Get spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    
    if (!sheet) {
      throw new Error('Sheet "Personnel Tracking" not found');
    }

    // Prepare data arrays
    const pIdx = SYSTEM_CONFIG.COLUMNS.PRINCIPAL._INDICES;
    const dIdx = SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES;
    const sIdx = SYSTEM_CONFIG.COLUMNS.STAFF._INDICES;

    // Create principal base data
    const principalBaseData = new Array(65).fill('');
    principalBaseData[pIdx.POST_STATION] = formData.postStation || '';
    principalBaseData[pIdx.FULL_NAME] = formData.fullName || '';
    principalBaseData[pIdx.RANK] = formData.rank || '';
    principalBaseData[pIdx.DESIGNATION] = formData.designation || '';
    principalBaseData[pIdx.DATE_OF_BIRTH] = formData.dateOfBirth ? parseDate_ddMMyyyy(formData.dateOfBirth) : '';
    principalBaseData[pIdx.AGE] = calculateAgeFromDateString(formData.dateOfBirth);
    principalBaseData[pIdx.SEX] = formData.sex || '';
    principalBaseData[pIdx.ASSUMPTION_DATE] = formData.assumptionDate ? parseDate_ddMMyyyy(formData.assumptionDate) : '';
    principalBaseData[pIdx.PASSPORT_NUMBER] = formData.principalPassport || '';
    principalBaseData[pIdx.PASSPORT_EXPIRATION] = formData.principalPassportExp ? parseDate_ddMMyyyy(formData.principalPassportExp) : '';
    principalBaseData[pIdx.PASSPORT_URL] = formData.principalPassportUrl || '';
    principalBaseData[pIdx.VISA_NUMBER] = formData.principalVisaNumber || '';
    principalBaseData[pIdx.VISA_EXPIRATION] = formData.principalVisaExp ? parseDate_ddMMyyyy(formData.principalVisaExp) : '';
    principalBaseData[pIdx.DIPLOMATIC_ID] = formData.principalDipId || '';
    principalBaseData[pIdx.DIPLOMATIC_ID_EXP] = formData.principalDipIdExp ? parseDate_ddMMyyyy(formData.principalDipIdExp) : '';
    principalBaseData[pIdx.DEPARTURE_DATE] = formData.principalDepartureDate ? parseDate_ddMMyyyy(formData.principalDepartureDate) : '';
    principalBaseData[pIdx.SOLO_PARENT] = formData.soloParent || 'No';
    principalBaseData[pIdx.SOLO_PARENT_URL] = formData.soloParentUrl || '';
    principalBaseData[pIdx.EXTENDED] = formData.extended || '';
    principalBaseData[pIdx.CURRENT_DEPARTURE_DATE] = formData.currentDepartureDate ? parseDate_ddMMyyyy(formData.currentDepartureDate) : '';
    principalBaseData[pIdx.EXTENSION_DETAILS] = formData.extensionDetails || '';

    // Create dependent data arrays
    const dependentDataList = (formData.dependents || []).map(dep => {
      const depRowArray = new Array(65).fill('');

      let depFullName = dep.lastName + ', ' + dep.firstName;
      if (dep.middleName) depFullName += ' ' + dep.middleName;
      if (dep.suffix) depFullName += ' ' + dep.suffix;

      depRowArray[dIdx.FULL_NAME] = depFullName;
      depRowArray[dIdx.RELATIONSHIP] = dep.relationship || '';
      depRowArray[dIdx.SEX] = dep.sex || '';
      depRowArray[dIdx.DATE_OF_BIRTH] = dep.dateOfBirth ? parseDate_ddMMyyyy(dep.dateOfBirth) : '';
      depRowArray[dIdx.AGE] = dep.age || calculateAgeFromDateString(dep.dateOfBirth);
      
      if (dep.dateOfBirth) {
        const depDOBDate = parseDate_ddMMyyyy(dep.dateOfBirth);
        if (depDOBDate) {
          const turns18 = new Date(depDOBDate);
          turns18.setFullYear(turns18.getFullYear() + 18);
          depRowArray[dIdx.TURNS_18_DATE] = turns18;
        }
      }
      
      depRowArray[dIdx.AT_POST] = dep.atPost || 'Yes';
      depRowArray[dIdx.NOTICE_OF_ARRIVAL] = dep.noticeOfArrivalDate ? parseDate_ddMMyyyy(dep.noticeOfArrivalDate) : '';
      depRowArray[dIdx.FAMILY_ALLOWANCE] = dep.receivesFamilyAllowance || 'No';
      depRowArray[dIdx.PASSPORT_NUMBER] = dep.passport || '';
      depRowArray[dIdx.PASSPORT_EXPIRATION] = dep.passportExp ? parseDate_ddMMyyyy(dep.passportExp) : '';
      depRowArray[dIdx.PASSPORT_URL] = dep.passportUrl || '';
      depRowArray[dIdx.VISA_NUMBER] = dep.visaNumber || '';
      depRowArray[dIdx.VISA_EXPIRATION] = dep.visaExp ? parseDate_ddMMyyyy(dep.visaExp) : '';
      depRowArray[dIdx.DIPLOMATIC_ID] = dep.dipId || '';
      depRowArray[dIdx.DIPLOMATIC_ID_EXP] = dep.dipIdExp ? parseDate_ddMMyyyy(dep.dipIdExp) : '';
      depRowArray[dIdx.DEPARTURE_DATE] = dep.departureDate ? parseDate_ddMMyyyy(dep.departureDate) : '';
      depRowArray[dIdx.PWD_STATUS] = dep.pwdStatus || 'No';
      depRowArray[dIdx.PWD_URL] = dep.pwdUrl || '';
      depRowArray[dIdx.APPROVAL_FAX_URL] = dep.approvalFaxUrl || '';
      
      return depRowArray;
    });

    // Create staff data arrays
    const staffDataList = (formData.privateStaff || []).map(staff => {
      const staffRowArray = new Array(65).fill('');

      let staffFullName = staff.lastName + ', ' + staff.firstName;
      if (staff.middleName) staffFullName += ' ' + staff.middleName;
      if (staff.suffix) staffFullName += ' ' + staff.suffix;

      staffRowArray[sIdx.FULL_NAME] = staffFullName;
      staffRowArray[sIdx.SEX] = staff.sex || '';
      staffRowArray[sIdx.DATE_OF_BIRTH] = staff.dateOfBirth ? parseDate_ddMMyyyy(staff.dateOfBirth) : '';
      staffRowArray[sIdx.AGE] = staff.age || calculateAgeFromDateString(staff.dateOfBirth);
      staffRowArray[sIdx.AT_POST] = staff.atPost || 'Yes';
      staffRowArray[sIdx.ARRIVAL_DATE] = staff.arrivalDate ? parseDate_ddMMyyyy(staff.arrivalDate) : '';
      staffRowArray[sIdx.PASSPORT_NUMBER] = staff.passport || '';
      staffRowArray[sIdx.PASSPORT_EXPIRATION] = staff.passportExp ? parseDate_ddMMyyyy(staff.passportExp) : '';
      staffRowArray[sIdx.PASSPORT_URL] = staff.passportUrl || '';
      staffRowArray[sIdx.VISA_NUMBER] = staff.visaNumber || '';
      staffRowArray[sIdx.VISA_EXPIRATION] = staff.visaExp ? parseDate_ddMMyyyy(staff.visaExp) : '';
      staffRowArray[sIdx.DIPLOMATIC_ID] = staff.dipId || '';
      staffRowArray[sIdx.DIPLOMATIC_ID_EXP] = staff.dipIdExp ? parseDate_ddMMyyyy(staff.dipIdExp) : '';
      staffRowArray[sIdx.DEPARTURE_DATE] = staff.departureDate ? parseDate_ddMMyyyy(staff.departureDate) : '';
      staffRowArray[sIdx.PWD_STATUS] = staff.pwdStatus || 'No';
      staffRowArray[sIdx.PWD_URL] = staff.pwdUrl || '';
      staffRowArray[sIdx.EMERGENCY_CONTACT] = staff.emergencyContact || '';
      
      return staffRowArray;
    });

    // Combine and write rows
    const rowsToWrite = [];
    const maxRows = Math.max(1, dependentDataList.length, staffDataList.length);
    const principalTargetRow = sheet.getLastRow() + 1;

    for (let i = 0; i < maxRows; i++) {
      const finalRow = [...principalBaseData];
      
      const depData = dependentDataList[i];
      const staffData = staffDataList[i];

      if (depData) {
        for (let j = dIdx.FULL_NAME; j <= dIdx.EXTENSION_DETAILS; j++) {
          if (depData[j]) {
            finalRow[j] = depData[j];
          }
        }
      }
      
      if (staffData) {
        for (let j = sIdx.FULL_NAME; j <= sIdx.EXTENSION_DETAILS; j++) {
          if (staffData[j]) {
            finalRow[j] = staffData[j];
          }
        }
      }

      rowsToWrite.push(finalRow);
    }

    // Write all rows at once
    if (rowsToWrite.length > 0) {
      sheet.getRange(principalTargetRow, 1, rowsToWrite.length, 65).setValues(rowsToWrite);
      Logger.log(`✅ ${rowsToWrite.length} row(s) for ${formData.fullName} recorded starting row ${principalTargetRow}`);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message: `✅ Personnel data recorded successfully starting row ${principalTargetRow}`,
      row: principalTargetRow
    };

  } catch (error) {
    Logger.log('❌ Error in savePersonnelData: ' + error);
    Logger.log('Error Stack: ' + error.stack);
    return {
      success: false,
      message: 'Failed to save data: ' + error.message
    };
  }
}

/**
 * Calculate age from date string
 */
function calculateAgeFromDateString(dateString) {
  if (!dateString) return '';
  
  try {
    const birthDate = new Date(dateString);
    if (isNaN(birthDate.getTime())) return '';
    
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
      age--;
    }
    
    return age >= 0 ? age : '';
  } catch (error) {
    return '';
  }
}

/**
 * Parse date in dd/mm/yyyy format
 */
function parseDate_ddMMyyyy(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  const parts = dateString.split(/[\/\-]/);

  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);

    if (!isNaN(day) && !isNaN(month) && !isNaN(year) && month >= 1 && month <= 12) {
      const daysInMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
      if (day >= 1 && day <= daysInMonth[month - 1]) {
        return new Date(year, month - 1, day);
      }
    }
  }

  const fallbackDate = new Date(dateString);
  if (!isNaN(fallbackDate.getTime())) {
    Logger.log("Warning: parseDate_ddMMyyyy used fallback for input: " + dateString);
    return fallbackDate;
  }

  Logger.log("Error: parseDate_ddMMyyyy failed to parse input: " + dateString);
  return null;
}

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

function createTrackingSheet(spreadsheet = null) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.insertSheet('Personnel Tracking');

    const maxColumns = trackingSheet.getMaxColumns();
    if (maxColumns < 65) {
      trackingSheet.insertColumnsAfter(maxColumns, 65 - maxColumns);
    }

    // Group headers
    trackingSheet.getRange(1, 1, 1, 21).merge().setValue('Principals')
      .setFontSize(12).setFontWeight('bold').setBackground('#a4c2f4')
      .setFontColor('black').setHorizontalAlignment('center');

    trackingSheet.getRange(1, 22, 1, 23).merge().setValue('Dependents')
      .setFontSize(12).setFontWeight('bold').setBackground('#93CCEA')
      .setFontColor('black').setHorizontalAlignment('center');

    trackingSheet.getRange(1, 46, 1, 20).merge().setValue('Private Staff')
      .setFontSize(12).setFontWeight('bold').setBackground('#f4cccc')
      .setFontColor('black').setHorizontalAlignment('center');

    const headers = [
      // Principal (1-21)
      'Post/Station', 'Principal Full Name', 'Rank', 'Designation', 'Date of Birth', 'Age', 'Sex',
      'Assumption Date', 'Principal Passport', 'Passport Expiration', 'Passport URL', 'Visa Number', 
      'Visa Expiration', 'Diplomatic/Consular ID', 'ID Expiration', 'Departure Date', 'Solo Parent', 
      'Solo Parent URL', 'Extended?', 'New Departure Date', 'Extension Details',
      
      // Dependent (22-44)
      'Dependent Name', 'Relationship', 'Sex', 'Date of Birth', 'Age', 'Turns 18 Date', 'At Post',
      'Notice of Arrival Date', 'Receives Family Allowance', 'Dependent Passport', 'Dependent Passport Exp', 
      'Passport URL', 'Visa Number', 'Visa Expiration', 'Diplomatic/Consular ID', 'ID Expiration', 
      'Departure Date', 'PWD Status', 'PWD URL', 'Approval Fax URL', 'Extended?', 'New Departure Date', 
      'Extension Details',
      
      // Staff (46-65)
      'Staff Name', 'Sex', 'Date of Birth', 'Age', 'At Post', 'Arrival Date',
      'Staff Passport', 'Staff Passport Exp', 'Passport URL', 'Visa Number', 'Visa Expiration', 
      'Diplomatic/Consular ID', 'ID Expiration', 'Departure Date', 'PWD Status', 'PWD URL', 
      'Emergency Contact', 'Extended?', 'New Departure Date', 'Extension Details'
    ];
  
    trackingSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  
    const headerRange = trackingSheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#a4c2f4').setFontColor('black').setFontWeight('bold')
      .setWrap(true).setHorizontalAlignment('center');

    trackingSheet.getRange(2, 22, 1, 23).setBackground('#93ccea').setFontColor('black');
    trackingSheet.getRange(2, 46, 1, 20).setBackground('#f4cccc').setFontColor('black');
  
    const columnWidths = [
      // Principal (1-21)
      200, 150, 180, 150, 120, 60, 80, 120, 120, 120, 250, 120, 120, 120, 120, 120, 100, 250, 100, 140, 300,
      
      // Dependent (22-44)
      200, 120, 80, 120, 60, 120, 80, 120, 120, 120, 120, 250, 120, 120, 120, 120, 120, 80, 250, 250, 100, 140, 300,
      
      // Staff (46-65)
      200, 80, 120, 60, 80, 120, 120, 120, 250, 120, 120, 120, 120, 120, 80, 250, 150, 100, 140, 300
    ];
    
    columnWidths.forEach((width, index) => {
      trackingSheet.setColumnWidth(index + 1, width);
    });
    
    const extendedValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'])
      .setAllowInvalid(false)
      .build();
    
    trackingSheet.getRange('S3:S').setDataValidation(extendedValidation);
    trackingSheet.getRange('AP3:AP').setDataValidation(extendedValidation);
    trackingSheet.getRange('BK3:BK').setDataValidation(extendedValidation);
    
    trackingSheet.getRange('T3:T').setNumberFormat('yyyy-mm-dd');
    trackingSheet.getRange('AQ3:AQ').setNumberFormat('yyyy-mm-dd');
    trackingSheet.getRange('BL3:BL').setNumberFormat('yyyy-mm-dd');
    
    trackingSheet.setFrozenRows(2);
    trackingSheet.setFrozenColumns(2);
    
    console.log('✅ Personnel Tracking sheet created successfully with 65 columns');
    return trackingSheet;
    
  } catch (error) {
    console.error('❌ Error creating tracking sheet:', error);
    throw error;
  }
}

function createPrincipalsSheet(spreadsheet = null) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const principalsSheet = ss.insertSheet('Principals List');

    const headers = ['Post/Station', 'Full Name', 'Date of Birth'];
    principalsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const headerRange = principalsSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a86e8').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center');

    principalsSheet.setColumnWidth(1, 200);
    principalsSheet.setColumnWidth(2, 200);
    principalsSheet.setColumnWidth(3, 120);
    
    principalsSheet.setFrozenRows(1);
    
    console.log('✅ Principals List sheet created successfully');
    return principalsSheet;
    
  } catch (error) {
    console.error('❌ Error creating principals sheet:', error);
    throw error;
  }
}

// ============================================================================
// UI FUNCTIONS
// ============================================================================

function showEntryForm() {
  const html = HtmlService.createHtmlOutputFromFile('EntryForm')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Personnel Entry Form');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Personnel Entry Form');
}

function showUpdateForm() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('UpdateForm')
      .setWidth(900)
      .setHeight(700)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '✏️ Update Personnel');
  } catch (error) {
    console.error('Error showing update form:', error);
    SpreadsheetApp.getUi().alert('Error opening form: ' + error.toString());
  }
}

function showTrackingSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let trackingSheet = ss.getSheetByName('Personnel Tracking');
  
    if (!trackingSheet) {
      trackingSheet = createTrackingSheet(ss);
    }
  
    trackingSheet.activate();
  
    if (trackingSheet.getLastRow() <= 2) {
      SpreadsheetApp.getUi().alert('📋 No personnel records found. Use "Open Entry Form" to add data.');
    }
  } catch (error) {
    console.error('Error showing tracking sheet:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function showPrincipalsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let principalsSheet = ss.getSheetByName('Principals List');
  
    if (!principalsSheet) {
      principalsSheet = createPrincipalsSheet(ss);
    }
  
    principalsSheet.activate();
  } catch (error) {
    console.error('Error showing principals sheet:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

function openAttendanceTracker() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('attendanceForm')
      .setWidth(1200)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '📊 Attendance Tracker');
  } catch (error) {
    console.error('Error showing attendance tracker:', error);
    SpreadsheetApp.getUi().alert('Error opening attendance tracker: ' + error.toString());
  }
}

function openArchiveViewer() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('archive')
      .setWidth(1200)
      .setHeight(800)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '📦 Archived Withdrawals');
  } catch (error) {
    console.error('Error showing archive viewer:', error);
    SpreadsheetApp.getUi().alert('Error opening archive: ' + error.toString());
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Personnel Tracker')
    .addItem('Open Entry Form', 'showEntryForm')
    .addItem('📊 Attendance Tracker', 'openAttendanceTracker')
    .addItem('View Records', 'showTrackingSheet')
    .addItem('Extension Form', 'showUpdateForm')
    .addSeparator()
    .addItem('Manage Principals', 'showPrincipalsSheet')
    .addItem('Update Principal List', 'updatePrincipal')
    .addToUi();
}

function updatePrincipal() {
  SpreadsheetApp.getUi().alert('Use "Manage Principals" to view and edit the Principals List sheet directly.');
}

// ============================================================================
// PRINCIPALS LIST FUNCTIONS
// ============================================================================

function getPrincipalsList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Principals List");
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    const principals = data
      .filter(row => row[1])
      .map(row => {
        let isoDob = "";
        const rawDob = row[2];

        if (rawDob instanceof Date) {
          isoDob = Utilities.formatDate(rawDob, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (typeof rawDob === "string" && rawDob.trim() !== "") {
          const parts = rawDob.split(/[\/\-]/);
          if (parts.length === 3) {
            const [mm, dd, yyyy] = parts;
            isoDob = `${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;
          }
        }

        return {
          postStation: (row[0] || "").toString().trim(),
          fullName: (row[1] || "").toString().trim(),
          dateOfBirth: isoDob
        };
      });

    Logger.log(JSON.stringify(principals.slice(0, 3), null, 2));
    return principals;

  } catch (err) {
    Logger.log("❌ getPrincipalsList error: " + err);
    return [];
  }
}

function getAllPrincipals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Personnel Tracking');
    
    if (!sheet) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const principals = [];
    const seenPrincipals = new Set();
    
    for (let i = 2; i < data.length; i++) {
      const principalName = data[i][1];
      const post = data[i][0];
      
      if (principalName && principalName !== '' && !seenPrincipals.has(principalName)) {
        seenPrincipals.add(principalName);
        principals.push({
          fullName: principalName,
          post: post || ''
        });
      }
    }

    return principals;

  } catch (error) {
    console.error('Error getting principals:', error);
    return [];
  }
}

/**
 * Find all rows for a principal
 */
function findPrincipalRows(sheet, principalName) {
  const nameColumn = 2;
  const data = sheet.getRange(3, nameColumn, sheet.getLastRow() - 2, 1).getValues();
  
  const matchingRows = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === principalName) {
      matchingRows.push(i + 3);
    }
  }
  return matchingRows;
}

/**
 * Find first principal row (for backward compatibility)
 */
function findPrincipalRow(sheet, principalName) {
  const rows = findPrincipalRows(sheet, principalName);
  return rows.length > 0 ? rows[0] : null;
}
