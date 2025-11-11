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
    const validation = validateForm_savePersonnel(formData); // Using our uniquely named validator
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
    formData = processAllFiles(formData); // This is in FileUpload.gs
    
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

    // Create principal base data - UPDATED TO 64 COLUMNS
    const principalBaseData = new Array(64).fill('');
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
    principalBaseData[pIdx.EXTENDED] = formData.extended || ''; // This is column 19 (S)
    principalBaseData[pIdx.CURRENT_DEPARTURE_DATE] = formData.currentDepartureDate ? parseDate_ddMMyyyy(formData.currentDepartureDate) : ''; // Col 20
    principalBaseData[pIdx.EXTENSION_DETAILS] = formData.extensionDetails || ''; // Col 21

    // Create dependent data arrays
    const dependentDataList = (formData.dependents || []).map(dep => {
      const depRowArray = new Array(64).fill(''); // UPDATED TO 64 COLUMNS

      let depFullName = dep.lastName + ', ' + dep.firstName;
      if (dep.middleName) depFullName += ' ' + dep.middleName;
      if (dep.suffix) depFullName += ' ' + dep.suffix;

      depRowArray[dIdx.FULL_NAME] = depFullName; // Index 21 (Col 22)
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
      depRowArray[dIdx.APPROVAL_FAX_URL] = dep.approvalFaxUrl || ''; // Index 40 (Col 41)
      depRowArray[dIdx.EXTENDED] = dep.extended || '';
      depRowArray[dIdx.CURRENT_DEPARTURE_DATE] = dep.currentDepartureDate ? parseDate_ddMMyyyy(dep.currentDepartureDate) : '';
      depRowArray[dIdx.EXTENSION_DETAILS] = dep.extensionDetails || ''; // Index 43 (Col 44)
      
      return depRowArray;
    });

    // Create staff data arrays
    const staffDataList = (formData.privateStaff || []).map(staff => {
      const staffRowArray = new Array(64).fill(''); // UPDATED TO 64 COLUMNS

      let staffFullName = staff.lastName + ', ' + staff.firstName;
      if (staff.middleName) staffFullName += ' ' + staff.middleName;
      if (staff.suffix) staffFullName += ' ' + staff.suffix;

      staffRowArray[sIdx.FULL_NAME] = staffFullName; // Index 44 (Col 45)
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
      staffRowArray[sIdx.EXTENDED] = staff.extended || '';
      staffRowArray[sIdx.CURRENT_DEPARTURE_DATE] = staff.currentDepartureDate ? parseDate_ddMMyyyy(staff.currentDepartureDate) : '';
      staffRowArray[sIdx.EXTENSION_DETAILS] = staff.extensionDetails || ''; // Index 63 (Col 64)
      
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
        // Copy dependent data (cols 22-44)
        for (let j = dIdx.FULL_NAME; j <= dIdx.EXTENSION_DETAILS; j++) {
          if (depData[j] !== undefined) {
            finalRow[j] = depData[j];
          }
        }
      }
      
      if (staffData) {
        // Copy staff data (cols 45-64)
        for (let j = sIdx.FULL_NAME; j <= sIdx.EXTENSION_DETAILS; j++) {
          if (staffData[j] !== undefined) {
            finalRow[j] = staffData[j];
          }
        }
      }

      rowsToWrite.push(finalRow);
    }

    // Write all rows at once - UPDATED TO 64 COLUMNS
    if (rowsToWrite.length > 0) {
      sheet.getRange(principalTargetRow, 1, rowsToWrite.length, 64).setValues(rowsToWrite);
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
 * Validates the flat formData object from savePersonnelData
 * This function name is unique to this file
 */
function validateForm_savePersonnel(formData) {
  const errors = [];
  const warnings = [];
  const config = SYSTEM_CONFIG.VALIDATION;

  function addError(field, message) {
    errors.push({ field, message });
  }

  // --- Principal Validation ---
  if (!formData.fullName) addError('Principal', 'Full Name is required. (Did you select a principal?)');
  if (!formData.rank) addError('Principal', 'Rank is required.');
  if (!formData.sex) addError('Principal', 'Sex is required.');
  if (!formData.assumptionDate) addError('Principal', 'Date of Assumption is required.');
  if (!formData.principalPassport) addError('Principal', 'Passport No. is required.');
  if (!formData.principalPassportExp) addError('Principal', 'Passport Expiration is required.');

  // --- Dependents Validation ---
  if (formData.dependents && formData.dependents.length > 0) {
    formData.dependents.forEach((dep, index) => {
      const depNum = index + 1;
      if (!dep.lastName || !dep.firstName) addError(`Dependent ${depNum}`, 'First and Last Name are required.');
      if (!dep.relationship) addError(`Dependent ${depNum}`, 'Relationship is required.');
      if (!dep.sex) addError(`Dependent ${depNum}`, 'Sex is required.');
      if (!dep.dateOfBirth) addError(`Dependent ${depNum}`, 'Date of Birth is required.');
      if (!dep.atPost) addError(`Dependent ${depNum}`, '"At Post" status is required.');
      if (!dep.receivesFamilyAllowance) addError(`Dependent ${depNum}`, '"Receives Family Allowance" status is required.');
      if (!dep.passport) addError(`Dependent ${depNum}`, 'Passport No. is required.');
      if (!dep.passportExp) addError(`Dependent ${depNum}`, 'Passport Expiration is required.');
      if (dep.pwdStatus === "") addError(`Dependent ${depNum}`, 'PWD Status is required.'); // Check for empty string
    });
  }

  // --- Private Staff Validation ---
  if (formData.privateStaff && formData.privateStaff.length > 0) {
    formData.privateStaff.forEach((staff, index) => {
      const staffNum = index + 1;
      if (!staff.lastName || !staff.firstName) addError(`Staff ${staffNum}`, 'First and Last Name are required.');
      if (!staff.sex) addError(`Staff ${staffNum}`, 'Sex is required.');
      if (!staff.dateOfBirth) {
        addError(`Staff ${staffNum}`, 'Date of Birth is required.');
      } else {
        // Check age requirement
        const age = calculateAgeFromDateString(staff.dateOfBirth); // This helper function is already in codeTester.gs
        if (age < config.MIN_AGE_STAFF) {
          addError(`Staff ${staffNum}`, `Staff member must be at least ${config.MIN_AGE_STAFF} years old. Current age is ${age}.`);
        }
      }
      if (!staff.atPost) addError(`Staff ${staffNum}`, '"At Post" status is required.');
      if (!staff.passport) addError(`Staff ${staffNum}`, 'Passport No. is required.');
      if (!staff.passportExp) addError(`Staff ${staffNum}`, 'Passport Expiration is required.');
      if (staff.pwdStatus === "") addError(`Staff ${staffNum}`, 'PWD Status is required.'); // Check for empty string
    });
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
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
 * Parse date in dd/mm/yyyy format OR yyyy-mm-dd
 */
function parseDate_ddMMyyyy(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  // Try yyyy-mm-dd first (standard HTML date)
  if (dateString.includes('-')) {
    const parts = dateString.split('-');
    if (parts.length === 3 && parts[0].length === 4) {
      const date = new Date(dateString);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }

  // Try dd/mm/yyyy
  const parts = dateString.split(/[\/\-]/);
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10); // month is 1-based
    const year = parseInt(parts[2], 10);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year) && year > 1900 && month >= 1 && month <= 12) {
       return new Date(year, month - 1, day);
    }
  }

  // Fallback attempt
  const fallbackDate = new Date(dateString);
  if (!isNaN(fallbackDate.getTime())) {
     Logger.log("Warning: parseDate_ddMMyyyy used fallback for input: " + dateString);
    return fallbackDate;
  }

  Logger.log("Error: parseDate_ddMMyyyy failed to parse input: " + dateString);
  return null;
}

// ============================================================================
// SHEET CREATION FUNCTIONS - UPDATED FOR 64 COLUMNS
// ============================================================================

/**
 * Creates the "Personnel Tracking" sheet with 64 columns (A-BL)
 * as per the new CONFIG.gs structure.
 */
function createTrackingSheet(spreadsheet = null) {
  try {
    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.insertSheet('Personnel Tracking');

    // Ensure we have exactly 64 columns
    const maxColumns = trackingSheet.getMaxColumns();
    if (maxColumns < 64) {
      trackingSheet.insertColumnsAfter(maxColumns, 64 - maxColumns);
    } else if (maxColumns > 64) {
      trackingSheet.deleteColumns(65, maxColumns - 64);
    }

    // Group headers (Row 1)
    trackingSheet.getRange(1, 1, 1, 21).merge().setValue('Principals')
      .setFontSize(12).setFontWeight('bold').setBackground('#a4c2f4')
      .setFontColor('black').setHorizontalAlignment('center');
    trackingSheet.getRange(1, 22, 1, 23).merge().setValue('Dependents')
      .setFontSize(12).setFontWeight('bold').setBackground('#93CCEA')
      .setFontColor('black').setHorizontalAlignment('center');
    trackingSheet.getRange(1, 45, 1, 20).merge().setValue('Private Staff')
      .setFontSize(12).setFontWeight('bold').setBackground('#f4cccc')
      .setFontColor('black').setHorizontalAlignment('center');
    
    // Column Headers (Row 2)
    const headers = [
      // Principal (Cols 1-21)
      'Post/Station', 'Principal Full Name', 'Rank', 'Designation', 'Date of Birth', 'Age', 'Sex',
      'Assumption Date', 'Principal Passport', 'Passport Expiration', 'Passport URL', 'Visa Number', 
      'Visa Expiration', 'Diplomatic/Consular ID', 'ID Expiration', 'End of Existing Employment Contract', 'Solo Parent', 
      'Solo Parent URL', 'Extended?', 'New End of Existing Employment Contract', 'Extension Details',
      
      // Dependent (Cols 22-44)
      'Dependent Name', 'Relationship', 'Sex', 'Date of Birth', 'Age', 'Turns 18 Date', 'At Post',
      'Notice of Arrival Date', 'Receives Family Allowance', 'Dependent Passport', 'Dependent Passport Exp', 
      'Passport URL', 'Visa Number', 'Visa Expiration', 'Diplomatic/Consular ID', 'ID Expiration', 
      'End of Existing Employment Contract', 'PWD Status', 'PWD URL', 'Approval Fax URL', 'Extended?', 'New End of Existing Employment Contract', 
      'Extension Details',
            
      // Staff (Cols 45-64)
      'Staff Name', 'Sex', 'Date of Birth', 'Age', 'At Post', 'Arrival Date',
      'Staff Passport', 'Staff Passport Exp', 'Passport URL', 'Visa Number', 'Visa Expiration', 
      'Diplomatic/Consular ID', 'ID Expiration', 'End of Existing Employment Contract', 'PWD Status', 'PWD URL', 
      'Emergency Contact', 'Extended?', 'New End of Existing Employment Contract', 'Extension Details'
    ];
    
    // Write all 64 headers to Row 2
    trackingSheet.getRange(2, 1, 1, 64).setValues([headers]);
  
    // Format Header Row 2
    trackingSheet.getRange(2, 1, 1, 64).setBackground('#a4c2f4').setFontColor('black').setFontWeight('bold')
      .setWrap(true).setHorizontalAlignment('center');
    
    trackingSheet.getRange(2, 22, 1, 23).setBackground('#93ccea'); // Dependent headers
    trackingSheet.getRange(2, 45, 1, 20).setBackground('#f4cccc'); // Staff headers
  
    // Column Widths
    const columnWidths = [
      // Principal (1-21)
      200, 150, 180, 150, 120, 60, 80, 120, 120, 120, 250, 120, 120, 120, 120, 120, 100, 250, 100, 140, 300,
      
      // Dependent (22-44)
      200, 120, 80, 120, 60, 120, 80, 120, 120, 120, 120, 250, 120, 120, 120, 120, 120, 80, 250, 250, 100, 140, 300,
            
      // Staff (45-64)
      200, 80, 120, 60, 80, 120, 120, 120, 250, 120, 120, 120, 120, 120, 80, 250, 150, 100, 140, 300
    ];
    
    // Set all 64 column widths
    columnWidths.forEach((width, index) => {
      trackingSheet.setColumnWidth(index + 1, width);
    });
    
    // Data Validations
    const extendedValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'])
      .setAllowInvalid(false)
      .build();
    
    trackingSheet.getRange('S3:S').setDataValidation(extendedValidation); // Principal Extended
    trackingSheet.getRange('AP3:AP').setDataValidation(extendedValidation); // Dependent Extended
    trackingSheet.getRange('BJ3:BJ').setDataValidation(extendedValidation); // Staff Extended (NOW COL 62/BJ)
    
    // Date Formats
    trackingSheet.getRange('T3:T').setNumberFormat('yyyy-mm-dd');
    trackingSheet.getRange('AQ3:AQ').setNumberFormat('yyyy-mm-dd');
    trackingSheet.getRange('BK3:BK').setNumberFormat('yyyy-mm-dd'); // Staff New Departure (NOW COL 63/BK)
    
    // Freeze Panes
    trackingSheet.setFrozenRows(2);
    trackingSheet.setFrozenColumns(2);
    
    console.log('✅ Personnel Tracking sheet created successfully with 64 columns');
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
    .setWidth(SYSTEM_CONFIG.UI.MODAL_WIDTH)
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
    SpreadsheetApp.getUi().showModalDialog(html, 'Status of Residency');
  } catch (error) {
    console.error('Error showing status of residency:', error);
    SpreadsheetApp.getUi().alert('Error opening status of residency: ' + error.toString());
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
  ui.createMenu('Menu')
    .addItem('Open Entry Form', 'showEntryForm')
    .addItem('Status of Residency', 'openAttendanceTracker')
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

/**
 * Gets the list of principals from the "Principals List" sheet.
 * This is used to populate the Entry Form dropdown.
 */
function getPrincipalsList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PRINCIPALS_LIST);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A:C

    const principals = data
      .filter(row => row[1] && row[1].toString().trim() !== "") // Filter out empty names
      .map(row => {
        let isoDob = "";
        const rawDob = row[2];

        if (rawDob instanceof Date) {
          isoDob = Utilities.formatDate(rawDob, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (typeof rawDob === "string" && rawDob.trim() !== "") {
          const parsed = parseDate_ddMMyyyy(rawDob); // Use our robust parser
          if(parsed) {
            isoDob = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        }

        return {
          postStation: (row[0] || "").toString().trim(),
          fullName: (row[1] || "").toString().trim(),
          dateOfBirth: isoDob
        };
      });
    
    return principals;

  } catch (err) {
    Logger.log("❌ getPrincipalsList error: " + err);
    return [];
  }
}

/**
 * Gets a unique list of principals already in the "Personnel Tracking" sheet.
 * This is used for the Update Form dropdown.
 */
function getAllPrincipals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) {
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    const data = sheet.getRange(3, 1, lastRow - 2, 2).getValues(); // Cols A and B
    const principals = [];
    const seenPrincipals = new Set();
    
    for (let i = 0; i < data.length; i++) {
      const post = data[i][0];
      const principalName = data[i][1];
      
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
 * Finds all rows (by row number) for a specific principal.
 */
function findPrincipalRows(sheet, principalName) {
  const nameColumn = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.FULL_NAME; // Column 2
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  const data = sheet.getRange(3, nameColumn, lastRow - 2, 1).getValues();
  
  const matchingRows = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === principalName) {
      matchingRows.push(i + 3); // +3 because data starts from row 3
    }
  }
  return matchingRows;
}

/**
 * Finds the *first* row for a principal.
 */
function findPrincipalRow(sheet, principalName) {
  const rows = findPrincipalRows(sheet, principalName);
  return rows.length > 0 ? rows[0] : null;
}

/**
 * Checks for a duplicate principal in the tracking sheet.
 */
function checkDuplicatePrincipal(fullName, dateOfBirth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) return { isDuplicate: false };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { isDuplicate: false };

    const nameCol = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.FULL_NAME;
    const dobCol = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.DATE_OF_BIRTH;
    
    const names = sheet.getRange(3, nameCol, lastRow - 2, 1).getValues();
    const dobs = sheet.getRange(3, dobCol, lastRow - 2, 1).getValues();
    const dobString = dateOfBirth ? parseDate_ddMMyyyy(dateOfBirth).toDateString() : '';

    for (let i = 0; i < names.length; i++) {
      if (names[i][0] === fullName) {
        const sheetDob = dobs[i][0] ? new Date(dobs[i][0]).toDateString() : '';
        if (sheetDob === dobString) {
          return {
            isDuplicate: true,
            message: `Principal "${fullName}" with DOB "${dobString}" already exists in row ${i + 3}.`
          };
        }
      }
    }
    return { isDuplicate: false };
  } catch (error) {
    Logger.log("Error checking duplicate: " + error);
    return { isDuplicate: false, message: "Error checking for duplicates: " + error.message };
  }
}

// ============================================================================
// SAMPLE DATA INSERTER
// ============================================================================

/**
 * Run this function from the Apps Script editor to add 7 rows of sample data.
 * It will clear rows 3-9 and insert the new data.
 */
function insertSampleData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Error: "Personnel Tracking" sheet not found. Please create it first (e.g., by running "showTrackingSheet").');
      return;
    }

    // Clear rows 3-9 to make space
    sheet.getRange('A3:BL9').clearContent();

    // The 7 rows of sample data, formatted as a JavaScript array
    const sampleData = [
      // Row 1 (Juan Dela Cruz)
      ["Washington D.C.", "DELA CRUZ, Juan M.", "Foreign Service Officer I", "First Secretary and Consul", new Date("1980-05-15"), 45, "Male", new Date("2023-08-01"), "P123456A", new Date("2028-05-14"), "https://drive.google.com/file/d/principal_pass_1", "V1234567", new Date("2028-08-01"), "D-00123", new Date("2028-08-01"), new Date("2029-07-31"), "No", "", "No", "", "", "DELA CRUZ, Maria L.", "Spouse", "Female", new Date("1982-03-10"), 43, "", "Yes", new Date("2023-08-01"), "Yes", "P765432B", new Date("2028-03-09"), "https://drive.google.com/file/d/dep_pass_1", "V1234568", new Date("2028-08-01"), "D-00124", new Date("2028-08-01"), new Date("2029-07-31"), "No", "", "", "No", "", "", "GARCIA, Ana R.", "Female", new Date("1995-01-20"), 30, "Yes", new Date("2023-08-15"), "S987654C", new Date("2030-01-19"), "https://drive.google.com/file/d/staff_pass_1", "V1234569", new Date("2026-08-15"), "D-00125", new Date("2026-08-15"), new Date("2029-07-31"), "No", "", "Maria Garcia - 09171234567", "No", "", ""],
      // Row 2 (Maria Clara Santos)
      ["Tokyo", "SANTOS, Maria Clara G.", "Chief of Mission II", "Ambassador", new Date("1975-11-02"), 50, "Female", new Date("2024-01-15"), "P234567B", new Date("2030-11-01"), "https://drive.google.com/file/d/principal_pass_2", "V2345678", new Date("2029-01-15"), "D-00201", new Date("2029-01-15"), new Date("2030-01-14"), "No", "", "No", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      // Row 3 (Robert Johnson)
      ["Singapore", "JOHNSON, Robert A.", "Foreign Service Officer IV", "Vice Consul", new Date("1990-02-28"), 35, "Male", new Date("2022-07-10"), "P345678C", new Date("2027-02-27"), "https://drive.google.com/file/d/principal_pass_3", "V3456789", new Date("2027-07-10"), "D-00301", new Date("2027-07-10"), new Date("2028-07-09"), "No", "", "No", "", "", "JOHNSON, Emily K.", "Child", "Female", new Date("2018-06-12"), 7, new Date("2036-06-12"), "Yes", new Date("2022-07-10"), "Yes", "P456789D", new Date("2027-06-11"), "https://drive.google.com/file/d/dep_pass_2", "V3456790", new Date("2027-07-10"), "D-00302", new Date("2027-07-10"), new Date("2028-07-09"), "No", "", "", "No", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      // Row 4 (Johnson's 2nd child)
      ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "JOHNSON, Michael R.", "Child", "Male", new Date("2020-10-30"), 5, new Date("2038-10-30"), "Yes", new Date("2022-07-10"), "Yes", "P567890E", new Date("2027-10-29"), "https://drive.google.com/file/d/dep_pass_3", "V3456791", new Date("2027-07-10"), "D-00303", new Date("2027-07-10"), new Date("2028-07-09"), "No", "", "", "No", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      // Row 5 (Wei Chen)
      ["Berlin", "CHEN, Wei L.", "Foreign Service Staff Officer I", "Administrative Officer", new Date("1985-09-01"), 40, "Male", new Date("2025-01-20"), "P678901F", new Date("2031-08-31"), "https://drive.google.com/file/d/principal_pass_4", "V4567890", new Date("2030-01-20"), "D-00401", new Date("2030-01-20"), new Date("2031-01-19"), "No", "", "No", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
      // Row 6 (Chen's staff)
      ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "DAVID, Leni S.", "Female", new Date("1990-04-16"), 35, "Yes", new Date("2025-02-01"), "S876543G", new Date("2029-04-15"), "https://drive.google.com/file/d/staff_pass_2", "V4567891", new Date("2028-02-01"), "D-00402", new Date("2028-02-01"), new Date("2031-01-19"), "Yes", "https://drive.google.com/file/d/staff_pwd_1", "Pedro David - 09187654321", "No", "", ""],
      // Row 7 (Fatima Al-Fassi)
      ["Riyadh", "AL-FASSI, Fatima Z.", "Foreign Service Officer II", "Consul", new Date("1988-07-07"), 37, "Female", new Date("2024-11-01"), "P789012G", new Date("2029-07-06"), "https://drive.google.com/file/d/principal_pass_5", "V5678901", new Date("2029-11-01"), "D-00501", new Date("2029-11-01"), new Date("2030-10-31"), "Yes", "https://drive.google.com/file/d/solo_parent_1", "No", "", "", "AL-FASSI, Omar Y.", "Child", "Male", new Date("2019-12-01"), 5, new Date("2037-12-01"), "Yes", new Date("2024-11-01"), "Yes", "P890123H", new Date("2029-11-30"), "https://drive.google.com/file/d/dep_pass_4", "V5678902", new Date("2029-11-01"), "D-00502", new Date("2029-11-01"), new Date("2030-10-31"), "No", "", "", "No", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
    ];

    // Write the data to the sheet starting at row 3, column 1
    sheet.getRange(3, 1, 7, 64).setValues(sampleData);
    
    // Set date and number formats
    sheet.getRange('E3:E9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('F3:F9').setNumberFormat('0');
    sheet.getRange('H3:H9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('J3:J9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('M3:M9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('O3:O9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('P3:P9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('Y3:Y9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('Z3:Z9').setNumberFormat('0');
    sheet.getRange('AA3:AA9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AC3:AC9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AF3:AF9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AI3:AI9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AK3:AK9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AL3:AL9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AU3:AU9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AV3:AV9').setNumberFormat('0');
    sheet.getRange('AX3:AX9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('AZ3:AZ9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('BC3:BC9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('BE3:BE9').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('BF3:BF9').setNumberFormat('yyyy-mm-dd');

    SpreadsheetApp.getUi().alert('✅ Sample data (7 rows) has been successfully inserted into "Personnel Tracking" sheet.');
    
  } catch (error) {
    Logger.log('Error inserting sample data: ' + error);
    SpreadsheetApp.getUi().alert('Error inserting sample data: ' + error.message);
  }
}
