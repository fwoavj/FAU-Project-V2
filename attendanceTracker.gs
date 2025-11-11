/**
 * This file is for functions related to the Attendance Tracker and Update Form.
 * Note: `COLUMN_CONFIG` is loaded from the CONFIG.gs file.
 */

// ============================================================================
// UPDATE FORM FUNCTIONS
// ============================================================================

/**
 * Gets personnel data for the "Update Form"
 */
function getPersonnelDataForUpdate(principalName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) {
      throw new Error('Personnel Tracking sheet not found');
    }

    const principalRow = findPrincipalRow(sheet, principalName); // findPrincipalRow is in codeTester.gs
    if (!principalRow) {
      throw new Error('Principal not found: ' + principalName);
    }
    
    // Use COLUMN_CONFIG for reliable column numbers
    const pCols = COLUMN_CONFIG; 
    const postStation = sheet.getRange(principalRow, pCols.COL_PRINCIPAL_POST).getValue();
    const originalDeparture = sheet.getRange(principalRow, pCols.COL_PRINCIPAL_DEPARTURE).getValue();
    const extended = sheet.getRange(principalRow, pCols.COL_PRINCIPAL_EXTENDED).getValue();
    const newDeparture = sheet.getRange(principalRow, pCols.COL_PRINCIPAL_CURRENT_DEPARTURE).getValue();
    const extensionDetails = sheet.getRange(principalRow, pCols.COL_PRINCIPAL_EXTENSION_DETAILS).getValue();
    
    return {
      postStation: postStation || '',
      originalDeparture: formatDateForClient(originalDeparture),
      extended: extended || '',
      newDeparture: formatDateForClient(newDeparture),
      extensionDetails: extensionDetails || ''
    };
  } catch (error) {
    Logger.log('Error in getPersonnelDataForUpdate: ' + error);
    throw new Error('Failed to get personnel data: ' + error.message);
  }
}

/**
 * Updates personnel data from the "Update Form"
 */
function updatePersonnelData(principalName, updateData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    if (!sheet) {
      throw new Error('Personnel Tracking sheet not found');
    }

    const principalRow = findPrincipalRow(sheet, principalName); // From codeTester.gs
    if (!principalRow) {
      throw new Error('Principal not found');
    }
    
    clearDataCache(); // Clear cache since we are writing data

    const pCols = COLUMN_CONFIG;
    const updates = [];

    // 1. Update Principal fields
    if (updateData.extended) {
      updates.push({ row: principalRow, col: pCols.COL_PRINCIPAL_EXTENDED, value: updateData.extended });
    }
    if (updateData.newDeparture) {
      updates.push({ row: principalRow, col: pCols.COL_PRINCIPAL_CURRENT_DEPARTURE, value: new Date(updateData.newDeparture) });
    }
    if (updateData.extensionDetails !== undefined) {
      let newDetails = updateData.extensionDetails;
      // Note: appendExtension logic was removed, uncomment if needed
      // if (updateData.appendExtension) { 
      //   const existingDetails = sheet.getRange(principalRow, pCols.COL_PRINCIPAL_EXTENSION_DETAILS).getValue();
      //   newDetails = existingDetails + '\n' + updateData.extensionDetails;
      // }
      updates.push({ row: principalRow, col: pCols.COL_PRINCIPAL_EXTENSION_DETAILS, value: newDetails });
    }

    // 2. Find all associated rows for this principal
    const allRows = findPrincipalRows(sheet, principalName); // From codeTester.gs
    
    // 3. Update all associated Dependent and Staff rows
    allRows.forEach(rowNum => {
      const rowData = sheet.getRange(rowNum, 1, 1, 64).getValues()[0];
      const depName = rowData[SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES.FULL_NAME];
      const staffName = rowData[SYSTEM_CONFIG.COLUMNS.STAFF._INDICES.FULL_NAME];

      // Update Dependent if one exists on this row
      if (depName && depName.toString().trim() !== '') {
        updates.push({ row: rowNum, col: pCols.COL_DEPENDENT_EXTENDED, value: updateData.extended });
        updates.push({ row: rowNum, col: pCols.COL_DEPENDENT_CURRENT_DEPARTURE, value: new Date(updateData.newDeparture) });
        updates.push({ row: rowNum, col: pCols.COL_DEPENDENT_EXTENSION_DETAILS, value: updateData.extensionDetails });
      }
      
      // Update Staff if one exists on this row
      if (staffName && staffName.toString().trim() !== '') {
        updates.push({ row: rowNum, col: pCols.COL_STAFF_EXTENDED, value: updateData.extended });
        updates.push({ row: rowNum, col: pCols.COL_STAFF_CURRENT_DEPARTURE, value: new Date(updateData.newDeparture) });
        updates.push({ row: rowNum, col: pCols.COL_STAFF_EXTENSION_DETAILS, value: updateData.extensionDetails });
      }
    });

    // 4. Execute batch update
    batchUpdateCells(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING, updates); // From DATA_ACCESS.gs

    // 5. Handle Attendance Extension
    if (updateData.extendAttendance === true) {
      const attendanceResult = extendAttendanceForFamily(principalName);
      return {
        success: true,
        message: `✅ Extension recorded successfully!\n\n${attendanceResult.message}`,
        attendanceExtended: attendanceResult.extended,
        attendanceSkipped: attendanceResult.skipped
      };
    }
    
    return {
      success: true,
      message: '✅ Extension recorded successfully!'
    };
    
  } catch (error) {
    console.error('Error updating personnel:', error);
    return {
      success: false,
      message: `Failed to update: ${error.message}`
    };
  }
}

/**
 * Gets a list of principals who have NOT departed
 */
function getExistingPrincipalsFromTracking() {
  try {
    const allPersonnel = getAllPersonnelDataCached(); // From DATA_ACCESS.gs
    if (allPersonnel.length === 0) {
      return [];
    }
    
    // 1. Get all unique principals from tracking sheet
    const allPrincipalsMap = new Map();
    allPersonnel.forEach(p => {
      if (p.principal && p.principal.fullName && !allPrincipalsMap.has(p.principal.fullName)) {
        allPrincipalsMap.set(p.principal.fullName, {
          fullName: p.principal.fullName,
          post: p.principal.postStation || ''
        });
      }
    });

    // 2. Get the list of departed principals
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const departureSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PRINCIPAL_DEPARTURE_LOG);
    const departedNames = new Set();
    
    if (departureSheet) {
      const lastRow = departureSheet.getLastRow();
      if (lastRow >= 2) {
        // Get data from Column B (Principal Name), starting from row 2
        const departedData = departureSheet.getRange(2, 2, lastRow - 1, 1).getValues();
        departedData.forEach(row => {
          if (row[0] && row[0].toString().trim() !== '') {
            departedNames.add(row[0].toString().trim());
          }
        });
      }
    } else {
      Logger.log('WARNING: "Principal Departure Log" sheet not found. Cannot filter departed principals.');
    }

    // 3. Filter the 'allPrincipals' list
    const activePrincipals = [];
    allPrincipalsMap.forEach((principal, fullName) => {
      if (!departedNames.has(fullName)) {
        activePrincipals.push(principal);
      }
    });

    Logger.log(`Filtered principals: Total ${allPrincipalsMap.size}, Departed ${departedNames.size}, Active ${activePrincipals.length}`);
    return activePrincipals;
  } catch (error) {
    console.error('Error in getExistingPrincipalsFromTracking:', error);
    return []; // Return empty on error
  }
}

/**
 * Extend status of residency for principal's active family members
 */
function extendAttendanceForFamily(principalName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
    const archiveSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
    if (!attendanceSheet) {
      return {
        success: false,
        message: 'Status of Residency sheet not found',
        extended: 0,
        skipped: 0
      };
    }

    // Get archived pairs
    const archivedPairs = new Set();
    if (archiveSheet && archiveSheet.getLastRow() >= 2) {
      const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 1).getValues();
      archiveData.forEach(row => {
        if (row[0]) archivedPairs.add(row[0].toString().trim());
      });
    }
    
    // Find all family members from the cached data
    const allPersonnel = getAllPersonnelDataCached(); // from DATA_ACCESS.gs
    const principalData = allPersonnel.filter(p => p.principal.fullName === principalName);

    const activeFamilyMembers = [];
    const archivedFamilyMembers = [];
    const seenMembers = new Set(); // To handle duplicates if any

    principalData.forEach(p => {
      // Check dependents
      p.dependents.forEach(dep => {
        const memberName = (dep.fullName || '').toString().trim();
        if (!memberName || seenMembers.has(memberName)) return;
        seenMembers.add(memberName);
        
        const pairName = `${principalName} - ${memberName}`;
        if (archivedPairs.has(pairName)) {
          archivedFamilyMembers.push({ name: memberName, type: 'Dependent', pairName: pairName });
        } else {
          activeFamilyMembers.push({ name: memberName, type: 'Dependent', pairName: pairName });
        }
      });
      
      // Check staff
      p.staff.forEach(staff => {
        const memberName = (staff.fullName || '').toString().trim();
        if (!memberName || seenMembers.has(memberName)) return;
        seenMembers.add(memberName);

        const pairName = `${principalName} - ${memberName}`;
        if (archivedPairs.has(pairName)) {
          archivedFamilyMembers.push({ name: memberName, type: 'Staff', pairName: pairName });
        } else {
          activeFamilyMembers.push({ name: memberName, type: 'Staff', pairName: pairName });
        }
      });
    });

    // Record attendance for active members only
    const currentDate = new Date();
    let recordedCount = 0;
    
    activeFamilyMembers.forEach(member => {
      try {
        const state = getPairState(member.pairName);
        const quarterValue = state.currentQuarter || 'Q1';
        
        const nextRow = attendanceSheet.getLastRow() + 1;
        attendanceSheet.getRange(nextRow, 1, 1, 5).setValues([[
          member.pairName,
          quarterValue,
          SYSTEM_CONFIG.ATTENDANCE.STATUS.AT_POST,
          'Principal extension - auto-recorded',
          currentDate
        ]]);
        
        attendanceSheet.getRange(nextRow, 5).setNumberFormat('yyyy-mm-dd hh:mm:ss');
        recordedCount++;
        
        Logger.log(`✅ Recorded status of residency for: ${member.pairName}`);
      } catch (error) {
        Logger.log(`⚠️ Error recording status of residency for ${member.pairName}: ${error}`);
      }
    });

    // Build result message
    let message = '';
    if (recordedCount > 0) {
      message += `✅ Status of Residency extended for ${recordedCount} active family member(s)`;
    }
    
    if (archivedFamilyMembers.length > 0) {
      message += `\n⚠️ Skipped ${archivedFamilyMembers.length} archived member(s)`;
    }
    
    if (recordedCount === 0 && archivedFamilyMembers.length === 0) {
      message = 'ℹ️ No active dependents/staff found to extend';
    }
    
    Logger.log(`Extension summary: ${recordedCount} extended, ${archivedFamilyMembers.length} skipped`);
    return {
      success: true,
      message: message,
      extended: recordedCount,
      skipped: archivedFamilyMembers.length,
      activeMembers: activeFamilyMembers,
      archivedMembers: archivedFamilyMembers
    };
  } catch (error) {
    console.error('Error extending family status of residency:', error);
    return {
      success: false,
      message: `Failed to extend status of residency: ${error.message}`,
      extended: 0,
      skipped: 0
    };
  }
}

/**
 * Format date for HTML input type="date"
 */
function formatDateForClient(dateValue) {
  if (!dateValue) return '';
  try {
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return '';
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
  } catch (error) {
    return '';
  }
}

/**
 * Withdraw entire family from the Update Form
 */
function withdrawEntireFamily(principalName, departureDate, reason) {
  try {
    console.log('🚪 Withdrawing entire family for:', principalName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    const principalDepartureLog = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PRINCIPAL_DEPARTURE_LOG);
    const attendanceLog = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
    const archivedWithdrawals = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
    
    if (!trackingSheet || !principalDepartureLog || !attendanceLog || !archivedWithdrawals) {
      throw new Error('Required sheets not found (Tracking, Departure Log, Status of Residency, or Archive)');
    }

    const principalRow = findPrincipalRow(trackingSheet, principalName); // From codeTester.gs
    if (!principalRow) {
      return {
        success: false,
        message: `Principal "${principalName}" not found`
      };
    }
    
    clearDataCache(); // Clear cache since we are writing data

    const today = new Date();
    const departureDateObj = new Date(departureDate);
    const post = trackingSheet.getRange(principalRow, COLUMN_CONFIG.COL_PRINCIPAL_POST).getValue();
    let withdrawnCount = 0;

    // Log principal to Principal Departure Log
    const principalLogLastRow = principalDepartureLog.getLastRow();
    principalDepartureLog.getRange(principalLogLastRow + 1, 1, 1, 6).setValues([[
      today,
      principalName,
      post,
      reason,
      departureDateObj,
      Session.getActiveUser().getEmail()
    ]]);
    withdrawnCount++;

    // Get dependents and staff from all rows associated with the principal
    const allPersonnel = getAllPersonnelData(); // Get fresh data
    const principalData = allPersonnel.filter(p => p.principal.fullName === principalName);

    const familyMembers = [];
    const seenMembers = new Set();
    
    principalData.forEach(p => {
      p.dependents.forEach(d => {
        if (d.fullName && !seenMembers.has(d.fullName)) {
          familyMembers.push({ name: d.fullName, type: 'Dependent' });
          seenMembers.add(d.fullName);
        }
      });
      p.staff.forEach(s => {
        if (s.fullName && !seenMembers.has(s.fullName)) {
          familyMembers.push({ name: s.fullName, type: 'Staff' });
          seenMembers.add(s.fullName);
        }
      });
    });

    // Process all family members
    familyMembers.forEach(member => {
      const pairName = principalName + ' - ' + member.name;
      const state = getPairState(pairName);
      
      // Only archive if not already archived
      if (state.lastRecord && state.lastRecord.status === SYSTEM_CONFIG.ATTENDANCE.STATUS.NOT_AT_POST) {
        Logger.log(`Skipping already withdrawn member: ${pairName}`);
        return;
      }
      
      const quarterValue = state.currentQuarter || 'Q1';
      
      // Add to Attendance Log
      const attendanceLogLastRow = attendanceLog.getLastRow();
      attendanceLog.getRange(attendanceLogLastRow + 1, 1, 1, 5).setValues([[
        pairName,
        quarterValue,
        SYSTEM_CONFIG.ATTENDANCE.STATUS.NOT_AT_POST,
        reason,
        departureDateObj
      ]]);
      
      // Add to Archived Withdrawals
      const archiveLastRow = archivedWithdrawals.getLastRow();
      archivedWithdrawals.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
        pairName,
        quarterValue,
        reason,
        departureDateObj,
        today
      ]]);
      withdrawnCount++;
    });

    SpreadsheetApp.flush();
    
    return {
      success: true,
      message: `✅ Successfully withdrawn ${withdrawnCount} person(s) (including principal)`,
      withdrawnCount: withdrawnCount
    };
  } catch (error) {
    console.error('❌ Withdrawal error:', error);
    return {
      success: false,
      message: 'Error withdrawing family: ' + error.message
    };
  }
}

/**
 * Get status of family members for Update Form
 * This function is now CORRECTED
 */
function getFamilyMembersStatus(principalName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
  
  const withdrawnInfo = {};
  
  // Get archived withdrawals
  if (archiveSheet && archiveSheet.getLastRow() >= 2) {
    const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 5).getValues();
    archiveData.forEach(row => {
      const fullName = (row[0] || '').toString().trim();
      const withdrawalDate = row[3]; // Column D: Withdrawal Date
      if (fullName.startsWith(principalName + ' - ')) {
        const memberName = fullName.substring(principalName.length + 3);
        withdrawnInfo[memberName] = { date: withdrawalDate ? new Date(withdrawalDate) : null };
      }
    });
  }

  // Get "Not At Post" from attendance log
  if (logSheet && logSheet.getLastRow() >= 2) {
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 5).getValues();
    const attendanceLog = {};
    logData.forEach(row => {
      const fullName = (row[0] || '').toString().trim();
      const attendance = (row[2] || '').toString().trim();
      const date = row[4];
      if (fullName.startsWith(principalName + ' - ') && date) {
        const logDate = new Date(date);
        if (!attendanceLog[fullName] || logDate > attendanceLog[fullName].date) {
          attendanceLog[fullName] = { attendance, date: logDate };
        }
      }
    });
    
    Object.keys(attendanceLog).forEach(fullName => {
      if (attendanceLog[fullName].attendance === SYSTEM_CONFIG.ATTENDANCE.STATUS.NOT_AT_POST) {
        const memberName = fullName.substring(principalName.length + 3);
        if (!withdrawnInfo[memberName]) {
          withdrawnInfo[memberName] = { date: attendanceLog[fullName].date };
        }
      }
    });
  }

  const active = [];
  const archived = [];
  
  // Get family members from cached data
  const allPersonnel = getAllPersonnelDataCached(); // from DATA_ACCESS.gs
  const principalData = allPersonnel.filter(p => p.principal.fullName === principalName);

  const familyMembers = new Map();
  principalData.forEach(p => {
    p.dependents.forEach(d => {
      if (d.fullName && !familyMembers.has(d.fullName)) familyMembers.set(d.fullName, 'Dependent');
    });
    p.staff.forEach(s => {
      if (s.fullName && !familyMembers.has(s.fullName)) familyMembers.set(s.fullName, 'Staff');
    });
  });

  familyMembers.forEach((type, name) => {
    if (withdrawnInfo[name]) {
      archived.push({
        name,
        type,
        withdrawalDate: withdrawnInfo[name].date ? Utilities.formatDate(withdrawnInfo[name].date, Session.getScriptTimeZone(), 'MMM dd, yyyy') : 'Unknown'
      });
    } else {
      active.push({ name, type });
    }
  });
  
  return { active, archived };
}

/**
 * Get family members for withdrawal confirmation
 */
function getFamilyMembersForWithdrawal(principalName) {
  // This function is nearly identical to getFamilyMembersStatus,
  // just returns data formatted for the withdrawal preview.
  try {
    const { active, archived } = getFamilyMembersStatus(principalName);
    return {
      principalName: principalName,
      active: active,
      withdrawn: archived
    };
  } catch (error) {
    console.error('Error getting family for withdrawal:', error);
    return { principalName: principalName, active: [], withdrawn: [] };
  }
}

// ============================================================================
// ATTENDANCE TRACKER LOGIC
// ============================================================================

/**
 * Saves an attendance record.
 * This function is now CORRECTED.
 */
function saveAttendance(pairName, status, remarks) {
  try {
    Logger.log(`Saving status of residency for ${pairName} with status ${status}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
    const archiveSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
    
    if (!logSheet) {
      return { success: false, message: SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG + ' not found' };
    }
    
    const state = getPairState(pairName);
    const quarter = state.currentQuarter;
    const today = new Date();
    
    // NOT At Post - Save to BOTH Archive AND Attendance Log
    if (status === SYSTEM_CONFIG.ATTENDANCE.STATUS.NOT_AT_POST) {
      if (!archiveSheet) {
        return { success: false, message: SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS + ' not found' };
      }
      
      // Save to Archived Withdrawals
      const archiveLastRow = archiveSheet.getLastRow();
      archiveSheet.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
        pairName, quarter, remarks || '', today, today
      ]]);
      
      // ALSO save to Attendance Log
      const logLastRow = logSheet.getLastRow();
      logSheet.getRange(logLastRow + 1, 1, 1, 5).setValues([[
        pairName, quarter, status, remarks || '', today
      ]]);
      
      // Clear cache because this is a data change
      clearDataCache();
      Logger.log('Saved to both Archive and Status of Residency Log');
      return { success: true, message: pairName + ' archived as withdrawn.' };
    }
    
    // At Post - Save to Attendance Log
    const logLastRow = logSheet.getLastRow();
    logSheet.getRange(logLastRow + 1, 1, 1, 5).setValues([[
      pairName, quarter, status, remarks || '', today
    ]]);
    
    Logger.log('Recorded to status of residency as ' + quarter);
    return {
      success: true,
      message: `Recorded ${quarter}. Next window opens in ${SYSTEM_CONFIG.ATTENDANCE.WINDOW_START_DAY} days.`
    };
    
  } catch (error) {
    Logger.log('Error in saveAttendance: ' + error.toString());
    return { success: false, message: 'Error: ' + error.message };
  }
}


/**
 * Gets the last state (quarter, date) for a given pair.
 * This function is CORRECTED.
 */
function getPairState(pairName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
  let lastRecord = null;

  if (logSheet && logSheet.getLastRow() >= 2) {
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 5).getValues();
    // Loop backwards to find the most recent entry
    for (let i = logData.length - 1; i >= 0; i--) {
      const name = (logData[i][0] || '').toString().trim();
      if (name === pairName) {
        lastRecord = {
          pairName: name,
          quarter: (logData[i][1] || 'Q0').toString().trim(),
          status: (logData[i][2] || '').toString().trim(),
          date: new Date(logData[i][4]) // Column E is the date
        };
        break; // Found the last record
      }
    }
  }

  // This pair has no record yet
  if (!lastRecord) {
    return { lastRecord: null, currentQuarter: 'Q1' };
  }
  
  // This pair has a record
  const lastQNum = parseInt(lastRecord.quarter.replace('Q', '').replace('+', '')) || 0;
  const nextQNum = lastQNum + 1;
  const nextQuarter = nextQNum > SYSTEM_CONFIG.ATTENDANCE.MAX_QUARTERS ?
                      (`Q${SYSTEM_CONFIG.ATTENDANCE.MAX_QUARTERS}+`) :
                      (`Q${nextQNum}`);
                      
  return { lastRecord: lastRecord, currentQuarter: nextQuarter };
}

/**
 * Calculates the window status based on the last status of residency record.
 */
function calculateWindowHelper(lastRecord, today) {
  const config = SYSTEM_CONFIG.ATTENDANCE;
  
  // Case 1: No record, this is a new pair
  if (!lastRecord || !lastRecord.date) {
    return {
      isNewPair: true,
      quarter: 'Q1',
      isOpen: true,
      status: 'Ready',
      message: 'Ready to Start'
    };
  }
  
  // Case 2: This pair has an existing record
  const lastDate = lastRecord.date;
  const lastQuarter = lastRecord.quarter;
  const daysSinceLastAttendance = Math.floor((today - lastDate) / (1000 * 60 * 60 * 24));
  
  const lastQNum = parseInt(lastQuarter.replace('Q', '').replace('+', '')) || 1;
  const nextQNum = lastQNum + 1;
  const nextQuarter = nextQNum > config.MAX_QUARTERS ? (`Q${config.MAX_QUARTERS}+`) : (`Q${nextQNum}`);
  
  const windowInfo = {
    isNewPair: false,
    quarter: nextQuarter,
    daysSinceLastAttendance: daysSinceLastAttendance,
    lastAttendanceDate: lastDate.toISOString(), // Send as string
    windowStartDay: config.WINDOW_START_DAY,
    windowEndDay: config.WINDOW_END_DAY,
    gracePeriodDays: config.GRACE_PERIOD_DAYS
  };
  
  // Check window status
  if (daysSinceLastAttendance >= config.WINDOW_START_DAY && daysSinceLastAttendance <= config.WINDOW_END_DAY) {
    const daysRemaining = config.WINDOW_END_DAY - daysSinceLastAttendance;
    windowInfo.isOpen = true;
    windowInfo.status = 'Open';
    windowInfo.daysRemaining = daysRemaining;
    windowInfo.message = `Window Open (${daysRemaining} ${daysRemaining === 1 ? 'day' : 'days'} left)`;
  } else if (daysSinceLastAttendance < config.WINDOW_START_DAY) {
    const daysUntilWindow = config.WINDOW_START_DAY - daysSinceLastAttendance;
    windowInfo.isOpen = false;
    windowInfo.status = 'Upcoming';
    windowInfo.message = `${daysUntilWindow} ${daysUntilWindow === 1 ? 'day' : 'days'} until window`;
  } else {
    // Window was missed
    const daysOverdue = daysSinceLastAttendance - config.WINDOW_END_DAY;
    windowInfo.isOpen = false;
    windowInfo.status = 'Missed';
    windowInfo.message = `Missed by ${daysOverdue} ${daysOverdue === 1 ? 'day' : 'days'}`;
  }
  
  return windowInfo;
}

/**
 * Loads all active dependent/staff pairs for the Attendance Tracker UI.
 * This function is now CORRECTED.
 */
function loadAttendanceData() {
  Logger.log('=== loadAttendanceData START ===');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archiveSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
    const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
    
    // 1. Get all personnel data
    const allPersonnel = getAllPersonnelDataCached(); // from DATA_ACCESS.gs
    if (allPersonnel.length === 0) {
      Logger.log('No personnel data found');
      return [];
    }

    // 2. Load Archived Withdrawals
    const archivedPairs = new Set();
    if (archiveSheet && archiveSheet.getLastRow() >= 2) {
      const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 1).getValues();
      archiveData.forEach(row => {
        const name = (row[0] || '').toString().trim();
        if (name) archivedPairs.add(name);
      });
      Logger.log('Loaded ' + archivedPairs.size + ' archived pairs');
    }
    
    // 3. Load Attendance Log
    const attendanceLog = {};
    if (logSheet && logSheet.getLastRow() >= 2) {
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 5).getValues();
      logData.forEach(row => {
        const name = (row[0] || '').toString().trim();
        const quarter = (row[1] || '').toString().trim();
        const attendance = (row[2] || '').toString().trim();
        const date = row[4]; // Column E
        if (name && date) {
          const logDate = new Date(date);
          if (!attendanceLog[name] || logDate > attendanceLog[name].date) {
            attendanceLog[name] = { date: logDate, quarter, attendance };
          }
        }
      });
      Logger.log('Loaded attendance log for ' + Object.keys(attendanceLog).length + ' pairs');
    }
    
    // 4. Process all personnel
    const attendancePairs = [];
    const today = new Date();
    
    allPersonnel.forEach(person => {
      const principalName = person.principal.fullName;
      if (!principalName) return;
      
      const post = person.principal.postStation;

      // Process Dependents
      person.dependents.forEach(dep => {
        const memberName = (dep.fullName || '').toString().trim();
        if (!memberName) return;
        
        const pairName = `${principalName} - ${memberName}`;
        
        // CHECK 1: Skip if in Archived Withdrawals
        if (archivedPairs.has(pairName)) return;
        
        // CHECK 2: Skip if last attendance was "Not At Post"
        const lastRecord = attendanceLog[pairName];
        if (lastRecord && lastRecord.attendance === SYSTEM_CONFIG.ATTENDANCE.STATUS.NOT_AT_POST) return;
        
        // This person is ACTIVE - add to attendance tracker
        const windowInfo = calculateWindowHelper(lastRecord, today);
        attendancePairs.push({
          name: pairName,
          principalName,
          memberName,
          memberType: 'Dependent',
          post,
          atPost: dep.atPost,
          quarter: windowInfo.quarter,
          status: 'Active',
          isNewPair: windowInfo.isNewPair,
          windowInfo,
          rowNumber: person.rowNumber
        });
      });

      // Process Staff
      person.staff.forEach(staff => {
        const memberName = (staff.fullName || '').toString().trim();
        if (!memberName) return;
        
        const pairName = `${principalName} - ${memberName}`;
        
        // CHECK 1: Skip if in Archived Withdrawals
        if (archivedPairs.has(pairName)) return;
        
        // CHECK 2: Skip if last attendance was "Not At Post"
        const lastRecord = attendanceLog[pairName];
        if (lastRecord && lastRecord.attendance === SYSTEM_CONFIG.ATTENDANCE.STATUS.NOT_AT_POST) return;
        
        // This person is ACTIVE - add to attendance tracker
        const windowInfo = calculateWindowHelper(lastRecord, today);
        attendancePairs.push({
          name: pairName,
          principalName,
          memberName,
          memberType: 'Staff',
          post,
          atPost: staff.atPost,
          quarter: windowInfo.quarter,
          status: 'Active',
          isNewPair: windowInfo.isNewPair,
          windowInfo,
          rowNumber: person.rowNumber
        });
      });
    });

    Logger.log('TOTAL ACTIVE PAIRS: ' + attendancePairs.length);
    return attendancePairs;
    
  } catch (error) {
    Logger.log('ERROR loading attendance data: ' + error);
    Logger.log('Stack: ' + error.stack);
    return [];
  }
}

/**
 * Auto-archives pairs that have missed their window and grace period.
 */
function autoArchiveMissedWindows() {
  try {
    Logger.log('=== AUTO-ARCHIVE MISSED WINDOWS START ===');
    const config = SYSTEM_CONFIG.ATTENDANCE;
    if (!config.AUTO_ARCHIVE_MISSED) {
      return { success: true, message: 'Auto-archive disabled', archived: 0 };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ATTENDANCE_LOG);
    const archiveSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
    
    if (!logSheet || !archiveSheet) {
      return { success: false, message: 'Required sheets (Log or Archive) not found' };
    }
    
    const attendanceData = loadAttendanceData(); // Gets all active pairs
    const today = new Date();
    let archivedCount = 0;
    const cutoffDays = config.WINDOW_END_DAY + config.GRACE_PERIOD_DAYS;
    
    attendanceData.forEach(function(pair) {
      // Skip new pairs or pairs that are already withdrawn (should be filtered by loadData, but double check)
      if (pair.isNewPair || pair.status === config.STATUS.WITHDRAWN) {
        return;
      }
      
      // Check if the window was missed and grace period is over
      if (pair.windowInfo && pair.windowInfo.daysSinceLastAttendance > cutoffDays) {
        Logger.log('Auto-archiving: ' + pair.name);
        
        // Calculate when window (and grace period) actually closed
        const lastAttendanceDate = new Date(pair.windowInfo.lastAttendanceDate);
        const withdrawalDate = new Date(lastAttendanceDate);
        withdrawalDate.setDate(withdrawalDate.getDate() + cutoffDays + 1);
        
        // Archive to sheet
        const archiveLastRow = archiveSheet.getLastRow();
        archiveSheet.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
          pair.name,
          pair.quarter,
          config.AUTO_ARCHIVE_REASON,
          withdrawalDate,
          today
        ]]);
        
        // Log "Not At Post"
        const logLastRow = logSheet.getLastRow();
        logSheet.getRange(logLastRow + 1, 1, 1, 5).setValues([[
          pair.name,
          pair.quarter,
          config.STATUS.NOT_AT_POST,
          config.AUTO_ARCHIVE_REASON,
          withdrawalDate // Use the calculated withdrawal date
        ]]);
        
        archivedCount++;
      }
    });
    
    if (archivedCount > 0) {
      clearDataCache(); // Clear cache if we made changes
    }
    
    Logger.log('Auto-archived ' + archivedCount + ' pairs');
    return { success: true, message: 'Auto-archived ' + archivedCount + ' pair(s)', archived: archivedCount };
  } catch (error) {
    Logger.log('ERROR in autoArchiveMissedWindows: ' + error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

/**
 * Loads all data from the Archived Withdrawals sheet for the archive viewer.
 */
function loadArchivedData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archiveSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.ARCHIVED_WITHDRAWALS);
    if (!archiveSheet || archiveSheet.getLastRow() < 2) {
      return [];
    }

    const data = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 5).getValues();
    const archivedData = [];

    data.forEach(row => {
      archivedData.push({
        name: (row[0] || '').toString().trim(),
        quarter: (row[1] || '').toString().trim(),
        withdrawalReason: (row[2] || '').toString().trim(),
        withdrawalDate: formatDateForClient(row[3]), // Col D
        archivedDate: formatDateForClient(row[4])  // Col E
      });
    });
    
    // Sort by most recent withdrawal date
    archivedData.sort((a, b) => new Date(b.withdrawalDate) - new Date(a.withdrawalDate));
    
    return archivedData;
  } catch (error) {
    Logger.log('Error loading archived data: ' + error);
    return [];
  }
}
