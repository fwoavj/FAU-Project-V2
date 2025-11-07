  function submitPersonnelData(formData) {
    try {
      // ADD VALIDATION
      const validation = validateFormData(formData);
      
      if (!validation.isValid) {
        const errorMessages = validation.errors.map(e => `${e.field}: ${e.message}`).join('\n');
        return {
          success: false,
          message: '❌ Validation failed:\n\n' + errorMessages
        };
      }
      
      // Show warnings but continue
      if (validation.warnings.length > 0) {
        const warningMessages = validation.warnings.map(w => `${w.field}: ${w.message}`).join('\n');
        Logger.log('⚠️ Warnings:\n' + warningMessages);
      }
      
      // CHECK FOR DUPLICATES
      const dupCheck = checkDuplicatePrincipal(formData.principal.fullName, formData.principal.dateOfBirth);
      if (dupCheck.isDuplicate) {
        return {
          success: false,
          message: '❌ Duplicate detected:\n\n' + dupCheck.message
        };
      }
      
      // PROCEED WITH ORIGINAL SUBMISSION CODE
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
      
      if (!sheet) {
        throw new Error('Sheet "Personnel Tracking" not found');
      }

      // Record principal
      const principalResult = recordPrincipal(sheet, formData.principal);
      
      if (!principalResult.success) {
        return principalResult;
      }

      const principalRow = principalResult.row;
      const principalName = formData.principal.fullName;
      
      // ADD AUDIT LOG
      auditPersonnelCreated('Principal', principalName, principalName);

      // Record dependents
      if (formData.dependents && formData.dependents.length > 0) {
        for (let i = 0; i < formData.dependents.length; i++) {
          const dependent = formData.dependents[i];
          const depResult = recordDependent(sheet, principalName, dependent, principalRow);
          
          if (depResult.success) {
            // ADD AUDIT LOG
            auditPersonnelCreated('Dependent', dependent.fullName, principalName);
          }
        }
      }

      // Record private staff
      if (formData.privateStaff && formData.privateStaff.length > 0) {
        for (let i = 0; i < formData.privateStaff.length; i++) {
          const staff = formData.privateStaff[i];
          const staffResult = recordPrivateStaff(sheet, principalName, staff, principalRow);
          
          if (staffResult.success) {
            // ADD AUDIT LOG
            auditPersonnelCreated('Staff', staff.fullName, principalName);
          }
        }
      }

      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: '✅ Personnel data submitted successfully!',
        principalRow: principalRow
      };

    } catch (error) {
      console.error('❌ Submission error:', error);
      // ADD ERROR AUDIT
      auditError('SUBMIT_PERSONNEL', 'System', 'N/A', error.message);
      
      return {
        success: false,
        message: 'Error: ' + error.message
      };
    }
  }

    // ============================================================================
    // PRINCIPAL RECORDING
    // ============================================================================

  function recordPrincipal(sheet, principal) {
    try {
      console.log('👤 Recording principal:', principal.fullName);
      
      // Check for duplicate
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][COLUMN_CONFIG.COL_PRINCIPAL_NAME - 1] === principal.fullName) {
          return {
            success: false,
            message: `Principal "${principal.fullName}" already exists in row ${i + 1}`
          };
        }
      }

      // Find next empty row
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;

      // Write principal data
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_PRINCIPAL_NAME).setValue(principal.fullName);
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_RANK).setValue(principal.rank || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_POSITION).setValue(principal.position || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_POST).setValue(principal.post || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_ARRIVAL_DATE).setValue(principal.arrivalDate || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_DEPARTURE_DATE).setValue(principal.departureDate || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_EXTENDED).setValue(principal.extended || 'No');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_CURRENT_DEPARTURE).setValue(principal.newDeparture || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_EXTENSION_DETAILS).setValue(principal.extensionDetails || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_VISA_NUMBER).setValue(principal.visaNumber || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_VISA_EXP).setValue(principal.visaExp || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_DIPLO_ID).setValue(principal.dipId || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_DIPLO_EXP).setValue(principal.dipIdExp || '');
      sheet.getRange(nextRow, COLUMN_CONFIG.COL_PWD_STATUS).setValue(principal.pwdStatus || 'No');

      console.log(`✅ Principal recorded in row ${nextRow}`);
      
      return {
        success: true,
        row: nextRow
      };

    } catch (error) {
      console.error('❌ Error recording principal:', error);
      return {
        success: false,
        message: 'Error recording principal: ' + error.message
      };
    }
    }

    // ============================================================================
    // DEPENDENT RECORDING
    // ============================================================================

  /**
   * ============================================================================
   * DEPENDENT RECORDING - FIXED
   * ============================================================================
   * This REPLACES the old recordDependent function.
   * It writes to the principal's *own row* instead of looping 50 rows down.
   */
  function recordDependent(sheet, principalName, dependent, principalRow) {
    try {
      console.log(`👨‍👩‍👧 Recording dependent: ${dependent.fullName} for ${principalName}`);
      
      // The principalRow IS the target row.
      const targetRow = principalRow; 
      
      if (!targetRow) {
        return {
          success: false,
          message: `Principal "${principalName}" not found or row not provided`
        };
      }
      
      // Check if a dependent already exists on this row
      const existingName = sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_NAME).getValue();
      if (existingName && existingName.toString().trim() !== '') {
        return {
          success: false,
          message: `A dependent (${existingName}) already exists for ${principalName} in row ${targetRow}. Only one dependent per principal is supported in this data model.`
        };
      }

      // Write dependent data to the *principal's row*
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_NAME).setValue(dependent.fullName);
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_RELATIONSHIP).setValue(dependent.relationship || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_AGE).setValue(dependent.age || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_DOB).setValue(dependent.dateOfBirth || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_TURNS_18).setValue(dependent.turns18 || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_AT_POST).setValue(dependent.atPost || 'No');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_ARRIVAL).setValue(dependent.arrivalDate || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_DEPARTURE).setValue(dependent.departureDate || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_VISA).setValue(dependent.visaNumber || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_VISA_EXP).setValue(dependent.visaExp || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_DIP_ID).setValue(dependent.dipId || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_DIP_EXP).setValue(dependent.dipIdExp || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_DEPENDENT_PWD).setValue(dependent.pwdStatus || 'No');

      console.log(`✅ Dependent recorded in row ${targetRow}`);
      
      return {
        success: true,
        row: targetRow
      };

    } catch (error) {
      console.error('❌ Error recording dependent:', error);
      return {
        success: false,
        message: 'Error recording dependent: ' + error.message
      };
    }
  }

  /**
   * ============================================================================
   * PRIVATE STAFF RECORDING - FIXED
   * ============================================================================
   * This REPLACES the old recordPrivateStaff function.
   * It writes to the principal's *own row* instead of looping 50 rows down.
   */
  function recordPrivateStaff(sheet, principalName, staff, principalRow) {
    try {
      console.log(`👷 Recording staff: ${staff.lastName}, ${staff.firstName} for ${principalName}`);
      
      const targetRow = principalRow;
      
      if (!targetRow) {
        return {
          success: false,
          message: `Principal "${principalName}" not found or row not provided`
        };
      }
      
      // Check if staff already exists on this row
      const existingName = sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_NAME).getValue();
      if (existingName && existingName.toString().trim() !== '') {
        return {
          success: false,
          message: `Staff (${existingName}) already exists for ${principalName} in row ${targetRow}. Only one staff member per principal is supported in this data model.`
        };
      }

      const fullName = `${staff.lastName}, ${staff.firstName}`;

      // Write staff data to the *principal's row*
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_NAME).setValue(fullName);
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_AGE).setValue(staff.age || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_DOB).setValue(staff.dateOfBirth || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_CITIZENSHIP).setValue(staff.citizenship || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_AT_POST).setValue(staff.atPost || 'No');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_ARRIVAL).setValue(staff.arrivalDate || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_DEPARTURE).setValue(staff.departureDate || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_VISA).setValue(staff.visaNumber || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_VISA_EXP).setValue(staff.visaExp || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_DIP_ID).setValue(staff.dipId || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_DIP_EXP).setValue(staff.dipIdExp || '');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_PWD).setValue(staff.pwdStatus || 'No');
      sheet.getRange(targetRow, COLUMN_CONFIG.COL_STAFF_EMERGENCY).setValue(staff.emergencyContact || '');

      console.log(`✅ Staff recorded in row ${targetRow}`);
      
      return {
        success: true,
        row: targetRow
      };

    } catch (error) {
      console.error('❌ Error recording staff:', error);
      return {
        success: false,
        message: 'Error recording staff: ' + error.message
      };
    }
  }

    // ============================================================================
    // HELPER FUNCTIONS
    // ============================================================================

  function findPrincipalRow(sheet, principalName) {
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLUMN_CONFIG.COL_PRINCIPAL_NAME - 1] === principalName) {
        return i + 1;
      }
    }
    
    return null;
    }

    // ============================================================================
    // UPDATE FUNCTIONS
    // ============================================================================

  function getPersonnelDataForUpdate(principalName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Personnel Tracking');
    
    if (!sheet) {
      throw new Error('Personnel Tracking sheet not found');
    }
    
    const principalRow = findPrincipalRow(sheet, principalName);
    
    if (!principalRow) {
      throw new Error('Principal not found: ' + principalName);
    }
    
    // ✅ FIX: Use COL_POST instead of COL_PRINCIPAL_POST
    const postStation = sheet.getRange(principalRow, COLUMN_CONFIG.COL_POST).getValue();
    const originalDeparture = sheet.getRange(principalRow, COLUMN_CONFIG.COL_DEPARTURE_DATE).getValue();
    const extended = sheet.getRange(principalRow, COLUMN_CONFIG.COL_EXTENDED).getValue();
    const newDeparture = sheet.getRange(principalRow, COLUMN_CONFIG.COL_CURRENT_DEPARTURE).getValue();
    const extensionDetails = sheet.getRange(principalRow, COLUMN_CONFIG.COL_EXTENSION_DETAILS).getValue();

    return {
      postStation: postStation || '',
      originalDeparture: originalDeparture ? Utilities.formatDate(new Date(originalDeparture), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      extended: extended || '',
      newDeparture: newDeparture ? Utilities.formatDate(new Date(newDeparture), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      extensionDetails: extensionDetails || ''
    };
  }



  function updatePersonnelData(principalName, updateData) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Personnel Tracking');
      
      if (!sheet) {
        throw new Error('Personnel Tracking sheet not found');
      }
      
      const principalRow = findPrincipalRow(sheet, principalName);
      
      if (!principalRow) {
        throw new Error('Principal not found');
      }
      
      // ✅ Update PRINCIPAL extension columns (P, Q, R)
      if (updateData.extended) {
        sheet.getRange(principalRow, 16).setValue(updateData.extended); // Column P
      }
      
      if (updateData.newDeparture) {
        sheet.getRange(principalRow, 17).setValue(new Date(updateData.newDeparture)); // Column Q
      }
      
      if (updateData.extensionDetails !== undefined) {
        const existingDetails = sheet.getRange(principalRow, 18).getValue(); // Column R
        
        let newDetails = updateData.extensionDetails;
        
        if (existingDetails && updateData.appendExtension) {
          newDetails = existingDetails + '\n' + updateData.extensionDetails;
        }
        
        sheet.getRange(principalRow, 18).setValue(newDetails); // Column R
      }
      
      // ✅ Update DEPENDENT extension columns (AK, AL, AM) if dependent exists
      const rowData = sheet.getRange(principalRow, 1, 1, 58).getValues()[0];
      const depName = rowData[18]; // Column S (index 18)
      
      if (depName && depName.toString().trim() !== '' && updateData.extended && updateData.newDeparture) {
        sheet.getRange(principalRow, 37).setValue(updateData.extended); // Column AK
        sheet.getRange(principalRow, 38).setValue(new Date(updateData.newDeparture)); // Column AL
        sheet.getRange(principalRow, 39).setValue(updateData.extensionDetails || ''); // Column AM
      }
      
      // ✅ Update STAFF extension columns (BD, BE, BF) if staff exists
      const staffName = rowData[39]; // Column AN (index 39)
      
      if (staffName && staffName.toString().trim() !== '' && updateData.extended && updateData.newDeparture) {
        sheet.getRange(principalRow, 56).setValue(updateData.extended); // Column BD
        sheet.getRange(principalRow, 57).setValue(new Date(updateData.newDeparture)); // Column BE
        sheet.getRange(principalRow, 58).setValue(updateData.extensionDetails || ''); // Column BF
      }
      
      SpreadsheetApp.flush();
      
      // ✅ If checkbox for extending attendance is checked
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

  function extendFamilyAttendance(sheet, principalName, principalRow, newDepartureDate) {
    try {
      console.log('👨‍👩‍👧‍👦 Extending family attendance to:', newDepartureDate);
      
      const maxSearchRows = 50;
      
      // Update dependents
      for (let i = 0; i < maxSearchRows; i++) {
        const checkRow = principalRow + i;
        const depName = sheet.getRange(checkRow, COLUMN_CONFIG.COL_DEPENDENT_NAME).getValue();
        const atPost = sheet.getRange(checkRow, COLUMN_CONFIG.COL_DEPENDENT_AT_POST).getValue();
        
        if (!depName) continue;
        
        if (atPost === 'Yes' || atPost === 'yes') {
          sheet.getRange(checkRow, COLUMN_CONFIG.COL_DEPENDENT_DEPARTURE).setValue(newDepartureDate);
          console.log(`  ✅ Extended dependent in row ${checkRow}`);
        }
      }
      
      // Update staff
      for (let i = 0; i < maxSearchRows; i++) {
        const checkRow = principalRow + i;
        const staffName = sheet.getRange(checkRow, COLUMN_CONFIG.COL_STAFF_NAME).getValue();
        const atPost = sheet.getRange(checkRow, COLUMN_CONFIG.COL_STAFF_AT_POST).getValue();
        
        if (!staffName) continue;
        
        if (atPost === 'Yes' || atPost === 'yes') {
          sheet.getRange(checkRow, COLUMN_CONFIG.COL_STAFF_DEPARTURE).setValue(newDepartureDate);
          console.log(`  ✅ Extended staff in row ${checkRow}`);
        }
      }
      
      return true;

    } catch (error) {
      console.error('❌ Error extending family:', error);
      return false;
    }
    }

// ============================================================================
// Replace the old getExistingPrincipalsFromTracking function in attendanceTracker.gs
// with this new version.
// ============================================================================

  function getExistingPrincipalsFromTracking() {
    try {
      // 1. Get all principals from the main tracking sheet
      // This function (getAllPrincipals) already returns the unique list
      const allPrincipals = getAllPrincipals(); 
      
      if (allPrincipals.length === 0) {
        return [];
      }

      // 2. Get the list of departed principals
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // This sheet name is based on your 'withdrawEntireFamily' function
      const departureSheet = ss.getSheetByName('Principal Departure Log'); 
      
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
        // Return the full list if the log sheet doesn't exist
        return allPrincipals;
      }

      // 3. Filter the 'allPrincipals' list
      const activePrincipals = allPrincipals.filter(principal => {
        // Keep the principal ONLY IF they are NOT in the departedNames set
        return !departedNames.has(principal.fullName);
      });
      
      Logger.log(`Filtered principals: Total ${allPrincipals.length}, Departed ${departedNames.size}, Active ${activePrincipals.length}`);
      return activePrincipals;

    } catch (error) {
      console.error('Error in getExistingPrincipalsFromTracking:', error);
      return []; // Return empty on error
    }
  }

    /**
     * Extend attendance for principal's active family members
     */
  function extendAttendanceForFamily(principalName) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const personnelSheet = ss.getSheetByName('Personnel Tracking');
      const attendanceSheet = ss.getSheetByName('Attendance Log for Dependent & Staff');
      const archiveSheet = ss.getSheetByName('Archived Withdrawals');
      
      if (!attendanceSheet) {
        return {
          success: false,
          message: 'Attendance sheet not found',
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
      
      // Find all family members
      const data = personnelSheet.getDataRange().getValues();
      const activeFamilyMembers = [];
      const archivedFamilyMembers = [];
      
      for (let i = 2; i < data.length; i++) {
        const rowPrincipal = data[i][1]; // Column B (index 1)
        
        if (rowPrincipal === principalName) {
          const dependent = data[i][18]; // Column S (position 19, index 18) ✅
          const staff = data[i][39];      // Column AN (position 40, index 39) ✅
          
          // Check dependent
          if (dependent && dependent.toString().trim()) {
            const pairName = `${principalName} - ${dependent}`;
            
            if (archivedPairs.has(pairName)) {
              archivedFamilyMembers.push({
                name: dependent,
                type: 'Dependent',
                pairName: pairName
              });
            } else {
              activeFamilyMembers.push({
                name: dependent,
                type: 'Dependent',
                pairName: pairName
              });
            }
          }
          
          // Check staff
          if (staff && staff.toString().trim()) {
            const pairName = `${principalName} - ${staff}`;
            
            if (archivedPairs.has(pairName)) {
              archivedFamilyMembers.push({
                name: staff,
                type: 'Staff',
                pairName: pairName
              });
            } else {
              activeFamilyMembers.push({
                name: staff,
                type: 'Staff',
                pairName: pairName
              });
            }
          }
        }
      }
      
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
            'At Post',
            'Principal extension - auto-recorded',
            currentDate
          ]]);
          
          attendanceSheet.getRange(nextRow, 5).setNumberFormat('yyyy-mm-dd hh:mm:ss');
          recordedCount++;
          
          Logger.log(`✅ Recorded attendance for: ${member.pairName}`);
        } catch (error) {
          Logger.log(`⚠️ Error recording attendance for ${member.pairName}: ${error}`);
        }
      });
      
      // Build result message
      let message = '';
      
      if (recordedCount > 0) {
        message += `✅ Attendance extended for ${recordedCount} active family member(s)`;
      }
      
      if (archivedFamilyMembers.length > 0) {
        message += `\n⚠️ Skipped ${archivedFamilyMembers.length} archived member(s)`;
      }
      
      if (recordedCount === 0 && archivedFamilyMembers.length === 0) {
        message = 'ℹ️ No dependents/staff found for this principal';
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
      console.error('Error extending family attendance:', error);
      return {
        success: false,
        message: `Failed to extend attendance: ${error.message}`,
        extended: 0,
        skipped: 0
      };
    }
    }

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

    // ============================================================================
    // DEPARTURE RECORDING - FIXED SYNTAX ERROR
    // ============================================================================

  function recordPrincipalDeparture(principalName, departureDate, reason) {
    try {
      console.log('✈️ Recording departure for:', principalName);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const trackingSheet = ss.getSheetByName('Personnel Tracking');
      const archiveSheet = ss.getSheetByName('Archived Withdrawals');
      
      if (!trackingSheet) {
        throw new Error('Sheet "Personnel Tracking" not found');
      }
      
      if (!archiveSheet) {
        throw new Error('Sheet "Archive" not found');
      }

      const principalRow = findPrincipalRow(trackingSheet, principalName);
      
      if (!principalRow) {
        return {
          success: false,
          message: `Principal "${principalName}" not found`
        };
      }

      // Update departure date in tracking sheet
      trackingSheet.getRange(principalRow, COLUMN_CONFIG.COL_DEPARTURE_DATE).setValue(departureDate);
      
      // Archive the principal and family
      const archiveResult = archivePrincipalAndFamily(trackingSheet, archiveSheet, principalName, principalRow, departureDate, reason);
      
      if (!archiveResult.success) {
        return archiveResult;
      }

      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: '✅ Departure recorded and archived successfully!',
        row: principalRow
      }; // FIXED: Removed extra closing brace here
      
    } catch (error) {
      console.error('❌ Departure recording error:', error);
      return {
        success: false,
        message: 'Error recording departure: ' + error.message
      };
    }
    }

    // ============================================================================
    // ARCHIVE FUNCTIONS
    // ============================================================================

  /**
   * ============================================================================
   * ARCHIVE FUNCTIONS - FIXED
   * ============================================================================
   * This REPLACES the old archivePrincipalAndFamily function.
   * It archives the principal's row and checks ONLY that same row
   * for a dependent/staff, matching the "horizontal" data model.
   */
  function archivePrincipalAndFamily(trackingSheet, archiveSheet, principalName, principalRow, departureDate, reason) {
    try {
      const archiveData = [];
      const timestamp = new Date();
      
      // 1. Get the ENTIRE row data for the principal in ONE call
      // We read all 58 columns (or to the last column)
      const rowData = trackingSheet.getRange(principalRow, 1, 1, trackingSheet.getLastColumn()).getValues()[0];
      
      // 2. Archive the principal
      // We add "Principal" as the type (assuming archive has an extra column for this)
      archiveData.push([...rowData, timestamp, reason, 'Principal']);
      
      // 3. Check for a dependent ON THE SAME ROW
      const depName = rowData[COLUMN_CONFIG.COL_DEPENDENT_NAME - 1]; // Use 0-based index
      if (depName && depName.toString().trim() !== '') {
        // Archive the same row, but mark as 'Dependent'
        archiveData.push([...rowData, timestamp, reason, 'Dependent']);
      }
      
      // 4. Check for staff ON THE SAME ROW
      const staffName = rowData[COLUMN_CONFIG.COL_STAFF_NAME - 1]; // Use 0-based index
      if (staffName && staffName.toString().trim() !== '') {
        // Archive the same row, but mark as 'Staff'
        archiveData.push([...rowData, timestamp, reason, 'Staff']);
      }

      // 🛑 The "for (let i = 0; i < maxSearchRows; i++)" LOOP is REMOVED
      
      // 5. Append all records (1-3 rows) to the archive
      if (archiveData.length > 0) {
        const archiveLastRow = archiveSheet.getLastRow();
        archiveSheet.getRange(archiveLastRow + 1, 1, archiveData.length, archiveData[0].length).setValues(archiveData);
      }
      
      console.log(`✅ Archived ${archiveData.length} records for ${principalName} from row ${principalRow}`);
      
      return {
        success: true,
        archivedCount: archiveData.length
      };

    } catch (error) {
      console.error('❌ Archive error:', error);
      return {
        success: false,
        message: 'Error archiving: ' + error.message
      };
    }
  }

    // ============================================================================
    // WITHDRAWAL FUNCTIONS
    // ============================================================================

  function withdrawEntireFamily(principalName, departureDate, reason) {
    try {
      console.log('🚪 Withdrawing entire family for:', principalName);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const trackingSheet = ss.getSheetByName('Personnel Tracking');
      const principalDepartureLog = ss.getSheetByName('Principal Departure Log');
      const attendanceLog = ss.getSheetByName('Attendance Log for Dependent & Staff');
      const archivedWithdrawals = ss.getSheetByName('Archived Withdrawals');
      
      if (!trackingSheet || !principalDepartureLog || !attendanceLog || !archivedWithdrawals) {
        throw new Error('Required sheets not found');
      }

      const principalRow = findPrincipalRow(trackingSheet, principalName);
      
      if (!principalRow) {
        return {
          success: false,
          message: `Principal "${principalName}" not found`
        };
      }

      const today = new Date();
      const post = trackingSheet.getRange(principalRow, COLUMN_CONFIG.COL_POST).getValue();
      let withdrawnCount = 1;
      
      // Log principal to Principal Departure Log
      const principalLogLastRow = principalDepartureLog.getLastRow();
      principalDepartureLog.getRange(principalLogLastRow + 1, 1, 1, 6).setValues([[
        today,
        principalName,
        post,
        reason,
        departureDate,
        Session.getActiveUser().getEmail()
      ]]);
      
      // Get dependents and staff from principal's row only
      const depName = trackingSheet.getRange(principalRow, COLUMN_CONFIG.COL_DEPENDENT_NAME).getValue();
      const staffName = trackingSheet.getRange(principalRow, COLUMN_CONFIG.COL_STAFF_NAME).getValue();
      
      // Process dependent if exists
      if (depName) {
        const pairName = principalName + ' - ' + depName;
        
        // Add to Attendance Log
        const attendanceLogLastRow = attendanceLog.getLastRow();
        attendanceLog.getRange(attendanceLogLastRow + 1, 1, 1, 5).setValues([[
          pairName,
          'Q1',
          'Not At Post',
          reason,
          departureDate
        ]]);
        
        // Add to Archived Withdrawals
        const archiveLastRow = archivedWithdrawals.getLastRow();
        archivedWithdrawals.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
          pairName,
          'Q1',
          reason,
          departureDate,
          today
        ]]);
        
        withdrawnCount++;
      }
      
      // Process staff if exists
      if (staffName) {
        const pairName = principalName + ' - ' + staffName;
        
        // Add to Attendance Log
        const attendanceLogLastRow = attendanceLog.getLastRow();
        attendanceLog.getRange(attendanceLogLastRow + 1, 1, 1, 5).setValues([[
          pairName,
          'Q1',
          'Not At Post',
          reason,
          departureDate
        ]]);
        
        // Add to Archived Withdrawals
        const archiveLastRow = archivedWithdrawals.getLastRow();
        archivedWithdrawals.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
          pairName,
          'Q1',
          reason,
          departureDate,
          today
        ]]);
        
        withdrawnCount++;
      }

      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: `✅ Successfully withdrawn ${withdrawnCount} person(s)`,
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

    // ============================================================================
    // DATA RETRIEVAL FUNCTIONS
    // ============================================================================

  function getAllPrincipals() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Personnel Tracking');
      
      if (!sheet) {
        return [];
      }

      const data = sheet.getDataRange().getValues();
      const principals = [];
      const seenPrincipals = new Set(); // ✅ Track unique principals
      
      // Start from row 3 (index 2) - skip header rows
      for (let i = 2; i < data.length; i++) {
        const principalName = data[i][1]; // Column B (index 1)
        const post = data[i][0]; // Column A (index 0)
        
        // ✅ Only add if:
        // 1. Principal name exists
        // 2. We haven't seen this principal before (skip duplicate rows for dependents/staff)
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

  function getFamilyMembersStatus(principalName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.getSheetByName('Personnel Tracking');
    const archiveSheet = ss.getSheetByName('Archived Withdrawals');
    const logSheet = ss.getSheetByName('Attendance Log for Dependent & Staff');
    
    const principalRow = findPrincipalRow(trackingSheet, principalName);
    if (!principalRow) return { active: [], archived: [] };

    const withdrawnInfo = {};
    
    // Get archived withdrawals
    if (archiveSheet && archiveSheet.getLastRow() >= 2) {
      const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 5).getValues();
      archiveData.forEach(row => {
        const fullName = (row[0] || '').toString().trim();
        const withdrawalDate = row[3];
        if (fullName) {
          const parts = fullName.split(' - ');
          if (parts.length === 2) {
            withdrawnInfo[parts[1].trim()] = { date: withdrawalDate ? new Date(withdrawalDate) : null };
          }
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
        if (fullName && date) {
          const logDate = new Date(date);
          if (!attendanceLog[fullName] || logDate > attendanceLog[fullName].date) {
            attendanceLog[fullName] = { attendance, date: logDate };
          }
        }
      });
      
      Object.keys(attendanceLog).forEach(fullName => {
        if (attendanceLog[fullName].attendance === 'Not At Post') {
          const parts = fullName.split(' - ');
          if (parts.length === 2 && !withdrawnInfo[parts[1].trim()]) {
            withdrawnInfo[parts[1].trim()] = { date: attendanceLog[fullName].date };
          }
        }
      });
    }

    const active = [];
    const archived = [];
    
    // ✅ Read ONLY from the principal's own row (not 50 rows below)
    const rowData = trackingSheet.getRange(principalRow, 1, 1, 58).getValues()[0];
    
    // Check dependent (Column S, index 18)
    const depName = rowData[18];
    if (depName && depName.toString().trim() !== '') {
      const name = depName.toString().trim();
      if (withdrawnInfo[name]) {
        archived.push({ 
          name, 
          type: 'Dependent', 
          withdrawalDate: withdrawnInfo[name].date ? Utilities.formatDate(withdrawnInfo[name].date, Session.getScriptTimeZone(), 'MMM dd, yyyy') : 'Unknown' 
        });
      } else {
        active.push({ name, type: 'Dependent' });
      }
    }
    
    // Check staff (Column AN, index 39)
    const staffName = rowData[39];
    if (staffName && staffName.toString().trim() !== '') {
      const name = staffName.toString().trim();
      if (withdrawnInfo[name]) {
        archived.push({ 
          name, 
          type: 'Staff', 
          withdrawalDate: withdrawnInfo[name].date ? Utilities.formatDate(withdrawnInfo[name].date, Session.getScriptTimeZone(), 'MMM dd, yyyy') : 'Unknown' 
        });
      } else {
        active.push({ name, type: 'Staff' });
      }
    }
    
    return { active, archived };
  }

  function getFamilyMembersForWithdrawal(principalName) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const trackingSheet = ss.getSheetByName('Personnel Tracking');
      const archiveSheet = ss.getSheetByName('Archived Withdrawals');
      const logSheet = ss.getSheetByName('Attendance Log for Dependent & Staff');
      
      if (!trackingSheet) {
        return { principalName: principalName, active: [], withdrawn: [] };
      }

      const principalRow = findPrincipalRow(trackingSheet, principalName);
      
      if (!principalRow) {
        return { principalName: principalName, active: [], withdrawn: [] };
      }

      // ✅ Get withdrawn info from archives
      const withdrawnInfo = {};
      
      if (archiveSheet && archiveSheet.getLastRow() >= 2) {
        const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 5).getValues();
        archiveData.forEach(row => {
          const fullName = (row[0] || '').toString().trim();
          const withdrawalDate = row[3];
          if (fullName) {
            const parts = fullName.split(' - ');
            if (parts.length === 2) {
              withdrawnInfo[parts[1].trim()] = { date: withdrawalDate ? new Date(withdrawalDate) : null };
            }
          }
        });
      }

      // ✅ Check attendance log for "Not At Post" status
      if (logSheet && logSheet.getLastRow() >= 2) {
        const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 5).getValues();
        const attendanceLog = {};
        logData.forEach(row => {
          const fullName = (row[0] || '').toString().trim();
          const attendance = (row[2] || '').toString().trim();
          const date = row[4];
          if (fullName && date) {
            const logDate = new Date(date);
            if (!attendanceLog[fullName] || logDate > attendanceLog[fullName].date) {
              attendanceLog[fullName] = { attendance, date: logDate };
            }
          }
        });
        
        Object.keys(attendanceLog).forEach(fullName => {
          if (attendanceLog[fullName].attendance === 'Not At Post') {
            const parts = fullName.split(' - ');
            if (parts.length === 2 && !withdrawnInfo[parts[1].trim()]) {
              withdrawnInfo[parts[1].trim()] = { date: attendanceLog[fullName].date };
            }
          }
        });
      }

      const active = [];
      const withdrawn = [];
      
      // ✅ Read ONLY from the principal's row (Column S = index 19, Column AN = index 40)
      const rowData = trackingSheet.getRange(principalRow, 1, 1, 58).getValues()[0];
      
      // Check dependent (Column S, index 18 in 0-based array)
      const depName = rowData[18]; // Column S
      if (depName && depName.toString().trim() !== '') {
        const name = depName.toString().trim();
        if (withdrawnInfo[name]) {
          withdrawn.push({ 
            name, 
            type: 'Dependent', 
            withdrawalDate: withdrawnInfo[name].date ? Utilities.formatDate(withdrawnInfo[name].date, Session.getScriptTimeZone(), 'MMM dd, yyyy') : 'Unknown' 
          });
        } else {
          active.push({ name, type: 'Dependent' });
        }
      }
      
      // Check staff (Column AN, index 39 in 0-based array)
      const staffName = rowData[39]; // Column AN
      if (staffName && staffName.toString().trim() !== '') {
        const name = staffName.toString().trim();
        if (withdrawnInfo[name]) {
          withdrawn.push({ 
            name, 
            type: 'Staff', 
            withdrawalDate: withdrawnInfo[name].date ? Utilities.formatDate(withdrawnInfo[name].date, Session.getScriptTimeZone(), 'MMM dd, yyyy') : 'Unknown' 
          });
        } else {
          active.push({ name, type: 'Staff' });
        }
      }
      
      return {
        principalName: principalName,
        active: active,
        withdrawn: withdrawn
      };

    } catch (error) {
      console.error('Error getting family for withdrawal:', error);
      return { principalName: principalName, active: [], withdrawn: [] };
    }
  }


  function getPrincipalsList() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("Principals List");
      if (!sheet) return [];

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];

      const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A:C

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


  function debugDOBColumns() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Principals List");
    const sample = sheet.getRange(2, 1, 1, 10).getValues()[0];
    Logger.log(sample);
  }

    // ============================================================================
    // UI FUNCTIONS
    // ============================================================================

  function showEntryForm() {
    const html = HtmlService.createHtmlOutputFromFile('EntryForm')
      .setWidth(1000)
      .setHeight(700)
      .setTitle('Personnel Entry Form');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Personnel Entry Form');
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
      .addItem('📊 Attendance Tracker', 'openAttendanceTracker')  // ADD THIS LINE
      .addItem('View Records', 'showTrackingSheet')
      .addItem('Extension Form', 'showUpdateForm')
      .addSeparator()
      .addItem('Manage Principals', 'showPrincipalsSheet')
      .addItem('Update Principal List', 'updatePrincipal')
      .addToUi();
    }

    // ============================================================================
    // LOAD ATTENDANCE DATA FOR UI - CORRECTED VERSION
    // ============================================================================
    // ============================================================================
    // SAVE ATTENDANCE - Matches your log sheet structure
    // ============================================================================
  function saveAttendance(pairName, status, remarks) {
    try {
      Logger.log('=== saveAttendance START ===');
      Logger.log('Pair: ' + pairName);
      Logger.log('Status: ' + status);
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const trackingSheet = ss.getSheetByName(CONFIG.SHEET_PERSONNEL_TRACKING);
      const logSheet = ss.getSheetByName(CONFIG.SHEET_ATTENDANCE_LOG);
      const archiveSheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVED_WITHDRAWALS);
      
      if (!trackingSheet) {
        return { success: false, message: CONFIG.SHEET_PERSONNEL_TRACKING + ' not found' };
      }
      
      if (!logSheet) {
        return { success: false, message: CONFIG.SHEET_ATTENDANCE_LOG + ' not found' };
      }
      
      const parts = pairName.split(' - ');
      if (parts.length !== 2) {
        return { success: false, message: 'Invalid pair name format' };
      }
      
      const principalName = parts[0].trim();
      const memberName = parts[1].trim();
      const fullPairName = pairName;
      
      const lastRow = trackingSheet.getLastRow();
      const principalCol = trackingSheet.getRange(3, 2, lastRow - 2, 1).getValues();
      const dependentCol = trackingSheet.getRange(3, 19, lastRow - 2, 1).getValues();
      const staffCol = trackingSheet.getRange(3, 40, lastRow - 2, 1).getValues();
      
      let memberType = '';
      let rowNumber = -1;
      
      for (let i = 0; i < principalCol.length; i++) {
        const principal = (principalCol[i][0] || '').toString().trim();
        
        if (principal === principalName) {
          const dependent = (dependentCol[i][0] || '').toString().trim();
          if (dependent === memberName) {
            memberType = 'Dependent';
            rowNumber = i + 3;
            break;
          }
          
          const staff = (staffCol[i][0] || '').toString().trim();
          if (staff === memberName) {
            memberType = 'Staff';
            rowNumber = i + 3;
            break;
          }
        }
      }
      
      if (rowNumber === -1) {
        return { success: false, message: 'Pair not found' };
      }
      
      const quarter = calculateCurrentQuarter(logSheet, fullPairName);
      const today = new Date();
      
      // NOT At Post - Save to BOTH Archive AND Attendance Log
      if (status === CONFIG.STATUS_NOT_AT_POST) {
        if (!archiveSheet) {
          return { success: false, message: CONFIG.SHEET_ARCHIVED_WITHDRAWALS + ' not found' };
        }
        
        // Save to Archived Withdrawals
        const archiveLastRow = archiveSheet.getLastRow();
        archiveSheet.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
          fullPairName, quarter, remarks || '', today, today
        ]]);
        
        // ALSO save to Attendance Log
        const logLastRow = logSheet.getLastRow();
        logSheet.getRange(logLastRow + 1, 1, 1, 5).setValues([[
          fullPairName, quarter, status, remarks || '', today
        ]]);
        
        // Update Personnel Tracking
        if (memberType === 'Dependent') {
          trackingSheet.getRange(rowNumber, 25).setValue(CONFIG.STATUS_NOT_AT_POST);
        } else {
          trackingSheet.getRange(rowNumber, 45).setValue(CONFIG.STATUS_NOT_AT_POST);
        }
        
        Logger.log('Saved to both Archive and Attendance Log');
        return { success: true, message: memberName + ' archived as withdrawn.' };
      }
      
      // At Post - Save to Attendance Log
      const logLastRow = logSheet.getLastRow();
      logSheet.getRange(logLastRow + 1, 1, 1, 5).setValues([[
        fullPairName, quarter, status, remarks || '', today
      ]]);
      
      if (memberType === 'Dependent') {
        trackingSheet.getRange(rowNumber, 25).setValue(CONFIG.STATUS_AT_POST);
      } else {
        trackingSheet.getRange(rowNumber, 45).setValue(CONFIG.STATUS_AT_POST);
      }
      
      Logger.log('Recorded to attendance log as ' + quarter);
      return { success: true, message: 'Recorded as ' + quarter + '. Next window: ' + CONFIG.ATTENDANCE_CYCLE_DAYS + ' days.' };
      
    } catch (error) {
      Logger.log('ERROR: ' + error.toString());
      return { success: false, message: 'Error: ' + error.message };
    }
    }


    // ============================================================================
    // CALCULATE CURRENT QUARTER from log history
    // ============================================================================
  function calculateCurrentQuarter(logSheet, pairName) {
    try {
      const logLastRow = logSheet.getLastRow();
      
      if (logLastRow < 2) {
        // No records yet, this is Q1
        return 'Q1';
      }
      
      // Read all log data (skip header)
      const logData = logSheet.getRange(2, 1, logLastRow - 1, 2).getValues();
      
      // Count how many times this pair has been recorded
      const attendanceCount = 0;
      
      for (let i = 0; i < logData.length; i++) {
        const name = (logData[i][0] || '').toString().trim();
        if (name === pairName) {
          attendanceCount++;
        }
      }
      
      // Next quarter is current count + 1
      const nextQuarter = attendanceCount + 1;
      
      if (nextQuarter > CONFIG.MAX_QUARTERS) {
      return 'Q' + CONFIG.MAX_QUARTERS + '+';
      }
      
      return 'Q' + nextQuarter;
      
    } catch (error) {
      Logger.log('Error calculating quarter: ' + error);
      return 'Q1';
    }
    }

    // ============================================================================
    // LOAD ATTENDANCE DATA - With window calculation from log
    // ============================================================================
  // Helper function to calculate window - moved here for scope
  function calculateWindowHelper(attendanceRecord, today) {
    if (!attendanceRecord) {
      return {
        isNewPair: true,
        quarter: 'Q1',
        isOpen: true,
        status: 'Ready',
        message: 'Ready to Start'
      };
    }
    
    const lastDate = attendanceRecord.date;
    const lastQuarter = attendanceRecord.quarter;
    const daysSinceLastAttendance = Math.floor((today - lastDate) / (1000 * 60 * 60 * 24));
    
    const lastQNum = parseInt(lastQuarter.replace('Q', '').replace('+', '')) || 1;
    const nextQNum = lastQNum + 1;
    const nextQuarter = nextQNum > CONFIG.MAX_QUARTERS ? ('Q' + CONFIG.MAX_QUARTERS + '+') : ('Q' + nextQNum);
    
    const windowInfo = {
      isNewPair: false,
      quarter: nextQuarter,
      daysSinceLastAttendance: daysSinceLastAttendance,
      lastAttendanceDate: lastDate,
      windowEndDay: CALCULATED.WINDOW_END_DAY,
      gracePeriodDays: CONFIG.DAYS_AFTER_MISSED_WINDOW
    };
    
    if (daysSinceLastAttendance >= CALCULATED.WINDOW_START_DAY && daysSinceLastAttendance <= CALCULATED.WINDOW_END_DAY) {
      const daysRemaining = CALCULATED.WINDOW_END_DAY - daysSinceLastAttendance;
      windowInfo.isOpen = true;
      windowInfo.status = 'Open';
      windowInfo.daysRemaining = daysRemaining;
      windowInfo.message = 'Window Open (Day ' + daysSinceLastAttendance + ')';
    } else if (daysSinceLastAttendance < CALCULATED.WINDOW_START_DAY) {
      const daysUntilWindow = CALCULATED.WINDOW_START_DAY - daysSinceLastAttendance;
      windowInfo.isOpen = false;
      windowInfo.status = 'Upcoming';
      windowInfo.message = daysUntilWindow + ' days until window';
    } else {
      const daysOverdue = daysSinceLastAttendance - CALCULATED.WINDOW_END_DAY;
      windowInfo.isOpen = false;
      windowInfo.status = 'Missed';
      windowInfo.message = 'Missed by ' + daysOverdue + ' days';
    }
    
    return windowInfo;
  }

// ============================================================================
// FIXED loadAttendanceData() Function
// ============================================================================
// This replaces the existing loadAttendanceData function in attendanceTracker.gs
// Starting around line 1339

  function loadAttendanceData() {
    Logger.log('=== loadAttendanceData START ===');
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const trackingSheet = ss.getSheetByName('Personnel Tracking');
      const logSheet = ss.getSheetByName('Attendance Log for Dependent & Staff');
      const archiveSheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVED_WITHDRAWALS);
      
      if (!trackingSheet) {
        Logger.log('ERROR: Personnel Tracking sheet not found');
        return [];
      }
      
      // ✅ Start reading from row 3
      const lastRow = trackingSheet.getLastRow();
      if (lastRow < 3) {
        Logger.log('No data rows');
        return [];
      }
      
      const data = trackingSheet.getRange(3, 1, lastRow - 2, 58).getValues(); // ← row 3 start
      
      // ============================================================================
      // LOAD ARCHIVED WITHDRAWALS - CHECK 1
      // ============================================================================
      const archivedPairs = new Set();
      if (archiveSheet && archiveSheet.getLastRow() >= 2) {
        const archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, 1).getValues();
        archiveData.forEach(row => {
          const name = (row[0] || '').toString().trim();
          if (name) {
            archivedPairs.add(name);
          }
        });
        Logger.log('Loaded ' + archivedPairs.size + ' archived pairs');
      }
      
      // ============================================================================
      // LOAD ATTENDANCE LOG - CHECK 2
      // ============================================================================
      const attendanceLog = {};
      if (logSheet && logSheet.getLastRow() >= 2) {
        const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 5).getValues();
        logData.forEach(row => {
          const name = (row[0] || '').toString().trim();
          const quarter = (row[1] || '').toString().trim();
          const attendance = (row[2] || '').toString().trim();
          const date = row[4];
          if (name && date) {
            const logDate = new Date(date);
            if (!attendanceLog[name] || logDate > attendanceLog[name].date) {
              attendanceLog[name] = { date: logDate, quarter, attendance };
            }
          }
        });
        Logger.log('Loaded attendance log for ' + Object.keys(attendanceLog).length + ' pairs');
      }
      
      const attendancePairs = [];
      const today = new Date();
      let skippedCount = 0;
      
      data.forEach((row, index) => {
        const principalName = (row[1] || '').toString().trim(); // ✅ Always trim
        if (!principalName) return;
        
        const post = (row[0] || '').toString().trim();

        // ========================================================================
        // DEPENDENT PROCESSING
        // ========================================================================
        const dependentName = (row[18] || '').toString().trim();
        if (dependentName) {
          const pairName = `${principalName} - ${dependentName}`;
          
          // ✅ CHECK 1: Skip if in Archived Withdrawals
          if (archivedPairs.has(pairName)) {
            Logger.log('SKIPPED (Archived): ' + pairName);
            skippedCount++;
            return; // Skip this dependent
          }
          
          // ✅ CHECK 2: Skip if last attendance was "Not At Post"
          const lastRecord = attendanceLog[pairName];
          if (lastRecord && lastRecord.attendance === 'Not At Post') {
            Logger.log('SKIPPED (Not At Post): ' + pairName);
            skippedCount++;
            return; // Skip this dependent
          }
          
          // ✅ This person is ACTIVE - add to attendance tracker
          const windowInfo = calculateWindowHelper(attendanceLog[pairName], today);
          attendancePairs.push({
            name: pairName,
            principalName,
            memberName: dependentName,
            memberType: 'Dependent',
            post,
            atPost: (row[24] || 'No').toString().trim(),
            quarter: windowInfo.quarter,
            status: 'Active',
            isNewPair: windowInfo.isNewPair,
            windowInfo,
            rowNumber: index + 3 // adjusted for row 3 start
          });
        }

        // ========================================================================
        // STAFF PROCESSING
        // ========================================================================
        const staffName = (row[39] || '').toString().trim();
        if (staffName) {
          const pairName = `${principalName} - ${staffName}`;
          
          // ✅ CHECK 1: Skip if in Archived Withdrawals
          if (archivedPairs.has(pairName)) {
            Logger.log('SKIPPED (Archived): ' + pairName);
            skippedCount++;
            return; // Skip this staff
          }
          
          // ✅ CHECK 2: Skip if last attendance was "Not At Post"
          const lastRecord = attendanceLog[pairName];
          if (lastRecord && lastRecord.attendance === 'Not At Post') {
            Logger.log('SKIPPED (Not At Post): ' + pairName);
            skippedCount++;
            return; // Skip this staff
          }
          
          // ✅ This person is ACTIVE - add to attendance tracker
          const windowInfo = calculateWindowHelper(attendanceLog[pairName], today);
          attendancePairs.push({
            name: pairName,
            principalName,
            memberName: staffName,
            memberType: 'Staff',
            post,
            atPost: (row[44] || 'No').toString().trim(),
            quarter: windowInfo.quarter,
            status: 'Active',
            isNewPair: windowInfo.isNewPair,
            windowInfo,
            rowNumber: index + 3
          });
        }
      });

      // ✅ Convert Date objects to strings for frontend
      attendancePairs.forEach(pair => {
        if (pair.windowInfo && pair.windowInfo.lastAttendanceDate instanceof Date) {
          pair.windowInfo.lastAttendanceDate = pair.windowInfo.lastAttendanceDate.toISOString();
        }
      });
      
      Logger.log('TOTAL PAIRS: ' + attendancePairs.length);
      Logger.log('SKIPPED: ' + skippedCount);
      return attendancePairs;
      
    } catch (error) {
      Logger.log('ERROR loading attendance data: ' + error);
      return [];
    }
  }


  // ============================================================================
  // CALCULATE WINDOW
  // ============================================================================
  function calculateWindow(attendanceRecord, today) {
    if (!attendanceRecord) {
    return {
      isNewPair: true,
      quarter: 'Q1',
      isOpen: true,
      status: 'Ready',
      message: 'Ready to Start'
    };
    }
    
    const lastDate = attendanceRecord.date;
    const lastQuarter = attendanceRecord.quarter;
    const daysSinceLastAttendance = Math.floor((today - lastDate) / (1000 * 60 * 60 * 24));
    
    const lastQNum = parseInt(lastQuarter.replace('Q', '').replace('+', '')) || 1;
    const nextQNum = lastQNum + 1;
    const nextQuarter = nextQNum > CONFIG.MAX_QUARTERS ? ('Q' + CONFIG.MAX_QUARTERS + '+') : ('Q' + nextQNum);
    
    const windowInfo = {
    isNewPair: false,
    quarter: nextQuarter,
    daysSinceLastAttendance: daysSinceLastAttendance,
    lastAttendanceDate: lastDate,
    windowEndDay: CALCULATED.WINDOW_END_DAY,         
    gracePeriodDays: CONFIG.DAYS_AFTER_MISSED_WINDOW
    };
    
    if (daysSinceLastAttendance >= CALCULATED.WINDOW_START_DAY && daysSinceLastAttendance <= CALCULATED.WINDOW_END_DAY) {
    const daysRemaining = CALCULATED.WINDOW_END_DAY - daysSinceLastAttendance;
    windowInfo.isOpen = true;
    windowInfo.status = 'Open';
    windowInfo.daysRemaining = daysRemaining;
    windowInfo.message = 'Window Open (Day ' + daysSinceLastAttendance + ')';
    } else if (daysSinceLastAttendance < CALCULATED.WINDOW_START_DAY) {
    const daysUntilWindow = CALCULATED.WINDOW_START_DAY - daysSinceLastAttendance;
    windowInfo.isOpen = false;
    windowInfo.status = 'Upcoming';
    windowInfo.message = daysUntilWindow + ' days until window';
    } else {
    const daysOverdue = daysSinceLastAttendance - CALCULATED.WINDOW_END_DAY;
    windowInfo.isOpen = false;
    windowInfo.status = 'Missed';
    windowInfo.message = 'Missed by ' + daysOverdue + ' days';
    }
    
    return windowInfo;
  }

  function autoArchiveMissedWindows() {
    try {
    Logger.log('=== AUTO-ARCHIVE MISSED WINDOWS START ===');
    
    if (!CONFIG.AUTO_ARCHIVE_MISSED_WINDOWS) {
      return { success: true, message: 'Auto-archive disabled', archived: 0 };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = ss.getSheetByName(CONFIG.SHEET_PERSONNEL_TRACKING);
    const logSheet = ss.getSheetByName(CONFIG.SHEET_ATTENDANCE_LOG);
    const archiveSheet = ss.getSheetByName(CONFIG.SHEET_ARCHIVED_WITHDRAWALS);
    
    if (!trackingSheet || !logSheet || !archiveSheet) {
      return { success: false, message: 'Required sheets not found' };
    }
    
    const attendanceData = loadAttendanceData();
    const today = new Date();
    let archivedCount = 0;
    const cutoffDays = CALCULATED.WINDOW_END_DAY + CONFIG.DAYS_AFTER_MISSED_WINDOW;
    
    attendanceData.forEach(function(pair) {
      if (pair.status === CONFIG.STATUS_WITHDRAWN || pair.isNewPair) {
        return;
      }
      
      if (pair.windowInfo && pair.windowInfo.daysSinceLastAttendance > cutoffDays) {
        Logger.log('Auto-archiving: ' + pair.name);
        
        // CALCULATE when window actually closed (withdrawal date)
        const lastAttendanceDate = pair.windowInfo.lastAttendanceDate || today;
        const withdrawalDate = new Date(lastAttendanceDate);
        withdrawalDate.setDate(withdrawalDate.getDate() + CALCULATED.WINDOW_END_DAY + 1);
        
        // Archive to sheet
        const archiveLastRow = archiveSheet.getLastRow();
          archiveSheet.getRange(archiveLastRow + 1, 1, 1, 5).setValues([[
          pair.name,
          pair.quarter,
          CONFIG.AUTO_ARCHIVE_REASON,
          withdrawalDate,  
          today
        ]]);
        
        // Log attendance
        const logLastRow = logSheet.getLastRow();
        logSheet.getRange(logLastRow + 1, 1, 1, 5).setValues([[
          pair.name,
          pair.quarter,
          CONFIG.STATUS_NOT_AT_POST,
          CONFIG.AUTO_ARCHIVE_REASON,
          withdrawalDate   // Use withdrawal date, not today
        ]]);
        
        // Update tracking
        if (pair.memberType === 'Dependent') {
          trackingSheet.getRange(pair.rowNumber, COLUMN_CONFIG.COL_DEPENDENT_AT_POST).setValue(CONFIG.STATUS_NOT_AT_POST);
        } else if (pair.memberType === 'Staff') {
          trackingSheet.getRange(pair.rowNumber, COLUMN_CONFIG.COL_STAFF_AT_POST).setValue(CONFIG.STATUS_NOT_AT_POST);
        }
        
        archivedCount++;
      }
    });
    
    Logger.log('Auto-archived ' + archivedCount + ' pairs');
    return { success: true, message: 'Auto-archived ' + archivedCount + ' pair(s)', archived: archivedCount };
    
    } catch (error) {
    Logger.log('ERROR: ' + error);
    return { success: false, message: 'Error: ' + error.message };
    }
  }
  //=====TEST=====
  function checkDuplicateFunctions() {
    // This will fail if there are duplicate function definitions
    Logger.log('loadAttendanceData exists: ' + typeof loadAttendanceData);
  }

  function testAttendanceLoad() {
    try {
      Logger.log('Starting test...');
      const result = loadAttendanceData();
      Logger.log('Result type: ' + typeof result);
      Logger.log('Is array: ' + Array.isArray(result));
      Logger.log('Length: ' + result.length);
      Logger.log('First item: ' + JSON.stringify(result[0]));
      return result;
    } catch (error) {
      Logger.log('ERROR: ' + error);
      Logger.log('Stack: ' + error.stack);
      return null;
    }
  }

  /**
 * Parses a date string assumed to be in dd/mm/yyyy format.
 * @param {string} dateString - The date string (e.g., "01/12/2025").
 * @returns {Date} A valid Date object or null.
 */
function parseDate_ddMMyyyy(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  // Split the date by / or -
  const parts = dateString.split(/[\/\-]/);

  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10); // month is 1-based
    const year = parseInt(parts[2], 10);

    // Check for valid numbers and create date
    // new Date() uses a 0-based month, so we use month - 1
    if (!isNaN(day) && !isNaN(month) && !isNaN(year) && month >= 1 && month <= 12) {
       // Basic validation for day based on month (doesn't handle leap years perfectly but good enough)
       const daysInMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
       if (day >= 1 && day <= daysInMonth[month - 1]) {
           return new Date(year, month - 1, day);
       }
    }
  }

  // Fallback attempt for other formats (like yyyy-mm-dd or browser default)
  const fallbackDate = new Date(dateString);
  if (!isNaN(fallbackDate.getTime())) {
     Logger.log("Warning: parseDate_ddMMyyyy used fallback for input: " + dateString);
    return fallbackDate;
  }

  Logger.log("Error: parseDate_ddMMyyyy failed to parse input: " + dateString);
  return null; // Invalid date
}
