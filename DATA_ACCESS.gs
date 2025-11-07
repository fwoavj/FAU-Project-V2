// ============================================================================
// DATA ACCESS MODULE - PERFORMANCE OPTIMIZED
// ============================================================================

/**
 * Get complete principal data with all family members in one efficient query
 */
function getPrincipalWithFamily(principalName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    
    if (!sheet) return null;
    
    const principalRow = findPrincipalRowEfficient(sheet, principalName);
    if (!principalRow) return null;
    
    const rowData = sheet.getRange(principalRow, 1, 1, 65).getValues()[0];
    
    const principalData = parsePrincipalRow(rowData);
    const dependents = [];
    const staff = [];
    
    const depName = rowData[SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES.FULL_NAME];
    if (depName && depName.toString().trim() !== '') {
      dependents.push(parseDependentRow(rowData, principalRow));
    }
    
    const staffName = rowData[SYSTEM_CONFIG.COLUMNS.STAFF._INDICES.FULL_NAME];
    if (staffName && staffName.toString().trim() !== '') {
      staff.push(parseStaffRow(rowData, principalRow));
    }
    
    return {
      principal: principalData,
      dependents: dependents,
      staff: staff,
      rowNumber: principalRow
    };
    
  } catch (error) {
    Logger.log('Error getting principal with family: ' + error);
    return null;
  }
}

/**
 * Efficient principal row finder using batch data retrieval
 */
function findPrincipalRowEfficient(sheet, principalName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    return null;
  }
  
  const nameColumn = SYSTEM_CONFIG.COLUMNS.PRINCIPAL.FULL_NAME;
  const names = sheet.getRange(3, nameColumn, lastRow - 2, 1).getValues();
  
  for (let i = 0; i < names.length; i++) {
    if (names[i][0] && names[i][0].toString().trim() === principalName) {
      return i + 3;
    }
  }
  
  return null;
}

/**
 * Parse principal row data into structured object
 */
function parsePrincipalRow(rowData) {
  const indices = SYSTEM_CONFIG.COLUMNS.PRINCIPAL._INDICES;
  
  return {
    postStation: rowData[indices.POST_STATION] || '',
    fullName: rowData[indices.FULL_NAME] || '',
    rank: rowData[indices.RANK] || '',
    designation: rowData[indices.DESIGNATION] || '',
    dateOfBirth: rowData[indices.DATE_OF_BIRTH] || '',
    age: rowData[indices.AGE] || '',
    sex: rowData[indices.SEX] || '',
    assumptionDate: rowData[indices.ASSUMPTION_DATE] || '',
    passportNumber: rowData[indices.PASSPORT_NUMBER] || '',
    passportExpiration: rowData[indices.PASSPORT_EXPIRATION] || '',
    passportUrl: rowData[indices.PASSPORT_URL] || '',
    visaNumber: rowData[indices.VISA_NUMBER] || '',
    visaExpiration: rowData[indices.VISA_EXPIRATION] || '',
    diplomaticId: rowData[indices.DIPLOMATIC_ID] || '',
    diplomaticIdExp: rowData[indices.DIPLOMATIC_ID_EXP] || '',
    departureDate: rowData[indices.DEPARTURE_DATE] || '',
    soloParent: rowData[indices.SOLO_PARENT] || 'No',
    soloParentUrl: rowData[indices.SOLO_PARENT_URL] || '',
    extended: rowData[indices.EXTENDED] || 'No',
    newDepartureDate: rowData[indices.CURRENT_DEPARTURE_DATE] || '',
    extensionDetails: rowData[indices.EXTENSION_DETAILS] || ''
  };
}

/**
 * Parse dependent row data into structured object
 */
function parseDependentRow(rowData, rowNumber) {
  const indices = SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES;
  
  return {
    fullName: rowData[indices.FULL_NAME] || '',
    relationship: rowData[indices.RELATIONSHIP] || '',
    sex: rowData[indices.SEX] || '',
    dateOfBirth: rowData[indices.DATE_OF_BIRTH] || '',
    age: rowData[indices.AGE] || '',
    turns18Date: rowData[indices.TURNS_18_DATE] || '',
    atPost: rowData[indices.AT_POST] || '',
    noticeOfArrival: rowData[indices.NOTICE_OF_ARRIVAL] || '',
    familyAllowance: rowData[indices.FAMILY_ALLOWANCE] || 'No',
    passportNumber: rowData[indices.PASSPORT_NUMBER] || '',
    passportExpiration: rowData[indices.PASSPORT_EXPIRATION] || '',
    passportUrl: rowData[indices.PASSPORT_URL] || '',
    visaNumber: rowData[indices.VISA_NUMBER] || '',
    visaExpiration: rowData[indices.VISA_EXPIRATION] || '',
    diplomaticId: rowData[indices.DIPLOMATIC_ID] || '',
    diplomaticIdExp: rowData[indices.DIPLOMATIC_ID_EXP] || '',
    departureDate: rowData[indices.DEPARTURE_DATE] || '',
    pwdStatus: rowData[indices.PWD_STATUS] || 'No',
    pwdUrl: rowData[indices.PWD_URL] || '',
    approvalFaxUrl: rowData[indices.APPROVAL_FAX_URL] || '',
    extended: rowData[indices.EXTENDED] || 'No',
    newDepartureDate: rowData[indices.CURRENT_DEPARTURE_DATE] || '',
    extensionDetails: rowData[indices.EXTENSION_DETAILS] || '',
    rowNumber: rowNumber
  };
}

/**
 * Parse staff row data into structured object
 */
function parseStaffRow(rowData, rowNumber) {
  const indices = SYSTEM_CONFIG.COLUMNS.STAFF._INDICES;
  
  return {
    fullName: rowData[indices.FULL_NAME] || '',
    sex: rowData[indices.SEX] || '',
    dateOfBirth: rowData[indices.DATE_OF_BIRTH] || '',
    age: rowData[indices.AGE] || '',
    atPost: rowData[indices.AT_POST] || '',
    arrivalDate: rowData[indices.ARRIVAL_DATE] || '',
    passportNumber: rowData[indices.PASSPORT_NUMBER] || '',
    passportExpiration: rowData[indices.PASSPORT_EXPIRATION] || '',
    passportUrl: rowData[indices.PASSPORT_URL] || '',
    visaNumber: rowData[indices.VISA_NUMBER] || '',
    visaExpiration: rowData[indices.VISA_EXPIRATION] || '',
    diplomaticId: rowData[indices.DIPLOMATIC_ID] || '',
    diplomaticIdExp: rowData[indices.DIPLOMATIC_ID_EXP] || '',
    departureDate: rowData[indices.DEPARTURE_DATE] || '',
    pwdStatus: rowData[indices.PWD_STATUS] || 'No',
    pwdUrl: rowData[indices.PWD_URL] || '',
    emergencyContact: rowData[indices.EMERGENCY_CONTACT] || '',
    extended: rowData[indices.EXTENDED] || 'No',
    newDepartureDate: rowData[indices.CURRENT_DEPARTURE_DATE] || '',
    extensionDetails: rowData[indices.EXTENSION_DETAILS] || '',
    rowNumber: rowNumber
  };
}

/**
 * Batch update multiple cells efficiently
 */
function batchUpdateCells(sheetName, updates) {
  if (!updates || updates.length === 0) {
    return {success: true, message: 'No updates to perform'};
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {success: false, message: 'Sheet not found: ' + sheetName};
    }
    
    const rowUpdates = {};
    
    updates.forEach(function(update) {
      if (!rowUpdates[update.row]) {
        rowUpdates[update.row] = [];
      }
      rowUpdates[update.row].push(update);
    });
    
    Object.keys(rowUpdates).forEach(function(row) {
      const rowNum = parseInt(row);
      const cellUpdates = rowUpdates[row];
      
      const cols = cellUpdates.map(u => u.col);
      const minCol = Math.min(...cols);
      const maxCol = Math.max(...cols);
      const numCols = maxCol - minCol + 1;
      
      const currentData = sheet.getRange(rowNum, minCol, 1, numCols).getValues()[0];
      
      cellUpdates.forEach(function(update) {
        const colIndex = update.col - minCol;
        currentData[colIndex] = update.value;
      });
      
      sheet.getRange(rowNum, minCol, 1, numCols).setValues([currentData]);
    });
    
    return {success: true, message: 'Updated ' + updates.length + ' cells'};
    
  } catch (error) {
    Logger.log('Error in batch update: ' + error);
    return {success: false, message: 'Error: ' + error.message};
  }
}

/**
 * Get all personnel data efficiently
 */
function getAllPersonnelData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    
    const allData = sheet.getRange(3, 1, lastRow - 2, 65).getValues();
    const personnel = [];
    
    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      const principalName = row[SYSTEM_CONFIG.COLUMNS.PRINCIPAL._INDICES.FULL_NAME];
      
      if (principalName && principalName.toString().trim() !== '') {
        const dependents = [];
        const staff = [];
        
        const depName = row[SYSTEM_CONFIG.COLUMNS.DEPENDENT._INDICES.FULL_NAME];
        if (depName && depName.toString().trim() !== '') {
          dependents.push(parseDependentRow(row, i + 3));
        }
        
        const staffName = row[SYSTEM_CONFIG.COLUMNS.STAFF._INDICES.FULL_NAME];
        if (staffName && staffName.toString().trim() !== '') {
          staff.push(parseStaffRow(row, i + 3));
        }

        personnel.push({
          principal: parsePrincipalRow(row),
          dependents: dependents,
          staff: staff,
          rowNumber: i + 3
        });
      }
    }
    
    return personnel;
    
  } catch (error) {
    Logger.log('Error getting all personnel: ' + error);
    return [];
  }
}

/**
 * Get expiring documents efficiently
 */
function getExpiringDocuments(warningDays) {
  try {
    const allPersonnel = getAllPersonnelData();
    const today = new Date();
    const warningDate = new Date();
    warningDate.setDate(warningDate.getDate() + warningDays);
    
    const expiring = {
      passports: [],
      visas: [],
      diplomaticIds: []
    };
    
    allPersonnel.forEach(function(person) {
      checkDocumentExpiration(person.principal, 'Principal', person.principal.fullName, expiring, today, warningDate);
      
      person.dependents.forEach(function(dep) {
        checkDocumentExpiration(dep, 'Dependent', dep.fullName, expiring, today, warningDate);
      });
      
      person.staff.forEach(function(staff) {
        checkDocumentExpiration(staff, 'Staff', staff.fullName, expiring, today, warningDate);
      });
    });
    
    return expiring;
    
  } catch (error) {
    Logger.log('Error getting expiring documents: ' + error);
    return {passports: [], visas: [], diplomaticIds: []};
  }
}

/**
 * Helper to check document expiration
 */
function checkDocumentExpiration(person, type, name, expiring, today, warningDate) {
  if (person.passportExpiration) {
    const passportExp = new Date(person.passportExpiration);
    if (passportExp > today && passportExp <= warningDate) {
      expiring.passports.push({
        type: type,
        name: name,
        expirationDate: passportExp,
        daysUntilExpiration: Math.floor((passportExp - today) / (1000 * 60 * 60 * 24))
      });
    }
  }
  
  if (person.visaExpiration) {
    const visaExp = new Date(person.visaExpiration);
    if (visaExp > today && visaExp <= warningDate) {
      expiring.visas.push({
        type: type,
        name: name,
        expirationDate: visaExp,
        daysUntilExpiration: Math.floor((visaExp - today) / (1000 * 60 * 60 * 24))
      });
    }
  }
  
  if (person.diplomaticIdExp) {
    const dipIdExp = new Date(person.diplomaticIdExp);
    if (dipIdExp > today && dipIdExp <= warningDate) {
      expiring.diplomaticIds.push({
        type: type,
        name: name,
        expirationDate: dipIdExp,
        daysUntilExpiration: Math.floor((dipIdExp - today) / (1000 * 60 * 60 * 24))
      });
    }
  }
}

/**
 * Get dependents turning 18 soon
 */
function getDependentsTurning18(alertDays) {
  try {
    const allPersonnel = getAllPersonnelData();
    const today = new Date();
    const alertDate = new Date();
    alertDate.setDate(alertDate.getDate() + Math.max(...alertDays));
    
    const turning18 = [];
    
    allPersonnel.forEach(function(person) {
      person.dependents.forEach(function(dep) {
        if (dep.dateOfBirth) {
          const birthDate = new Date(dep.dateOfBirth);
          const turns18Date = new Date(birthDate);
          turns18Date.setFullYear(turns18Date.getFullYear() + 18);
          
          const daysUntil18 = Math.floor((turns18Date - today) / (1000 * 60 * 60 * 24));
          
          if (daysUntil18 >= 0 && daysUntil18 <= Math.max(...alertDays)) {
            turning18.push({
              principalName: person.principal.fullName,
              dependentName: dep.fullName,
              relationship: dep.relationship,
              turns18Date: turns18Date,
              daysUntil18: daysUntil18,
              shouldAlert: alertDays.includes(daysUntil18)
            });
          }
        }
      });
    });
    
    turning18.sort((a, b) => a.daysUntil18 - b.daysUntil18);
    
    return turning18;
    
  } catch (error) {
    Logger.log('Error getting dependents turning 18: ' + error);
    return [];
  }
}

/**
 * Cache wrapper for expensive operations
 */
const DataCache = {
  cache: {},
  
  get: function(key) {
    const cached = this.cache[key];
    if (!cached) {
      return null;
    }
    
    const now = new Date().getTime();
    if (now - cached.timestamp > SYSTEM_CONFIG.PERFORMANCE.CACHE_DURATION * 1000) {
      delete this.cache[key];
      return null;
    }
    
    return cached.data;
  },
  
  set: function(key, data) {
    this.cache[key] = {
      data: data,
      timestamp: new Date().getTime()
    };
  },
  
  clear: function(key) {
    if (key) {
      delete this.cache[key];
    } else {
      this.cache = {};
    }
  }
};

/**
 * Cached version of getAllPersonnelData
 */
function getAllPersonnelDataCached() {
  const cacheKey = 'all_personnel';
  let data = DataCache.get(cacheKey);
  
  if (!data) {
    data = getAllPersonnelData();
    DataCache.set(cacheKey, data);
  }
  
  return data;
}

/**
 * Clear cache when data changes
 */
function clearDataCache() {
  DataCache.clear();
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
