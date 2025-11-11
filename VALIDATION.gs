// ============================================================================
// DATA VALIDATION MODULE
// ============================================================================

/**
 * Main validation function - validates complete form data
 */
function validateFormData(formData) {
  const errors = [];
  const warnings = [];
  
  try {
    if (formData.principal) {
      const principalValidation = validatePrincipal(formData.principal);
      errors.push(...principalValidation.errors);
      warnings.push(...principalValidation.warnings);
    } else {
      errors.push({field: 'principal', message: 'Principal data is required'});
    }
    
    if (formData.dependents && formData.dependents.length > 0) {
      formData.dependents.forEach((dependent, index) => {
        const depValidation = validateDependent(dependent, formData.principal);
        const prefix = `Dependent ${index + 1}`;
        errors.push(...depValidation.errors.map(e => ({...e, field: `${prefix}: ${e.field}`})));
        warnings.push(...depValidation.warnings.map(w => ({...w, field: `${prefix}: ${w.field}`})));
      });
    }
    
    if (formData.privateStaff && formData.privateStaff.length > 0) {
      formData.privateStaff.forEach((staff, index) => {
        const staffValidation = validateStaff(staff, formData.principal);
        const prefix = `Staff ${index + 1}`;
        errors.push(...staffValidation.errors.map(e => ({...e, field: `${prefix}: ${e.field}`})));
        warnings.push(...staffValidation.warnings.map(w => ({...w, field: `${prefix}: ${w.field}`})));
      });
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors,
      warnings: warnings
    };
    
  } catch (error) {
    Logger.log('Validation error: ' + error);
    return {
      isValid: false,
      errors: [{field: 'system', message: 'Validation system error: ' + error.message}],
      warnings: []
    };
  }
}

/**
 * Validate principal data
 */
function validatePrincipal(principal) {
  const errors = [];
  const warnings = [];
  
  if (!validateRequiredField(principal.fullName, 'Full Name')) {
    errors.push({field: 'fullName', message: 'Full name is required'});
  } else {
    const nameValidation = validateName(principal.fullName);
    if (!nameValidation.isValid) {
      errors.push({field: 'fullName', message: nameValidation.message});
    }
  }
  
  if (!validateRequiredField(principal.rank, 'Rank')) {
    errors.push({field: 'rank', message: 'Rank is required'});
  }
  
  if (!validateRequiredField(principal.post, 'Post/Station')) {
    errors.push({field: 'post', message: 'Post/Station is required'});
  }
  
  if (!validateRequiredField(principal.dateOfBirth, 'Date of Birth')) {
    errors.push({field: 'dateOfBirth', message: 'Date of birth is required'});
  } else {
    const dobValidation = validateDateOfBirth(
      principal.dateOfBirth,
      SYSTEM_CONFIG.VALIDATION.MIN_AGE_PRINCIPAL,
      SYSTEM_CONFIG.VALIDATION.MAX_AGE_PRINCIPAL
    );
    if (!dobValidation.isValid) {
      errors.push({field: 'dateOfBirth', message: dobValidation.message});
    }
  }
  
  if (principal.arrivalDate && principal.departureDate) {
    const dateLogic = validateDateLogic(principal.arrivalDate, principal.departureDate, 'Arrival', 'Departure');
    if (!dateLogic.isValid) {
      errors.push({field: 'dates', message: dateLogic.message});
    }
  }
  
  if (principal.passportNumber && principal.passportExp) {
    const passportValidation = validateDocument(
      principal.passportExp,
      'Passport',
      SYSTEM_CONFIG.VALIDATION.PASSPORT_WARNING_DAYS
    );
    if (!passportValidation.isValid) {
      warnings.push({field: 'passportExp', message: passportValidation.message});
    }
  }
  
  if (principal.visaNumber && principal.visaExp) {
    const visaValidation = validateDocument(
      principal.visaExp,
      'Visa',
      SYSTEM_CONFIG.VALIDATION.VISA_WARNING_DAYS
    );
    if (!visaValidation.isValid) {
      warnings.push({field: 'visaExp', message: visaValidation.message});
    }
    
    if (principal.departureDate) {
      const visaDepartureCheck = validateDateLogic(principal.visaExp, principal.departureDate, 'Visa expiration', 'Departure date');
      if (!visaDepartureCheck.isValid) {
        warnings.push({field: 'visaExp', message: 'Visa expires before departure date'});
      }
    }
  }
  
  if (principal.dipId && principal.dipIdExp) {
    const dipIdValidation = validateDocument(
      principal.dipIdExp,
      'Diplomatic ID',
      SYSTEM_CONFIG.VALIDATION.DIPLOMATIC_ID_WARNING_DAYS
    );
    if (!dipIdValidation.isValid) {
      warnings.push({field: 'dipIdExp', message: dipIdValidation.message});
    }
  }
  
  if (principal.extended === 'Yes') {
    if (!principal.newDeparture) {
      errors.push({field: 'newDeparture', message: 'Current departure date required when extended'});
    }
    if (!principal.extensionDetails || principal.extensionDetails.trim() === '') {
      warnings.push({field: 'extensionDetails', message: 'Extension details recommended'});
    }
  }
  
  if (principal.soloParent === 'Yes' && !principal.soloParentFile && !principal.soloParentUrl) {
    warnings.push({field: 'soloParentUrl', message: 'Solo parent documentation recommended'});
  }
  
  return {errors, warnings};
}

/**
 * Validate dependent data
 */
function validateDependent(dependent, principal) {
  const errors = [];
  const warnings = [];
  
  if (!validateRequiredField(dependent.fullName, 'Full Name')) {
    errors.push({field: 'fullName', message: 'Full name is required'});
  } else {
    const nameValidation = validateName(dependent.fullName);
    if (!nameValidation.isValid) {
      errors.push({field: 'fullName', message: nameValidation.message});
    }
  }
  
  if (!validateRequiredField(dependent.relationship, 'Relationship')) {
    errors.push({field: 'relationship', message: 'Relationship is required'});
  }
  
  if (!validateRequiredField(dependent.dateOfBirth, 'Date of Birth')) {
    errors.push({field: 'dateOfBirth', message: 'Date of birth is required'});
  } else {
    const dobValidation = validateDateOfBirth(
      dependent.dateOfBirth,
      SYSTEM_CONFIG.VALIDATION.MIN_AGE_DEPENDENT,
      SYSTEM_CONFIG.VALIDATION.MAX_AGE_DEPENDENT
    );
    if (!dobValidation.isValid) {
      errors.push({field: 'dateOfBirth', message: dobValidation.message});
    }
    
    const turns18Validation = checkTurning18(dependent.dateOfBirth);
    if (turns18Validation.isTurning18Soon) {
      warnings.push({
        field: 'age',
        message: `⚠️ CRITICAL: Will turn 18 in ${turns18Validation.daysUntil18} days (${turns18Validation.turns18Date}). May lose dependent status!`,
        priority: 'HIGH'
      });
    }
    if (turns18Validation.isOver18) {
      errors.push({
        field: 'age',
        message: `Dependent is over 18 years old (turned 18 on ${turns18Validation.turns18Date}). Cannot be registered as dependent.`
      });
    }
  }
  
  if (dependent.arrivalDate && dependent.departureDate) {
    const dateLogic = validateDateLogic(dependent.arrivalDate, dependent.departureDate, 'Arrival', 'Departure');
    if (!dateLogic.isValid) {
      errors.push({field: 'dates', message: dateLogic.message});
    }
  }
  
  if (dependent.arrivalDate && principal && principal.arrivalDate) {
    const arrivalCheck = validateDateLogic(principal.arrivalDate, dependent.arrivalDate, 'Principal arrival', 'Dependent arrival');
    if (!arrivalCheck.isValid) {
      warnings.push({field: 'arrivalDate', message: 'Dependent arrival date is before principal arrival date'});
    }
  }
  
  if (dependent.passportNumber && dependent.passportExp) {
    const passportValidation = validateDocument(
      dependent.passportExp,
      'Passport',
      SYSTEM_CONFIG.VALIDATION.PASSPORT_WARNING_DAYS
    );
    if (!passportValidation.isValid) {
      warnings.push({field: 'passportExp', message: passportValidation.message});
    }
  }
  
  if (dependent.visaNumber && dependent.visaExp) {
    const visaValidation = validateDocument(
      dependent.visaExp,
      'Visa',
      SYSTEM_CONFIG.VALIDATION.VISA_WARNING_DAYS
    );
    if (!visaValidation.isValid) {
      warnings.push({field: 'visaExp', message: visaValidation.message});
    }
    
    if (dependent.departureDate) {
      const visaDepartureCheck = validateDateLogic(dependent.visaExp, dependent.departureDate, 'Visa expiration', 'Departure date');
      if (!visaDepartureCheck.isValid) {
        warnings.push({field: 'visaExp', message: 'Visa expires before departure date'});
      }
    }
  }
  
  if (dependent.pwdStatus === 'Yes') {
    if (!dependent.pwdFile && !dependent.pwdUrl) {
      warnings.push({field: 'pwdUrl', message: 'PWD documentation recommended for PWD dependent'});
    }
    if (!dependent.approvalFaxFile && !dependent.approvalFaxUrl) {
      warnings.push({field: 'approvalFax', message: 'Approval documentation recommended for PWD dependent'});
    }
  }
  
  if (dependent.extended === 'Yes') {
    if (!dependent.newDeparture) {
      errors.push({field: 'newDeparture', message: 'Current departure date required when extended'});
    }
  }
  
  return {errors, warnings};
}

/**
 * Validate staff data
 */
function validateStaff(staff, principal) {
  const errors = [];
  const warnings = [];
  
  if (!validateRequiredField(staff.fullName, 'Full Name')) {
    errors.push({field: 'fullName', message: 'Full name is required'});
  } else {
    const nameValidation = validateName(staff.fullName);
    if (!nameValidation.isValid) {
      errors.push({field: 'fullName', message: nameValidation.message});
    }
  }
  
  if (!validateRequiredField(staff.dateOfBirth, 'Date of Birth')) {
    errors.push({field: 'dateOfBirth', message: 'Date of birth is required'});
  } else {
    const dobValidation = validateDateOfBirth(
      staff.dateOfBirth,
      SYSTEM_CONFIG.VALIDATION.MIN_AGE_STAFF,
      SYSTEM_CONFIG.VALIDATION.MAX_AGE_STAFF
    );
    if (!dobValidation.isValid) {
      errors.push({field: 'dateOfBirth', message: dobValidation.message});
    }
  }
  
  if (staff.arrivalDate && staff.departureDate) {
    const dateLogic = validateDateLogic(staff.arrivalDate, staff.departureDate, 'Arrival', 'Departure');
    if (!dateLogic.isValid) {
      errors.push({field: 'dates', message: dateLogic.message});
    }
  }
  
  if (staff.passportNumber && staff.passportExp) {
    const passportValidation = validateDocument(
      staff.passportExp,
      'Passport',
      SYSTEM_CONFIG.VALIDATION.PASSPORT_WARNING_DAYS
    );
    if (!passportValidation.isValid) {
      warnings.push({field: 'passportExp', message: passportValidation.message});
    }
  }
  
  if (staff.visaNumber && staff.visaExp) {
    const visaValidation = validateDocument(
      staff.visaExp,
      'Visa',
      SYSTEM_CONFIG.VALIDATION.VISA_WARNING_DAYS
    );
    if (!visaValidation.isValid) {
      warnings.push({field: 'visaExp', message: visaValidation.message});
    }
  }
  
  if (!staff.emergencyContact || staff.emergencyContact.trim() === '') {
    warnings.push({field: 'emergencyContact', message: 'Emergency contact information recommended'});
  }
  
  if (staff.pwdStatus === 'Yes' && !staff.pwdFile && !staff.pwdUrl) {
    warnings.push({field: 'pwdUrl', message: 'PWD documentation recommended for PWD staff'});
  }
  
  return {errors, warnings};
}

// ============================================================================
// HELPER VALIDATION FUNCTIONS
// ============================================================================

function validateRequiredField(value, fieldName) {
  return value !== null && value !== undefined && value.toString().trim() !== '';
}

function validateName(name) {
  if (!name || typeof name !== 'string') {
    return {isValid: false, message: 'Name must be a text value'};
  }
  
  const trimmed = name.trim();
  
  if (trimmed.length < SYSTEM_CONFIG.VALIDATION.MIN_NAME_LENGTH) {
    return {isValid: false, message: `Name must be at least ${SYSTEM_CONFIG.VALIDATION.MIN_NAME_LENGTH} characters`};
  }
  
  if (trimmed.length > SYSTEM_CONFIG.VALIDATION.MAX_NAME_LENGTH) {
    return {isValid: false, message: `Name must be less than ${SYSTEM_CONFIG.VALIDATION.MAX_NAME_LENGTH} characters`};
  }
  
  const namePattern = /^[a-zA-Z\s\-'.]+$/;
  if (!namePattern.test(trimmed)) {
    return {isValid: false, message: 'Name contains invalid characters'};
  }
  
  return {isValid: true};
}

function validateDateOfBirth(dob, minAge, maxAge) {
  if (!dob) {
    return {isValid: false, message: 'Date of birth is required'};
  }
  
  let birthDate;
  try {
    birthDate = new Date(dob);
    if (isNaN(birthDate.getTime())) {
      return {isValid: false, message: 'Invalid date format'};
    }
  } catch (error) {
    return {isValid: false, message: 'Invalid date format'};
  }
  
  const today = new Date();
  const age = calculateAge(birthDate);
  
  const maxPastDate = new Date();
  maxPastDate.setFullYear(maxPastDate.getFullYear() - SYSTEM_CONFIG.VALIDATION.MAX_PAST_YEARS);
  if (birthDate < maxPastDate) {
    return {isValid: false, message: 'Date of birth is too far in the past'};
  }
  
  if (birthDate > today) {
    return {isValid: false, message: 'Date of birth cannot be in the future'};
  }
  
  if (age < minAge) {
    return {isValid: false, message: `Age must be at least ${minAge} years old (currently ${age})`};
  }
  
  if (age > maxAge) {
    return {isValid: false, message: `Age must be at most ${maxAge} years old (currently ${age})`};
  }
  
  return {isValid: true, age: age};
}

function validateDateLogic(earlierDate, laterDate, earlierLabel, laterLabel) {
  if (!earlierDate || !laterDate) {
    return {isValid: true};
  }
  
  try {
    const date1 = new Date(earlierDate);
    const date2 = new Date(laterDate);
    
    if (isNaN(date1.getTime()) || isNaN(date2.getTime())) {
      return {isValid: false, message: 'Invalid date format'};
    }
    
    if (date1 >= date2) {
      return {isValid: false, message: `${earlierLabel} date must be before ${laterLabel} date`};
    }
    
    return {isValid: true};
  } catch (error) {
    return {isValid: false, message: 'Error validating dates'};
  }
}

function validateDocument(expirationDate, documentType, warningDays) {
  if (!expirationDate) {
    return {isValid: true};
  }
  
  try {
    const expDate = new Date(expirationDate);
    if (isNaN(expDate.getTime())) {
      return {isValid: false, message: `Invalid ${documentType} expiration date format`};
    }
    
    const today = new Date();
    const daysUntilExpiration = Math.floor((expDate - today) / (1000 * 60 * 60 * 24));
    
    if (daysUntilExpiration < 0) {
      return {isValid: false, message: `${documentType} has expired`};
    }
    
    if (daysUntilExpiration <= warningDays) {
      return {isValid: false, message: `${documentType} expires in ${daysUntilExpiration} days`};
    }
    
    return {isValid: true};
  } catch (error) {
    return {isValid: false, message: `Error validating ${documentType} expiration`};
  }
}

function checkTurning18(dateOfBirth) {
  if (!dateOfBirth) {
    return {isTurning18Soon: false, isOver18: false};
  }
  
  try {
    const birthDate = new Date(dateOfBirth);
    if (isNaN(birthDate.getTime())) {
      return {isTurning18Soon: false, isOver18: false};
    }
    
    const today = new Date();
    const age = calculateAge(birthDate);
    
    const turns18Date = new Date(birthDate);
    turns18Date.setFullYear(turns18Date.getFullYear() + 18);
    
    const daysUntil18 = Math.floor((turns18Date - today) / (1000 * 60 * 60 * 24));
    
    if (age >= 18) {
      return {
        isTurning18Soon: false,
        isOver18: true,
        turns18Date: Utilities.formatDate(turns18Date, Session.getScriptTimeZone(), 'MMM dd, yyyy'),
        daysUntil18: daysUntil18
      };
    }
    
    const alertDays = SYSTEM_CONFIG.ALERTS ? SYSTEM_CONFIG.ALERTS.AGE_18_ALERTS.DAYS_BEFORE : [90, 60, 30, 14, 7];
    const maxAlertDays = Math.max(...alertDays);
    
    if (daysUntil18 <= maxAlertDays && daysUntil18 >= 0) {
      return {
        isTurning18Soon: true,
        isOver18: false,
        turns18Date: Utilities.formatDate(turns18Date, Session.getScriptTimeZone(), 'MMM dd, yyyy'),
        daysUntil18: daysUntil18
      };
    }
    
    return {
      isTurning18Soon: false,
      isOver18: false,
      turns18Date: Utilities.formatDate(turns18Date, Session.getScriptTimeZone(), 'MMM dd, yyyy'),
      daysUntil18: daysUntil18
    };
    
  } catch (error) {
    Logger.log('Error checking age 18: ' + error);
    return {isTurning18Soon: false, isOver18: false};
  }
}

function calculateAge(birthDate) {
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();
  const monthDiff = today.getMonth() - birthDate.getMonth();
  
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
    age--;
  }
  
  return age;
}

/**
 * Check for duplicate personnel
 */
function checkDuplicatePrincipal(fullName, dateOfBirth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.PERSONNEL_TRACKING);
    
    if (!sheet) {
      return {isDuplicate: false};
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      return {isDuplicate: false};
    }
    
    const data = sheet.getRange(3, SYSTEM_CONFIG.COLUMNS.PRINCIPAL.FULL_NAME, lastRow - 2, 2).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const existingName = data[i][0] ? data[i][0].toString().trim() : '';
      const existingDOB = data[i][1];
      
      if (!existingName) continue;
      
      if (existingName.toLowerCase() === fullName.toLowerCase()) {
        if (existingDOB && dateOfBirth) {
          const dobMatch = new Date(existingDOB).getTime() === new Date(dateOfBirth).getTime();
          if (dobMatch) {
            return {
              isDuplicate: true,
              matchType: 'exact',
              existingName: existingName,
              row: i + 3,
              message: `Exact match found: "${existingName}" with same date of birth`
            };
          }
        }
        return {
          isDuplicate: true,
          matchType: 'name',
          existingName: existingName,
          row: i + 3,
          message: `Principal with name "${existingName}" already exists`
        };
      }
      
      const similarity = calculateStringSimilarity(fullName.toLowerCase(), existingName.toLowerCase());
      if (similarity >= SYSTEM_CONFIG.VALIDATION.FUZZY_MATCH_THRESHOLD) {
        return {
          isDuplicate: true,
          matchType: 'similar',
          existingName: existingName,
          similarity: Math.round(similarity * 100),
          row: i + 3,
          message: `Similar name found: "${existingName}" (${Math.round(similarity * 100)}% match)`
        };
      }
    }
    
    return {isDuplicate: false};
    
  } catch (error) {
    Logger.log('Error checking duplicate: ' + error);
    return {isDuplicate: false, error: error.message};
  }
}

/**
 * Calculate string similarity using Levenshtein distance
 */
function calculateStringSimilarity(str1, str2) {
  const longer = str1.length > str2.length ? str1 : str2;
  const shorter = str1.length > str2.length ? str2 : str1;
  
  if (longer.length === 0) {
    return 1.0;
  }
  
  const editDistance = levenshteinDistance(longer, shorter);
  return (longer.length - editDistance) / longer.length;
}

/**
 * Calculate Levenshtein distance between two strings
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];
  
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }
  
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  
  return matrix[str2.length][str1.length];
}
