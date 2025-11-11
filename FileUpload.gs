// ============================================================================
// FILE UPLOAD MODULE - UPDATED FOR SUBFOLDERS
// ============================================================================

/**
 * Save a base64 file to a specific Google Drive folder
 * @param {string} fileName - The original name of the file
 * @param {string} base64Data - The base64-encoded file content
 * @param {string} mimeType - The MIME type of the file
 * @param {string} personName - The name of the person (for file naming)
 * @param {string} fileType - The type of file (e.g., "Passport", "PWD")
 * @param {string} targetFolderId - The specific Google Drive Folder ID to save to
 * @returns {string} The URL of the saved file
 */
function saveFileToDrive(fileName, base64Data, mimeType, personName, fileType, targetFolderId) {
  try {
    // Default to the main folder ID if no specific one is provided
    const folderId = targetFolderId || SYSTEM_CONFIG.DRIVE.FOLDER_ID;
    const folder = DriveApp.getFolderById(folderId);
    
    Logger.log(`Saving file to Drive: ${fileName} in folder ID: ${folderId}`);

    const base64Content = base64Data.includes(',') 
      ? base64Data.split(',')[1] 
      : base64Data;
    
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Content),
      mimeType,
      fileName
    );
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const finalFileName = personName.replace(/[^a-zA-Z0-9]/g, '_') + '_' + fileType + '_' + timestamp + '_' + fileName;
    
    const file = folder.createFile(blob);
    file.setName(finalFileName);
    
    file.setSharing(
      SYSTEM_CONFIG.SECURITY.FILE_SHARING.ACCESS,
      SYSTEM_CONFIG.SECURITY.FILE_SHARING.PERMISSION
    );
    
    const fileUrl = file.getUrl();
    Logger.log('File saved successfully: ' + fileUrl);
    return fileUrl;
    
  } catch (error) {
    Logger.log(`Error saving file to Drive (Folder ID: ${targetFolderId}): ${error}`);
    throw new Error('Failed to save file: ' + error.message);
  }
}

/**
 * Process principal passport file
 */
function processPrincipalPassportFile(principal) {
  if (principal.principalPassportFile && principal.principalPassportFile.data) {
    try {
      const fileUrl = saveFileToDrive(
        principal.principalPassportFile.name,
        principal.principalPassportFile.data,
        principal.principalPassportFile.mimeType,
        principal.fullName,
        'Passport',
        SYSTEM_CONFIG.DRIVE.SUBFOLDERS.PRINCIPAL_PASSPORTS // <-- Use specific folder
      );
      principal.principalPassportUrl = fileUrl;
      Logger.log('Principal passport saved: ' + fileUrl);
    } catch (error) {
      Logger.log('Error processing principal passport: ' + error);
      principal.principalPassportUrl = 'ERROR: ' + error.message;
    }
  }
  return principal;
}

/**
 * Process solo parent file
 */
function processSoloParentFile(principal) {
  if (principal.soloParentFile && principal.soloParentFile.data) {
    try {
      const fileUrl = saveFileToDrive(
        principal.soloParentFile.name,
        principal.soloParentFile.data,
        principal.soloParentFile.mimeType,
        principal.fullName,
        'SoloParent',
        SYSTEM_CONFIG.DRIVE.SUBFOLDERS.SOLO_PARENT // <-- Use specific folder
      );
      principal.soloParentUrl = fileUrl;
      Logger.log('Solo parent document saved: ' + fileUrl);
    } catch (error) {
      Logger.log('Error processing solo parent document: ' + error);
      principal.soloParentUrl = 'ERROR: ' + error.message;
    }
  }
  return principal;
}

/**
 * Process dependent files (passport, PWD, approval fax)
 */
function processDependentFiles(dependents) {
  if (!dependents || dependents.length === 0) return dependents;
  
  for (let i = 0; i < dependents.length; i++) {
    const dep = dependents[i];
    const dependentName = (dep.firstName + '_' + dep.lastName).replace(/[^a-zA-Z0-9]/g, '_');
    
    // Process passport file
    if (dep.passportFile && dep.passportFile.data) {
      try {
        const fileUrl = saveFileToDrive(
          dep.passportFile.name,
          dep.passportFile.data,
          dep.passportFile.mimeType,
          dependentName,
          'Passport',
          SYSTEM_CONFIG.DRIVE.SUBFOLDERS.DEPENDENT_PASSPORTS // <-- Use specific folder
        );
        dep.passportUrl = fileUrl;
        Logger.log('Dependent passport saved for ' + dependentName + ': ' + fileUrl);
      } catch (error) {
        Logger.log('Error processing dependent passport for ' + dependentName + ': ' + error);
        dep.passportUrl = 'ERROR: ' + error.message;
      }
    }
    
    // Process PWD file
    if (dep.pwdFile && dep.pwdFile.data) {
      try {
        const fileUrl = saveFileToDrive(
          dep.pwdFile.name,
          dep.pwdFile.data,
          dep.pwdFile.mimeType,
          dependentName,
          'PWD',
          SYSTEM_CONFIG.DRIVE.SUBFOLDERS.PWD // <-- Use specific folder
        );
        dep.pwdUrl = fileUrl;
        Logger.log('Dependent PWD saved for ' + dependentName + ': ' + fileUrl);
      } catch (error) {
        Logger.log('Error processing dependent PWD for ' + dependentName + ': ' + error);
        dep.pwdUrl = 'ERROR: ' + error.message;
      }
    }
    
    // Process approval fax file
    if (dep.approvalFaxFile && dep.approvalFaxFile.data) {
      try {
        const fileUrl = saveFileToDrive(
          dep.approvalFaxFile.name,
          dep.approvalFaxFile.data,
          dep.approvalFaxFile.mimeType,
          dependentName,
          'ApprovalFax',
          SYSTEM_CONFIG.DRIVE.SUBFOLDERS.APPROVAL_FAX // <-- Use specific folder
        );
        dep.approvalFaxUrl = fileUrl;
        Logger.log('Approval fax saved for ' + dependentName + ': ' + fileUrl);
      } catch (error) {
        Logger.log('Error processing approval fax for ' + dependentName + ': ' + error);
        dep.approvalFaxUrl = 'ERROR: ' + error.message;
      }
    }
  }
  
  return dependents;
}

/**
 * Process staff files (passport, PWD)
 */
function processStaffFiles(staff) {
  if (!staff || staff.length === 0) return staff;
  
  for (let i = 0; i < staff.length; i++) {
    const staffMember = staff[i];
    const staffName = (staffMember.firstName + '_' + staffMember.lastName).replace(/[^a-zA-Z0-9]/g, '_');
    
    // Process passport file
    if (staffMember.passportFile && staffMember.passportFile.data) {
      try {
        const fileUrl = saveFileToDrive(
          staffMember.passportFile.name,
          staffMember.passportFile.data,
          staffMember.passportFile.mimeType,
          staffName,
          'Passport',
          SYSTEM_CONFIG.DRIVE.SUBFOLDERS.DEPENDENT_PASSPORTS // <-- Use specific folder (assuming staff passports go with dependent passports)
        );
        staffMember.passportUrl = fileUrl;
        Logger.log('Staff passport saved for ' + staffName + ': ' + fileUrl);
      } catch (error) {
        Logger.log('Error processing staff passport for ' + staffName + ': ' + error);
        staffMember.passportUrl = 'ERROR: ' + error.message;
      }
    }
    
    // Process PWD file
    if (staffMember.pwdFile && staffMember.pwdFile.data) {
      try {
        const fileUrl = saveFileToDrive(
          staffMember.pwdFile.name,
          staffMember.pwdFile.data,
          staffMember.pwdFile.mimeType,
          staffName,
          'PWD',
          SYSTEM_CONFIG.DRIVE.SUBFOLDERS.PWD // <-- Use specific folder
        );
        staffMember.pwdUrl = fileUrl;
        Logger.log('Staff PWD saved for ' + staffName + ': ' + fileUrl);
      } catch (error) {
        Logger.log('Error processing staff PWD for ' + staffName + ': ' + error);
        staffMember.pwdUrl = 'ERROR: ' + error.message;
      }
    }
  }
  
  return staff;
}

/**
 * Process all files for form submission
 */
function processAllFiles(formData) {
  try {
    Logger.log('Processing all file uploads...');
    
    // Process principal files
    if (formData.principalPassportFile) {
      formData = processPrincipalPassportFile(formData);
    }
    
    if (formData.soloParent === 'Yes' && formData.soloParentFile) {
      formData = processSoloParentFile(formData);
    }
    
    // Process dependent files
    if (formData.dependents && formData.dependents.length > 0) {
      formData.dependents = processDependentFiles(formData.dependents);
    }
    
    // Process staff files
    if (formData.privateStaff && formData.privateStaff.length > 0) {
      formData.privateStaff = processStaffFiles(formData.privateStaff);
    }
    
    Logger.log('All files processed successfully');
    return formData;
    
  } catch (error) {
    Logger.log('Error processing files: ' + error);
    throw error;
  }
}

/**
 * Delete file from Drive (used for cleanup if needed)
 */
function deleteFileFromDrive(fileUrl) {
  try {
    const fileId = extractFileIdFromUrl(fileUrl);
    if (fileId) {
      const file = DriveApp.getFileById(fileId);
      file.setTrashed(true);
      Logger.log('File deleted: ' + fileUrl);
      return true;
    }
    return false;
  } catch (error) {
    Logger.log('Error deleting file: ' + error);
    return false;
  }
}

/**
 * Extract file ID from Google Drive URL
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  try {
    // This regex looks for a long string of letters, numbers, hyphens, and underscores
    // which is typical for Google Drive IDs
    const match = url.match(/[-\w]{25,}/); 
    return match ? match[0] : null;
  } catch (error) {
    Logger.log('Error extracting file ID: ' + error);
    return null;
  }
}

/**
 * Validate file before upload
 */
function validateFile(fileData) {
  const maxSize = 10 * 1024 * 1024; // 10MB
  const validTypes = ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png'];
  
  if (!fileData || !fileData.data) {
    return {isValid: false, message: 'No file data provided'};
  }
  
  if (fileData.size > maxSize) {
    return {isValid: false, message: 'File size exceeds 10MB limit'};
  }
  
  if (!validTypes.includes(fileData.mimeType)) {
    return {isValid: false, message: 'Invalid file type. Only PDF, JPG, and PNG are allowed'};
  }
  
  return {isValid: true};
}

/**
 * Get file info from Drive URL
 */
function getFileInfo(fileUrl) {
  try {
    const fileId = extractFileIdFromUrl(fileUrl);
    if (!fileId) return null;
    
    const file = DriveApp.getFileById(fileId);
    return {
      name: file.getName(),
      size: file.getSize(),
      mimeType: file.getMimeType(),
      url: file.getUrl(),
      createdDate: file.getDateCreated(),
      lastUpdated: file.getLastUpdated()
    };
  } catch (error) {
    Logger.log('Error getting file info: ' + error);
    return null;
  }
}

/**
 * Get upload statistics
 */
function getUploadStatistics() {
  try {
    const folder = DriveApp.getFolderById(SYSTEM_CONFIG.DRIVE.FOLDER_ID);
    const files = folder.getFiles();
    
    let totalFiles = 0;
    let totalSize = 0;
    const fileTypes = {};
    
    while (files.hasNext()) {
      const file = files.next();
      totalFiles++;
      totalSize += file.getSize();
      
      const mimeType = file.getMimeType();
      fileTypes[mimeType] = (fileTypes[mimeType] || 0) + 1;
    }
    
    return {
      totalFiles: totalFiles,
      totalSize: totalSize,
      totalSizeMB: (totalSize / (1024 * 1024)).toFixed(2),
      fileTypes: fileTypes
    };
  } catch (error) {
    Logger.log('Error getting upload statistics: ' + error);
    return null;
  }
}
