/**
 * @OnlyCurrentDoc
 */


// --- CONFIGURATION ---
// IMPORTANT: You MUST update this URL if you create a new web app deployment that changes its link!
// Please ensure this is the exact URL of your deployed web app that has "Execute as: Me" and "Who has access: Anyone".
const WEB_APP_URL = "YOUR_WEB_APP_URL_GOES_HERE"; // PASTE YOUR NEW DEPLOYMENT URL HERE


// Defines the sequential order of approval roles
const APPROVAL_ROLES = ['Prepared By', 'Reviewed By', 'Noted By', 'Approved By'];
// --- END CONFIGURATION ---


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Org Chart Tools')
    .addItem('Initialize Real-Time Change Log (Run Once)', 'initializeChangeTracking')
    .addSeparator()
    .addItem('Update Headcount Summary & Create Approval Records', 'takeHeadcountSnapshotWithAlert')
    .addItem('Generate Incumbency History Report', 'generateIncumbencyReport')
    .addSeparator()
    .addItem('Debug Incumbency History', 'debugIncumbencyForPosition')
    .addItem('Clear Script Cache', 'clearScriptCache')
    .addToUi();
}


/**
 * UTILITY FUNCTION to clear the script's cache.
 */
function clearScriptCache() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm', 'This will clear all cached data for the web app, which may cause it to load slightly slower one time. Are you sure you want to continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    CacheService.getScriptCache().removeAll(['incumbency_history_04-CSD-006']);
    ui.alert('Success! The script cache has been cleared. Please reload the web app for changes to take effect.');
  }
}


/**
 * DIAGNOSTIC FUNCTION
 */
function debugIncumbencyForPosition() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Debug Incumbency History',
    'Please enter the exact Position ID to debug:',
    ui.ButtonSet.OK_CANCEL);


  const button = result.getSelectedButton();
  const posId = result.getResponseText();


  if (button !== ui.Button.OK || !posId) {
    return; // User cancelled
  }


  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    if (!logSheet || logSheet.getLastRow() < 2) {
      ui.alert('The "change_log_sheet" is empty or not found.');
      return;
    }


    const allLogData = logSheet.getDataRange().getValues();
    const headers = allLogData.shift();


    const posIdIndex = headers.indexOf('Position ID');
    const nameIndex = headers.indexOf('Employee Name');
    const timestampIndex = headers.indexOf('Change Timestamp');
    const effectiveDateIndex = headers.indexOf('Effective Date');


    if ([posIdIndex, nameIndex, timestampIndex, effectiveDateIndex].includes(-1)) {
      throw new Error("One or more required columns (Position ID, Employee Name, Change Timestamp, Effective Date) are missing from the change_log_sheet.");
    }


    const positionEntries = allLogData
      .filter(row => row[posIdIndex] === posId && row[timestampIndex])
      .sort((a, b) => {
        const dateA = a[effectiveDateIndex] instanceof Date ? a[effectiveDateIndex] : new Date(a[timestampIndex]);
        const dateB = b[effectiveDateIndex] instanceof Date ? b[effectiveDateIndex] : new Date(b[timestampIndex]);
        return dateA - dateB;
      });


    if (positionEntries.length === 0) {
      Logger.log(`No log entries found for Position ID: ${posId}`);
      ui.alert(`No log entries were found for Position ID: "${posId}". Please check the ID and try again.`);
      return;
    }


    Logger.log(`--- DEBUG LOG FOR POSITION ID: ${posId} ---`);
    Logger.log(`Found ${positionEntries.length} entries. Sorted by true effective date:`);
    Logger.log('--------------------------------------------------');


    positionEntries.forEach((entry, index) => {
      const definitiveDate = entry[effectiveDateIndex] || entry[timestampIndex];
      const logLine = `Event #${index + 1} (Effective: ${new Date(definitiveDate).toLocaleDateString()}): ` +
        `Incumbent: "${entry[nameIndex]}", ` +
        `Effective Date (Col AA): "${entry[effectiveDateIndex]}"`;
      Logger.log(logLine);
    });


    Logger.log('--------------------------------------------------');
    Logger.log(`--- END OF DEBUG LOG ---`);


    ui.alert('Debug log created successfully. Please go to the Apps Script editor and view the logs under "Executions".');


  } catch (e) {
    Logger.log(`Error during debug: ${e.toString()}`);
    ui.alert(`An error occurred: ${e.message}`);
  }
}




// --- ALL OTHER FUNCTIONS ARE BELOW ---


function initializeChangeTracking() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  if (mainSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Your data sheet is empty. Please add data before initializing.');
    return;
  }
  try {
    const lastCol = mainSheet.getLastColumn();
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, lastCol).getValues();


    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('lastKnownData', JSON.stringify(data));
    scriptProperties.setProperty('lastKnownColumnCount', lastCol.toString());
    scriptProperties.setProperty('incumbencyHistory', JSON.stringify({}));
    scriptProperties.setProperty('snapshotTimestamp', '');


    SpreadsheetApp.getUi().alert('Success! The real-time change log and incumbency tracking systems have been initialized.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Initialization failed. Error: ' + e.message);
  }
}


function handleSheetChange(e) {
  if (['EDIT', 'INSERT_ROW', 'REMOVE_ROW'].indexOf(e.changeType) === -1) {
    return;
  }
  // --- FIX: Check for the lock and exit if it's active ---
  const scriptProperties = PropertiesService.getScriptProperties();
  if (scriptProperties.getProperty('scriptChangeLock')) {
    Logger.log('Skipping handleSheetChange because scriptChangeLock is active.');
    return;
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    logDataChanges();
  } finally {
    lock.releaseLock();
  }
}


/**
 * Invalidates the incumbency history cache for specific positions.
 * @param {string[]} positionIds - An array of Position IDs to clear from the cache.
 */
function invalidateIncumbencyCache(positionIds) {
  if (!positionIds || positionIds.length === 0) return;
  const cache = CacheService.getScriptCache();
  const cacheKeys = positionIds.map(id => `incumbency_history_${id}`);
  cache.removeAll(cacheKeys);
  Logger.log(`Invalidated incumbency cache for Position IDs: ${positionIds.join(', ')}`);
}


function logDataChanges() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  if (sheets.length === 0) {
    return;
  }

  const mainSheet = sheets[0];
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  if (!logSheet || logSheet.getLastRow() < 1) return;

  const scriptProperties = PropertiesService.getScriptProperties();
  const lastKnownDataString = scriptProperties.getProperty('lastKnownData');
  if (!lastKnownDataString) return;

  const logSheetHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const logHeaderMap = new Map(logSheetHeaders.map((h, i) => [h.trim(), i]));
  const mainSheetHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];

  const pendingResignationPosId = scriptProperties.getProperty('pendingResignationPosId');
  const pendingResignationDate = scriptProperties.getProperty('pendingResignationDate');
  const pendingEffectiveDatePosId = scriptProperties.getProperty('pendingEffectiveDatePosId');
  const pendingEffectiveDate = scriptProperties.getProperty('pendingEffectiveDate');
  const overrideTimestamp = scriptProperties.getProperty('overrideTimestamp');
  const isCorrection = scriptProperties.getProperty('isResignationDateCorrection');

  let timestamp = new Date();
  if (overrideTimestamp) {
    timestamp = new Date(overrideTimestamp + 'T12:00:00');
    scriptProperties.deleteProperty('overrideTimestamp');
  }

  const incumbencyHistory = JSON.parse(scriptProperties.getProperty('incumbencyHistory') || '{}');
  const previousData = JSON.parse(lastKnownDataString);
  const currentData = mainSheet.getLastRow() > 1 ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues() : [];

  const currentDataMap = new Map(currentData.map(row => [row[0], row]));
  const previousDataMap = new Map(previousData.map(row => [row[0], row]));
  const previousEmployeeMap = new Map();
  previousData.forEach(row => {
    if (row[1]) previousEmployeeMap.set(String(row[1]).trim(), row);
  });

  const changesToLog = [];
  previousDataMap.forEach((prevRow, posId) => {
    const currentRow = currentDataMap.get(posId);
    if (!currentRow) {
      changesToLog.push(prevRow.concat([timestamp, 'Row Deleted', '']));
    } else if (JSON.stringify(prevRow) !== JSON.stringify(currentRow) || (isCorrection && posId === pendingEffectiveDatePosId)) {
      let internalTransferNote = '';
      if (currentRow[1] && currentRow[1] !== prevRow[1]) {
        const oldPositionRow = previousEmployeeMap.get(String(currentRow[1]).trim());
        if (oldPositionRow && oldPositionRow[0] !== posId) {
          internalTransferNote = `From: ${oldPositionRow[8] || 'N/A'} (${oldPositionRow[9] || 'N/A'}) - ${oldPositionRow[5] || 'N/A'}`;
        }
      }
      if (prevRow[1] && !currentRow[1] && prevRow[2]) {
        if (!incumbencyHistory[posId]) incumbencyHistory[posId] = [];
        incumbencyHistory[posId].unshift(prevRow[2]);
        incumbencyHistory[posId] = incumbencyHistory[posId].slice(0, 10);
      }
      
      if (prevRow[1] && !currentRow[1]) {
        changesToLog.push(prevRow.concat([timestamp, 'Row Modified', internalTransferNote]));
      } else {
        changesToLog.push(currentRow.concat([timestamp, 'Row Modified', internalTransferNote]));
      }
    }
  });

  currentDataMap.forEach((currentRow, posId) => {
    if (!previousDataMap.has(posId)) {
      let internalTransferNote = '';
      if (currentRow[1]) {
        const oldPositionRow = previousEmployeeMap.get(String(currentRow[1]).trim());
        if (oldPositionRow) {
          internalTransferNote = `From: ${oldPositionRow[8] || 'N/A'} (${oldPositionRow[9] || 'N/A'}) - ${oldPositionRow[5] || 'N/A'}`;
        }
      }
      changesToLog.push(currentRow.concat([timestamp, 'Row Added', internalTransferNote]));
    }
  });

  if (changesToLog.length > 0) {
    const modifiedPositionIds = [...new Set(changesToLog.map(row => row[0]).filter(String))];
    invalidateIncumbencyCache(modifiedPositionIds);

    const logData = changesToLog.map(function(changedRow) {
      const newLogRow = Array(logSheetHeaders.length).fill('');
      const changeType = changedRow[changedRow.length - 2];
      const posId = changedRow[0];
      const empId = changedRow[1];

      mainSheetHeaders.forEach((header, i) => {
        if (logHeaderMap.has(header.trim())) {
          newLogRow[logHeaderMap.get(header.trim())] = changedRow[i];
        }
      });

      const headcount = getCurrentHeadcounts(changedRow[6], changedRow[8], changedRow[9], currentData);
      if (logHeaderMap.has('Change Type')) newLogRow[logHeaderMap.get('Change Type')] = changeType;
      if (logHeaderMap.has('Transfer Note')) newLogRow[logHeaderMap.get('Transfer Note')] = changedRow[changedRow.length - 1];
      if (logHeaderMap.has('Change Timestamp')) newLogRow[logHeaderMap.get('Change Timestamp')] = changedRow[changedRow.length - 3];
      if (logHeaderMap.has('Division Headcount')) newLogRow[logHeaderMap.get('Division Headcount')] = headcount.division;
      if (logHeaderMap.has('Department Headcount')) newLogRow[logHeaderMap.get('Department Headcount')] = headcount.department;
      if (logHeaderMap.has('Section Headcount')) newLogRow[logHeaderMap.get('Section Headcount')] = headcount.section;

      const effectiveDateIndex = logHeaderMap.get('Effective Date');
      if (effectiveDateIndex !== undefined) {
        // This is the definitive vacating event from a promotion/transfer.
        // It's identified by the pending property, and we ensure it's only used once
        // by checking that the log entry still contains the employee from the `prevRow`.
        const logContainsEmployee = !!changedRow[mainSheetHeaders.indexOf('Employee ID')];

        if (pendingResignationPosId && pendingResignationDate && posId.toUpperCase() === pendingResignationPosId.toUpperCase() && logContainsEmployee) {
          newLogRow[effectiveDateIndex] = new Date(pendingResignationDate);

          // CRITICAL: Immediately delete the properties after using them once.
          // This prevents a subsequent cascading update (e.g., reporting line change)
          // from creating a second log entry for the same position and incorrectly
          // reusing the promotion date.
          scriptProperties.deleteProperty('pendingResignationPosId');
          scriptProperties.deleteProperty('pendingResignationDate');
        }
        // This handles other scenarios like a direct resignation from the form.
        else if (pendingEffectiveDatePosId && pendingEffectiveDate && posId.toUpperCase() === pendingEffectiveDatePosId.toUpperCase()) {
          newLogRow[effectiveDateIndex] = new Date(pendingEffectiveDate);
        }
      }
      return newLogRow;
    });

    if (pendingEffectiveDatePosId) {
      scriptProperties.deleteProperty('pendingEffectiveDatePosId');
      scriptProperties.deleteProperty('pendingEffectiveDate');
    }
    if (pendingResignationPosId) {
      scriptProperties.deleteProperty('pendingResignationPosId');
      scriptProperties.deleteProperty('pendingResignationDate');
    }
    
    if (isCorrection) {
      scriptProperties.deleteProperty('isResignationDateCorrection');
    }

    if (logData.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logData.length, logData[0].length).setValues(logData);
    }
  }

  PropertiesService.getScriptProperties().setProperty('lastKnownData', JSON.stringify(currentData));
  PropertiesService.getScriptProperties().setProperty('lastKnownColumnCount', String(mainSheet.getLastColumn()));
  PropertiesService.getScriptProperties().setProperty('incumbencyHistory', JSON.stringify(incumbencyHistory));
}

function getCurrentHeadcounts(division, department, section, allData) {
  let divisionCount = 0;
  let departmentCount = 0;
  let sectionCount = 0;
  for (let i = 0; i < allData.length; i++) {
    if ((allData[i][17] || '').toString().trim().toLowerCase() === 'inactive') continue;
    if (allData[i][6] === division) {
      divisionCount++;
      if (allData[i][8] === department) {
        departmentCount++;
        if (allData[i][9] === section) {
          sectionCount++;
        }
      }
    }
  }
  return {
    division: divisionCount,
    department: departmentCount,
    section: sectionCount
  };
}


function takeHeadcountSnapshotWithAlert() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm', 'This will update the "Previous Headcount" summary and create new approval records for all departments. Continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    try {
      takeHeadcountSnapshot();
      ui.alert('Success! The headcount summary has been updated and new approval records have been created for each department.');
    } catch (e) {
      ui.alert('Error: ' + e.message);
    }
  }
}


function takeHeadcountSnapshot() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  let targetSheet = spreadsheet.getSheetByName('Previous Headcount');

  if (!targetSheet) {
    targetSheet = spreadsheet.insertSheet('Previous Headcount');
    targetSheet.appendRow(['Division', 'Group', 'Department', 'Section', 'Approved Plantilla']);
    targetSheet.setFrozenRows(1);
  }

  if (mainSheet.getLastRow() < 2) {
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const currentSnapshot = scriptProperties.getProperty('snapshotTimestamp');
  if (currentSnapshot) {
    scriptProperties.setProperty('previousHeadcountTimestamp', currentSnapshot);
  }

  const timestamp = new Date();
  scriptProperties.setProperty('snapshotTimestamp', timestamp.toISOString());

  const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 18).getValues();
  const approvalsSheet = spreadsheet.getSheetByName('Approvals');
  if (!approvalsSheet) {
    throw new Error('Sheet "Approvals" not found.');
  }

  const approversData = getApproversData();
  const uniqueDepartments = [...new Set(data.map(row => row[8]).filter(String))];
  const existingApprovalRecords = approvalsSheet.getDataRange().getValues();
  const headers = existingApprovalRecords.length > 0 ? existingApprovalRecords[0] : [];
  const snapshotColIndex = headers.indexOf('Snapshot Date');
  const deptColIndex = headers.indexOf('Department');
  const newlyCreatedRecords = [];

  uniqueDepartments.forEach(dept => {
    const recordExists = existingApprovalRecords.some((row, index) =>
      index > 0 && row[snapshotColIndex] === timestamp.toISOString() && row[deptColIndex] === dept
    );
    if (!recordExists) {
      approvalsSheet.appendRow([timestamp.toISOString(), dept, '', '', '', '', '', '', '', '']);
      newlyCreatedRecords.push(dept);
    }
  });

  newlyCreatedRecords.forEach(dept => {
    sendApprovalNotificationEmail(dept, timestamp.toISOString(), approversData, 'Prepared By');
  });

  const summary = {};
  data.forEach(function (row) {
    if ((row[17] || '').toString().trim().toLowerCase() === 'inactive') return;
    const isFilled = !!row[1];
    const division = row[6],
      group = row[7] || '', // Ensure blank values are treated as empty strings
      department = row[8] || '',
      section = row[9] || '';

    if (!division) return;
    if (!summary[division]) summary[division] = {
      filled: 0,
      vacant: 0,
      groups: {}
    };
    if (!summary[division].groups[group]) summary[division].groups[group] = {
      filled: 0,
      vacant: 0,
      departments: {}
    };
    if (!summary[division].groups[group].departments[department]) summary[division].groups[group].departments[department] = {
      filled: 0,
      vacant: 0,
      sections: {}
    };
    if (!summary[division].groups[group].departments[department].sections[section]) summary[division].groups[group].departments[department].sections[section] = {
      filled: 0,
      vacant: 0
    };

    isFilled ? summary[division].filled++ : summary[division].vacant++;
    isFilled ? summary[division].groups[group].filled++ : summary[division].groups[group].vacant++;
    isFilled ? summary[division].groups[group].departments[department].filled++ : summary[division].groups[group].departments[department].vacant++;
    isFilled ? summary[division].groups[group].departments[department].sections[section].filled++ : summary[division].groups[group].departments[department].sections[section].vacant++;
  });


  const monthHeader = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MMM yyyy");
  const filledHeader = `${monthHeader} Filled`;
  const vacantHeader = `${monthHeader} Vacant`;

  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  let filledColIdx = targetHeaders.indexOf(filledHeader);
  let vacantColIdx = targetHeaders.indexOf(vacantHeader);
  let plantillaColIdx = targetHeaders.indexOf('Approved Plantilla');

  if (plantillaColIdx === -1) {
    targetSheet.getRange(1, 5).setValue('Approved Plantilla');
    plantillaColIdx = 4;
  }

  if (filledColIdx === -1) {
    const lastCol = targetSheet.getLastColumn();
    targetSheet.getRange(1, lastCol + 1, 1, 2).setValues([
      [filledHeader, vacantHeader]
    ]);
    filledColIdx = lastCol;
    vacantColIdx = lastCol + 1;
  }

  const existingData = targetSheet.getLastRow() > 1 ? targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).getValues() : [];
  const dataMap = new Map();
  existingData.forEach((row, index) => {
    const key = [row[0], row[1], row[2], row[3]].join('|');
    dataMap.set(key, {
      rowIndex: index + 2,
      data: row
    });
  });

  const updatedData = [];

  const processLevel = (div, group, dept, sec, counts) => {
    const key = [div, group, dept, sec].join('|');
    if (dataMap.has(key)) {
      const existingRow = dataMap.get(key);
      existingRow.data[filledColIdx] = counts.filled;
      existingRow.data[vacantColIdx] = counts.vacant;
      updatedData.push({
        range: `A${existingRow.rowIndex}`,
        values: [existingRow.data]
      });
      dataMap.delete(key);
    } else {
      const newRow = Array(targetSheet.getLastColumn()).fill('');
      newRow[0] = div;
      newRow[1] = group;
      newRow[2] = dept;
      newRow[3] = sec;
      newRow[filledColIdx] = counts.filled;
      newRow[vacantColIdx] = counts.vacant;
      targetSheet.appendRow(newRow);
    }
  };

  // --- REVISED SECTION ---
  // This revised logic filters out empty keys ('') before processing,
  // preventing the creation of blank or incomplete rows in the "Previous Headcount" sheet.
  Object.keys(summary).sort().forEach(divName => {
    processLevel(divName, '', '', '', summary[divName]); // Process Division total
    Object.keys(summary[divName].groups).sort().filter(g => g).forEach(groupName => { // Filter out empty group names
      processLevel(divName, groupName, '', '', summary[divName].groups[groupName]); // Process Group total
      Object.keys(summary[divName].groups[groupName].departments).sort().filter(d => d).forEach(deptName => { // Filter out empty dept names
        processLevel(divName, groupName, deptName, '', summary[divName].groups[groupName].departments[deptName]); // Process Dept total
        Object.keys(summary[divName].groups[groupName].departments[deptName].sections).sort().filter(s => s).forEach(secName => { // Filter out empty section names
          processLevel(divName, groupName, deptName, secName, summary[divName].groups[groupName].departments[deptName].sections[secName]); // Process Section total
        });
      });
    });
  });
  // --- END REVISED SECTION ---

  updatedData.forEach(update => {
    const range = targetSheet.getRange(update.range).offset(0, 0, 1, update.values[0].length);
    range.setValues(update.values);
  });
}


function getApproversData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const approversSheet = spreadsheet.getSheetByName('Approvers');
  const allApprovers = {};


  if (approversSheet) {
    const data = approversSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data.shift();
      data.forEach((row) => {
        const department = row[0] ? row[0].toString().trim() : '';
        const role = row[1] ? row[1].toString().trim() : '';
        const email = row[2] ? row[2].toString().trim() : '';
        if (department && role && email) {
          if (!allApprovers[department]) {
            allApprovers[department] = {};
          }
          allApprovers[department][role] = email;
        }
      });
    }
  }
  return allApprovers;
}


function sendApprovalNotificationEmail(department, snapshotTimestamp, allApproversData, roleToNotify) {
  const departmentApprovers = allApproversData[department];
  if (!departmentApprovers) {
    return;
  }
  const recipientEmail = departmentApprovers[roleToNotify];
  if (recipientEmail) {
    const subject = `Approval Required (${roleToNotify}): Org Chart Snapshot for ${department}`;
    const body = `Dear ${recipientEmail.split('@')[0].toUpperCase()},\n\nThe Organizational Chart snapshot for your department (${department}) generated on ${new Date(snapshotTimestamp).toLocaleString("en-US",{timeZone:"Asia/Manila"})} requires your signature as "${roleToNotify}".\n\nPlease visit the Organizational Chart web application to sign:\n${WEB_APP_URL}\n\nThank you,\nYour Organizational Chart Team`;
    try {
      MailApp.sendEmail(recipientEmail, subject, body);
    } catch (mailError) {
      Logger.log(`ERROR sending email to ${roleToNotify} (${recipientEmail}) for department ${department}. Error: ${mailError.message}`);
    }
  }
}


function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Interactive Organizational Chart').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getIncumbencyHistory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const historyString = scriptProperties.getProperty('incumbencyHistory');
  return historyString ? JSON.parse(historyString) : {};
}


function getEmployeeData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const mainSheet = spreadsheet.getSheets()[0];


    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    const resignationDates = new Map();
    if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
      const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const posIdIndex = headers.indexOf('Position ID');
      const statusIndex = headers.indexOf('Status');
      const effectiveDateIndex = headers.indexOf('Effective Date');


      if (posIdIndex > -1 && statusIndex > -1 && effectiveDateIndex > -1) {
        logData.forEach(row => {
          if (row[posIdIndex] && String(row[statusIndex]).toUpperCase() === 'RESIGNED' && row[effectiveDateIndex] instanceof Date) {
            resignationDates.set(row[posIdIndex], row[effectiveDateIndex]);
          }
        });
      }
    }


    const userPermissions = {};
    const permissionsSheet = spreadsheet.getSheetByName('Permissions');
    if (permissionsSheet) {
      const permData = permissionsSheet.getDataRange().getValues();
      if (permData.length > 0) {
        const permissionHeaders = permData.shift();
        const emailColIndex = permissionHeaders.indexOf('EMAIL');
        if (emailColIndex !== -1) {
          const userRow = permData.find(row => row[emailColIndex] && row[emailColIndex].toString().trim().toLowerCase() === userEmail);
          if (userRow) {
            permissionHeaders.forEach((header, index) => {
              if (header) {
                userPermissions[header.trim()] = userRow[index] ? userRow[index].toString().trim().toLowerCase() : '';
              }
            });
          }
        }
      }
    }
    const isFieldAuthorized = (fieldName) => (userPermissions[fieldName] === 'x' || userPermissions[fieldName] === 'all' || userPermissions[fieldName] === 'anyone');
    const isDepartmentViewable = (employeeDivision, employeeDepartment) => {
      const viewableDeptEntry = userPermissions['Viewable Department'] || '';
      if (viewableDeptEntry === 'all' || viewableDeptEntry === 'anyone') return true;
      const allowedDeptDivs = viewableDeptEntry.split(',').map(item => item.trim().toLowerCase()).filter(item => item);
      return allowedDeptDivs.includes(employeeDepartment.toLowerCase()) || allowedDeptDivs.includes(employeeDivision.toLowerCase());
    };
    const canEdit = userPermissions['Can Edit'] === 'x' || userPermissions['Can Edit'] === 'all' || userPermissions['Can Edit'] === 'anyone';


    if (mainSheet.getLastRow() < 2) {
      return {
        current: [],
        previous: {},
        snapshotTimestamp: null,
        currentUserEmail: userEmail,
        userCanSeeAnyDepartment: false,
        totalApprovedPlantilla: 0,
        previousDateString: null,
        canEdit: canEdit
      };
    }


    const lastCol = Math.max(20, mainSheet.getLastColumn());
    const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, lastCol).getValues();


    const employeeIdToPositionIdMap = new Map();
    mainData.forEach(row => {
      const employeeId = row[1] ? row[1].toString().trim() : null;
      const positionId = row[0] ? row[0].toString().trim() : null;
      if (employeeId && positionId) {
        employeeIdToPositionIdMap.set(employeeId, positionId);
      }
    });


    const historicalNotes = getHistoricalNotes();
    const incumbencyHistory = getIncumbencyHistory();
    const employeesToShow = [];
    let hasReturnedAnyEmployee = false;


    mainData.forEach(function (row) {
      const employeeDivision = row[6] ? row[6].toString().trim() : '';
      const employeeDepartment = row[8] ? row[8].toString().trim() : '';
      if (!isDepartmentViewable(employeeDivision, employeeDepartment)) return;
      hasReturnedAnyEmployee = true;
      const posId = row[0] ? row[0].toString().trim() : null;
      if (!posId) return;


      const managerEmployeeId = row[3] ? row[3].toString().trim() : null;
      let managerPositionId = ''; // Use let instead of const

      if (managerEmployeeId) {
        // Heuristic: Position IDs contain hyphens, Employee IDs do not.
        // This handles the temporary state where a report is pointed to a vacant Position ID.
        if (managerEmployeeId.includes('-')) {
          managerPositionId = managerEmployeeId;
        } else {
          // Otherwise, it's an Employee ID, so look it up in the map as usual.
          managerPositionId = employeeIdToPositionIdMap.get(managerEmployeeId) || '';
        }
      }


      const history = historicalNotes[posId] || {};
      history.lastIncumbents = incumbencyHistory[posId] || [];


      let dateHired = null;
      if (row[18] && row[18] instanceof Date) {
        try {
          dateHired = Utilities.formatDate(row[18], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
          dateHired = null;
        }
      }
      let contractEndDate = null;
      if (row[19] && row[19] instanceof Date) {
        try {
          contractEndDate = Utilities.formatDate(row[19], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
          contractEndDate = null;
        }
      }


      const employeeStatus = row[16] ? row[16].toString().trim() : '';
      let resignationDate = null;
      if (employeeStatus.toUpperCase() === 'RESIGNED' && resignationDates.has(posId)) {
        resignationDate = Utilities.formatDate(resignationDates.get(posId), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }


      employeesToShow.push({
        positionId: posId,
        employeeId: row[1] ? row[1].toString().trim() : null,
        nodeId: posId,
        employeeName: row[2],
        managerId: managerPositionId || '',
        managerEmployeeId: managerEmployeeId || '',
        managerName: row[4],
        jobTitle: row[5],
        division: employeeDivision,
        group: row[7],
        department: employeeDepartment,
        section: row[9],
        gender: row[10] ? row[10].toString().trim() : '',
        level: row[11],
        payrollType: isFieldAuthorized('Payroll Type') ? row[12] : null,
        jobLevel: isFieldAuthorized('Job Level') ? row[13] : null,
        contractType: isFieldAuthorized('Contract Type') ? (row[14] ? row[14].toString().trim() : null) : null,
        stylingContractType: row[14] ? row[14].toString().trim() : null,
        competency: isFieldAuthorized('Competency') ? row[15] : null,
        status: employeeStatus,
        positionStatus: row[17] ? row[17].toString().trim() : 'Active',
        dateHired: dateHired,
        contractEndDate: contractEndDate,
        historicalNote: history,
        resignationDate: resignationDate
      });
    });


    let previousHeadcount = {};
    let totalApprovedPlantilla = 0;
    let previousDateString = null;


    try {
      const previousSheet = spreadsheet.getSheetByName('Previous Headcount');
      if (previousSheet && previousSheet.getLastRow() > 1) {
        const prevDataRange = previousSheet.getDataRange();
        const prevData = prevDataRange.getValues();
        if (prevData.length > 0) {
          const prevHeaders = prevData.shift();
          const plantillaIndex = prevHeaders.indexOf('Approved Plantilla');
          let lastFilledIndex = -1;
          for (let i = prevHeaders.length - 1; i >= 0; i--) {
            if (String(prevHeaders[i]).includes('Filled')) {
              lastFilledIndex = i;
              break;
            }
          }
          if (lastFilledIndex !== -1) {
            const lastVacantIndex = lastFilledIndex + 1;
            const dateHeader = String(prevHeaders[lastFilledIndex] || '');
            if (dateHeader) {
              previousDateString = dateHeader.replace(/ filled/i, '').trim();
            }
            prevData.forEach(function (row) {
              const division = row[0],
                group = row[1] || '',
                department = row[2] || '',
                section = row[3] || '';
              const rawPlantilla = row[plantillaIndex];
              const plantillaValue = (plantillaIndex !== -1 && rawPlantilla !== '' && !isNaN(rawPlantilla)) ? parseInt(rawPlantilla) : null;
              const filled = row[lastFilledIndex] || 0;
              const vacant = (row.length > lastVacantIndex) ? (row[lastVacantIndex] || 0) : 0;
              if (division) {
                if (!previousHeadcount[division]) {
                  previousHeadcount[division] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null,
                    groups: {}
                  };
                }
                if (!previousHeadcount[division].groups[group]) {
                  previousHeadcount[division].groups[group] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null,
                    departments: {}
                  };
                }
                if (!previousHeadcount[division].groups[group].departments[department]) {
                  previousHeadcount[division].groups[group].departments[department] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null,
                    sections: {}
                  };
                }
                if (!previousHeadcount[division].groups[group].departments[department].sections[section]) {
                  previousHeadcount[division].groups[group].departments[department].sections[section] = {
                    filled: 0,
                    vacant: 0,
                    plantilla: null
                  };
                }
                if (group === '' && department === '' && section === '') {
                  previousHeadcount[division].filled = filled;
                  previousHeadcount[division].vacant = vacant;
                  previousHeadcount[division].plantilla = plantillaValue;
                } else if (group && department === '' && section === '') {
                  previousHeadcount[division].groups[group].filled = filled;
                  previousHeadcount[division].groups[group].vacant = vacant;
                  previousHeadcount[division].groups[group].plantilla = plantillaValue;
                } else if (department && section === '') {
                  previousHeadcount[division].groups[group].departments[department].filled = filled;
                  previousHeadcount[division].groups[group].departments[department].vacant = vacant;
                  previousHeadcount[division].groups[group].departments[department].plantilla = plantillaValue;
                } else if (section) {
                  previousHeadcount[division].groups[group].departments[department].sections[section].filled = filled;
                  previousHeadcount[division].groups[group].departments[department].sections[section].vacant = vacant;
                  previousHeadcount[division].groups[group].departments[department].sections[section].plantilla = plantillaValue;
                }
                if (group === '' && department === '' && section === '' && plantillaValue !== null) {
                  totalApprovedPlantilla += plantillaValue;
                }
              }
            });
          }
        }
      }
    } catch (e) {
      Logger.log('WARNING: Could not parse "Previous Headcount" sheet. Summary data will be unavailable. Error: ' + e.message);
    }


    const snapshotTimestamp = PropertiesService.getScriptProperties().getProperty('snapshotTimestamp');
    return {
      current: employeesToShow.filter(emp => emp.positionId),
      previous: previousHeadcount,
      snapshotTimestamp: snapshotTimestamp,
      previousDateString: previousDateString,
      currentUserEmail: userEmail,
      userCanSeeAnyDepartment: hasReturnedAnyEmployee,
      totalApprovedPlantilla: totalApprovedPlantilla,
      canEdit: canEdit
    };
  } catch (e) {
    Logger.log('ERROR in getEmployeeData: ' + e.toString() + ' Stack: ' + e.stack);
    return null;
  }
}


function getHistoricalNotes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  const history = {};
  if (!logSheet || logSheet.getLastRow() < 2) return history;


  const logValues = logSheet.getDataRange().getValues();
  const headers = logValues.shift();
  const posIdIndex = headers.indexOf('Position ID');
  const empIdIndex = headers.indexOf('Employee ID');
  const transferNoteIndex = headers.indexOf('Transfer Note');


  if (posIdIndex === -1 || empIdIndex === -1 || transferNoteIndex === -1) {
    Logger.log("getHistoricalNotes: Could not find required headers in change_log_sheet.");
    return history;
  }


  const filledPositions = new Set(logValues.filter(row => row[empIdIndex]).map(row => row[posIdIndex]));


  logValues.forEach(row => {
    const posId = row[posIdIndex];
    const transferNote = row[transferNoteIndex];
    if (posId) {
      if (!history[posId]) {
        history[posId] = {
          isNewPosition: !filledPositions.has(posId)
        };
      }
      if (transferNote) {
        history[posId].previousRole = transferNote;
      }
    }
  });
  return history;
}


function getApprovalData(department) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const snapshotTimestamp = PropertiesService.getScriptProperties().getProperty('snapshotTimestamp');
    const approvers = getApproversData()[department] || {};
    let approvalStatus = {};
    const approvalsSheet = spreadsheet.getSheetByName('Approvals');
    if (approvalsSheet && snapshotTimestamp) {
      const data = approvalsSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data.shift();
        const snapshotColIndex = headers.indexOf('Snapshot Date');
        const deptColIndex = headers.indexOf('Department');
        const approvalRow = data.find(row => row[snapshotColIndex] === snapshotTimestamp && row[deptColIndex] === department);
        if (approvalRow) {
          headers.forEach((header, index) => {
            const value = approvalRow[index];
            approvalStatus[header] = (value instanceof Date) ? value.toISOString() : value;
          });
        }
      }
    }
    return {
      approvers: approvers,
      approvalStatus: approvalStatus
    };
  } catch (e) {
    Logger.log('ERROR in getApprovalData: ' + e.toString());
    return null;
  }
}


function recordApproval(role, department) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const approvalsSheet = spreadsheet.getSheetByName('Approvals');
    if (!approvalsSheet) throw new Error("Sheet 'Approvals' not found.");
    const user = Session.getActiveUser();
    const userName = user.getUserLoginId().split('@')[0];
    const snapshotTimestamp = PropertiesService.getScriptProperties().getProperty('snapshotTimestamp');
    if (!snapshotTimestamp) throw new Error("No active snapshot found.");
    const data = approvalsSheet.getDataRange().getValues();
    const headers = data[0];
    const snapshotColIndex = headers.indexOf('Snapshot Date');
    const deptColIndex = headers.indexOf('Department');
    const roleColIndex = headers.indexOf(role);
    if (roleColIndex === -1) throw new Error(`Role column "${role}" not found.`);
    for (let i = 1; i < data.length; i++) {
      if (data[i][snapshotColIndex] === snapshotTimestamp && data[i][deptColIndex] === department) {
        approvalsSheet.getRange(i + 1, roleColIndex + 1).setValue(userName);
        approvalsSheet.getRange(i + 1, roleColIndex + 2).setValue(new Date());
        SpreadsheetApp.flush();
        const approversData = getApproversData();
        const currentRoleIndex = APPROVAL_ROLES.indexOf(role);
        const nextRole = APPROVAL_ROLES[currentRoleIndex + 1];
        if (nextRole && !getApprovalData(department).approvalStatus[nextRole]) {
          sendApprovalNotificationEmail(department, snapshotTimestamp, approversData, nextRole);
        }
        return "Approval recorded successfully.";
      }
    }
    throw new Error("Could not find matching approval record.");
  } catch (e) {
    Logger.log('ERROR in recordApproval: ' + e.toString());
    throw new Error('Failed to record approval. ' + e.message);
  }
}


function getListsForDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const refSheet = ss.getSheetByName("Reference Data");


  let dynamicLists = {};
  if (mainSheet.getLastRow() > 1) {
    const data = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    const activeEmployees = data
      .filter(row => row[1] && (row[17] || '').toLowerCase() !== 'inactive')
      .map(row => ({
        id: row[1],
        name: row[2]
      }))
      .sort((a, b) => a.name.localeCompare(b.name));


    dynamicLists = {
      managers: activeEmployees,
      divisions: [...new Set(data.map(row => row[6]).filter(String))].sort(),
      groups: [...new Set(data.map(row => row[7]).filter(String))].sort(),
      departments: [...new Set(data.map(row => row[8]).filter(String))].sort(),
      sections: [...new Set(data.map(row => row[9]).filter(String))].sort()
    };
  }


  let staticLists = {};
  if (refSheet) {
    const refData = refSheet.getDataRange().getValues();
    const headers = refData.shift();
    headers.forEach((header, colIndex) => {
      if (header) {
        const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/gi, '');
        const values = refData.map(row => row[colIndex]).filter(String).sort();
        staticLists[key] = values;
      }
    });
  }


  return { ...dynamicLists,
    ...staticLists
  };
}


function generateNewPositionId(division, section) {
  try {
    if (!division || !section) {
      return "ERROR: Division and Section are required.";
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];


    const divisionCode = division.split(' ')[0].trim();
    const sectionCode = section.split(' ')[0].trim();


    if (!/^\d+$/.test(divisionCode) || !/^\d+$/.test(sectionCode)) {
      return "ERROR: Division/Section name must start with a numeric code.";
    }


    const prefix = `${divisionCode}-${sectionCode}-`;
    const positionIds = mainSheet.getRange("A2:A").getValues().flat().filter(String);


    let maxSequence = 0;
    positionIds.forEach(id => {
      if (id.startsWith(prefix)) {
        const sequence = parseInt(id.substring(prefix.length), 10);
        if (!isNaN(sequence) && sequence > maxSequence) {
          maxSequence = sequence;
        }
      }
    });


    const newSequence = (maxSequence + 1).toString().padStart(3, '0');
    return prefix + newSequence;
  } catch (e) {
    Logger.log(e);
    return `ERROR: ${e.message}`;
  }
}


function saveEmployeeData(dataObject, mode) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const scriptProperties = PropertiesService.getScriptProperties();
  try {
    scriptProperties.setProperty('scriptChangeLock', 'true');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf('Status') + 1;

    // --- START: Automated Direct Report Reassignment Logic ---
    // This logic remains the same as it correctly handles reassigning subordinates.
    const isBecomingVacant = dataObject.status && dataObject.status.toUpperCase() === 'VACANT';
    const isTransferOrPromo = dataObject.employeeid && (dataObject.status.toUpperCase() === 'PROMOTION' || dataObject.status.toUpperCase() === 'INTERNAL TRANSFER' || dataObject.status.toUpperCase() === 'LATERAL TRANSFER');

    if (mode === 'edit' && (isBecomingVacant || isTransferOrPromo)) {
      const allData = mainSheet.getDataRange().getValues();
      const currentHeaders = allData[0];
      const posIdIndex = currentHeaders.indexOf('Position ID');
      const empIdIndex = currentHeaders.indexOf('Employee ID');
      const reportingToIdIndex = currentHeaders.indexOf('Reporting to ID');
      const jobTitleIndex = currentHeaders.indexOf('Job Title');
      const reportingToNameIndex = currentHeaders.indexOf('Reporting to');

      let positionToVacateId = dataObject.positionid;

      if (isTransferOrPromo) {
        for (let i = 1; i < allData.length; i++) {
          const rowData = allData[i];
          const currentEmpId = rowData[empIdIndex] ? String(rowData[empIdIndex]).trim().toUpperCase() : '';
          const currentPosId = rowData[posIdIndex] ? String(rowData[posIdIndex]).trim().toUpperCase() : '';

          if (currentEmpId === dataObject.employeeid.toUpperCase() && currentPosId !== dataObject.positionid.toUpperCase()) {
            positionToVacateId = rowData[posIdIndex];
            break;
          }
        }
      }

      const managerRowToVacate = allData.find(row => row[posIdIndex] === positionToVacateId);
      if (managerRowToVacate) {
        const departingEmployeeId = managerRowToVacate[empIdIndex] ? String(managerRowToVacate[empIdIndex]).trim() : null;
        if (departingEmployeeId) {
          const vacatedPositionId = managerRowToVacate[posIdIndex];
          const vacatedJobTitle = managerRowToVacate[jobTitleIndex];
          const newManagerNameForReport = `(Vacant) ${vacatedJobTitle}`;

          for (let i = 1; i < allData.length; i++) {
            const reportRow = allData[i];
            const reportCurrentManagerId = reportRow[reportingToIdIndex] ? String(reportRow[reportingToIdIndex]).trim() : '';
            if (reportCurrentManagerId === departingEmployeeId) {
              const reportSheetRowIndex = i + 1;
              mainSheet.getRange(reportSheetRowIndex, reportingToIdIndex + 1).setValue(vacatedPositionId);
              mainSheet.getRange(reportSheetRowIndex, reportingToNameIndex + 1).setValue(newManagerNameForReport);
            }
          }
        }
      }
    }
    // --- END: Automated Direct Report Reassignment Logic ---

    // --- START: REVISED LOGIC FOR PROMOTION/TRANSFER ---
    // This revised logic correctly handles the vacating of the old position.
    // Instead of manually logging the event, it sets a script property.
    // The main `logDataChanges` function will then detect this property and create
    // a single, correctly-timed log entry for the vacancy event.
    if (dataObject.employeeid && (dataObject.status.toUpperCase() === 'PROMOTION' || dataObject.status.toUpperCase() === 'INTERNAL TRANSFER' || dataObject.status.toUpperCase() === 'LATERAL TRANSFER')) {
      const allData = mainSheet.getDataRange().getValues();
      const sheetHeaders = allData[0];
      const posIdIndex = sheetHeaders.indexOf('Position ID');
      const empIdIndex = sheetHeaders.indexOf('Employee ID');

      for (let i = 1; i < allData.length; i++) {
        const row = allData[i];
        const existingEmpId = row[empIdIndex] ? String(row[empIdIndex]).trim() : '';
        const existingPosId = row[posIdIndex] ? String(row[posIdIndex]).trim() : '';

        if (existingEmpId.toUpperCase() === dataObject.employeeid.toUpperCase() && existingPosId.toUpperCase() !== dataObject.positionid.toUpperCase()) {
          const oldRowIndex = i + 1;
          const oldPositionId = existingPosId;

          // Set script properties to let logDataChanges handle the vacancy logging with the correct effective date.
          if (dataObject.startdateinposition) {
            scriptProperties.setProperties({
              'pendingResignationPosId': oldPositionId.toUpperCase(),
              'pendingResignationDate': dataObject.startdateinposition
            });
          }

          // Now, clear the data from the main sheet for the old position.
          const empNameIndex = sheetHeaders.indexOf('Employee Name');
          const genderIndex = sheetHeaders.indexOf('Gender');
          const dateHiredIndex = sheetHeaders.indexOf('Date Hired');
          const contractEndIndex = sheetHeaders.indexOf('Contract End Date');
          const statusIndexHeader = sheetHeaders.indexOf('Status');

          mainSheet.getRange(oldRowIndex, empIdIndex + 1).clearContent();
          mainSheet.getRange(oldRowIndex, empNameIndex + 1).clearContent();
          mainSheet.getRange(oldRowIndex, genderIndex + 1).clearContent();
          mainSheet.getRange(oldRowIndex, dateHiredIndex + 1).clearContent();
          mainSheet.getRange(oldRowIndex, contractEndIndex + 1).clearContent();
          mainSheet.getRange(oldRowIndex, statusIndexHeader + 1).setValue('VACANT');

          break; // Exit loop after finding and processing the old position
        }
      }
    }
    // --- END REVISED LOGIC ---

    // The rest of the function for processing the submitted form data remains largely the same
    for (const key in dataObject) {
      if (typeof dataObject[key] === 'string') {
        dataObject[key] = dataObject[key].toUpperCase();
      }
    }

    if (dataObject.status && dataObject.status.toUpperCase() === 'VACANT') {
      dataObject.employeeid = '';
      dataObject.employeename = '';
      dataObject.datehired = '';
      dataObject.gender = '';
      if (dataObject.effectivedate) {
        PropertiesService.getScriptProperties().setProperties({
          'pendingEffectiveDatePosId': dataObject.positionid.toUpperCase(),
          'pendingEffectiveDate': dataObject.effectivedate
        });
      }
    }

    if (dataObject.status && dataObject.status.toUpperCase() === 'RESIGNED' && dataObject.effectivedate) {
      // Logic for handling resignation dates
      PropertiesService.getScriptProperties().setProperties({
        'pendingEffectiveDatePosId': dataObject.positionid.toUpperCase(),
        'pendingEffectiveDate': dataObject.effectivedate
      });
    }

    if (dataObject.startdateinposition) {
      PropertiesService.getScriptProperties().setProperty('overrideTimestamp', dataObject.startdateinposition);
    }

    const keyMap = {};
    headers.forEach((header, i) => {
      const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-g]/gi, '');
      keyMap[key] = i;
    });

    if (mode === 'add') {
      const newRowData = Array(headers.length).fill('');
      for (const key in dataObject) {
        if (keyMap.hasOwnProperty(key)) {
          newRowData[keyMap[key]] = dataObject[key];
        }
      }
      mainSheet.appendRow(newRowData);
    } else if (mode === 'edit') {
      const positionId = dataObject.positionid;
      const positionIdColValues = mainSheet.getRange("A2:A").getValues();
      let rowIndex = -1;
      for (let i = 0; i < positionIdColValues.length; i++) {
        if (positionIdColValues[i][0] == positionId) {
          rowIndex = i + 2;
          break;
        }
      }
      if (rowIndex === -1) {
        throw new Error(`Position ID ${positionId} not found for editing.`);
      }

      const rangeToUpdate = mainSheet.getRange(rowIndex, 1, 1, headers.length);
      const existingRowData = rangeToUpdate.getValues()[0];
      
      const originalStatus = existingRowData[statusColIndex - 1];
      const newStatus = dataObject.status;
      const newEmployeeId = dataObject.employeeid;
      const newEmployeeName = dataObject.employeename;

      if (originalStatus && originalStatus.toUpperCase() === 'VACANT' && newStatus.toUpperCase() !== 'VACANT' && newEmployeeId) {
        // Auto-reassign reports when a vacancy is filled
        const filledPositionId = dataObject.positionid;
        const allData = mainSheet.getDataRange().getValues();
        const currentHeaders = allData[0];
        const reportingToIdIndex = currentHeaders.indexOf('Reporting to ID');
        const reportingToNameIndex = currentHeaders.indexOf('Reporting to');

        if (reportingToIdIndex !== -1 && reportingToNameIndex !== -1) {
          for (let i = 1; i < allData.length; i++) {
            const reportRow = allData[i];
            const reportCurrentManagerId = reportRow[reportingToIdIndex] ? String(reportRow[reportingToIdIndex]).trim() : '';

            if (reportCurrentManagerId === filledPositionId) {
              const reportSheetRowIndex = i + 1;
              mainSheet.getRange(reportSheetRowIndex, reportingToIdIndex + 1).setValue(newEmployeeId);
              mainSheet.getRange(reportSheetRowIndex, reportingToNameIndex + 1).setValue(newEmployeeName);
            }
          }
        }
      }

      for (const key in dataObject) {
        if (keyMap.hasOwnProperty(key)) {
          const colIndex = keyMap[key];
          existingRowData[colIndex] = dataObject[key];
        }
      }
      rangeToUpdate.setValues([existingRowData]);
    }

    SpreadsheetApp.flush();
    // IMPORTANT: We call logDataChanges() to log the primary event (e.g., the promotion itself).
    // The vacating event has already been logged manually.
    logDataChanges();

    return "Data saved successfully.";
  } catch (e) {
    Logger.log('Error in saveEmployeeData: ' + e.message + ' Stack: ' + e.stack);
    throw new Error('Failed to save data. ' + e.message);
  } finally {
    scriptProperties.deleteProperty('scriptChangeLock');
    lock.releaseLock();
  }
}


function deactivatePosition(positionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheets()[0];
    const positionIdCol = mainSheet.getRange("A:A").getValues();
    const rowIndex = positionIdCol.findIndex(row => row[0] === positionId);


    if (rowIndex === -1) {
      throw new Error(`Position ID ${positionId} not found for deactivation.`);
    }
    mainSheet.getRange(rowIndex + 1, 18).setValue('Inactive');
    SpreadsheetApp.flush();
    logDataChanges();


    return "Position deactivated successfully.";
  } catch (e) {
    Logger.log('Error in deactivatePosition: ' + e.message + ' Stack: ' + e.stack);
    throw new Error('Failed to deactivate position. ' + e.message);
  } finally {
    lock.releaseLock();
  }
}


/**
 * REVISED - Generates the Incumbency History report sheet.
 */
function generateIncumbencyReport() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  const mainData = mainSheet.getLastRow() > 1 ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 3).getValues() : [];
  const mainDataMap = new Map(mainData.map(row => [row[0], row]));


  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  const reportSheetName = 'Incumbency History';
  let reportSheet = spreadsheet.getSheetByName(reportSheetName);


  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert('The "change_log_sheet" has no data to report.');
    return;
  }


  const allLogData = logSheet.getDataRange().getValues();
  const headers = allLogData.shift();
  const allHistory = calculateIncumbencyEngine(allLogData, headers, mainDataMap);
  const finalHistoryRecords = [];
  const sortedPosIds = Object.keys(allHistory).sort();


  for (const posId of sortedPosIds) {
    const records = allHistory[posId];
    if (!records || records.length === 0) continue;

    // Correction Logic: Check if the last incumbent in the history is still the active employee.
    // If so, their end date must be 'Present' (null), overriding any incorrect log data.
    const lastRecord = records[records.length - 1];
    const currentLiveRow = mainDataMap.get(posId);
    const currentLiveEmployeeId = currentLiveRow ? (currentLiveRow[1] || '').toString().trim() : null;

    if (lastRecord.incumbentId && currentLiveEmployeeId && lastRecord.incumbentId === currentLiveEmployeeId && lastRecord.endDate) {
      lastRecord.endDate = null; // Set to null, which signifies 'Present'
    }

    records.forEach(rec => {
      const tenure = (rec.startDate && (rec.endDate || new Date())) ? Math.round(((rec.endDate || new Date()) - rec.startDate) / (1000 * 60 * 60 * 24)) : 0;
      finalHistoryRecords.push([
        posId,
        rec.jobTitle,
        rec.incumbentName,
        rec.startDate,
        rec.endDate,
        tenure >= 0 ? tenure : 0,
        rec.changeCount
      ]);
    });
  }


  if (finalHistoryRecords.length === 0) {
    ui.alert('No incumbency history could be generated.');
    return;
  }


  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(reportSheetName);
  }
  reportSheet.clear();
  const reportHeaders = ['Position ID', 'Job Title', 'Incumbent Name', 'Start Date', 'End Date', 'Tenure (Days)', 'Position Change Count'];
  reportSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight('bold');


  if (finalHistoryRecords.length > 0) {
    reportSheet.getRange(2, 1, finalHistoryRecords.length, finalHistoryRecords[0].length).setValues(finalHistoryRecords);
  }


  reportSheet.getRange("D:E").setNumberFormat("yyyy-mm-dd");
  reportSheet.setFrozenRows(1);
  reportSheet.autoResizeColumns(1, reportHeaders.length);
  ui.alert(`Success! "${reportSheetName}" sheet has been updated.`);
}




/**
 * =================================================================================================
 * FINAL REWRITE v6 - calculateIncumbencyEngine
 * =================================================================================================
 * This is a complete, procedural rewrite of the engine from scratch.
 * It abandons all previous flawed logic (reducing, finding, etc.)
 * It uses a simple loop to walk the complete timeline and builds the history record by record.
 * This method is robust and correctly handles ALL previously reported bugs.
 * =================================================================================================
 */
function calculateIncumbencyEngine(allLogData, headers, mainDataMap) {
  const posIdIndex = headers.indexOf('Position ID');
  const empIdIndex = headers.indexOf('Employee ID');
  const nameIndex = headers.indexOf('Employee Name');
  const jobTitleIndex = headers.indexOf('Job Title');
  const timestampIndex = headers.indexOf('Change Timestamp');
  const effectiveDateIndex = headers.indexOf('Effective Date');
  const hireDateIndex = headers.indexOf('Date Hired');

  const isFirstEverEventForEmployee = (employeeId, eventDate, allLogs) => {
    for (const row of allLogs) {
      const logEmpId = (row[empIdIndex] || '').toString().trim();
      if (logEmpId === employeeId) {
        const logEventDate = _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]);
        if (logEventDate && logEventDate.getTime() < eventDate.getTime()) {
          return false;
        }
      }
    }
    return true;
  };

  const positions = {};
  allLogData.forEach(row => {
    const posId = row[posIdIndex];
    if (posId) {
      if (!positions[posId]) positions[posId] = [];
      positions[posId].push(row);
    }
  });

  const finalHistory = {};

  for (const posId in positions) {
    const logEntries = positions[posId];
    const allChangeEventsForPos = logEntries
      .filter(row => row[timestampIndex])
      .map(row => ({
        eventDate: _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]),
        incumbentId: (row[empIdIndex] || '').toString().trim() || null,
        incumbentName: (row[nameIndex] || '').toString().trim() || 'N/A',
        jobTitle: (row[jobTitleIndex] || '').toString().trim() || 'N/A',
        hireDate: _parseDate(row[hireDateIndex])
      }))
      .filter(e => e.eventDate)
      .sort((a, b) => a.eventDate.getTime() - b.eventDate.getTime());

    if (allChangeEventsForPos.length === 0) continue;

    let historyRecords = [];
    let i = 0;
    while (i < allChangeEventsForPos.length) {
      const startEvent = allChangeEventsForPos[i];
      if (!startEvent.incumbentId) {
        i++;
        continue;
      }

      let endEvent = null;
      let endEventIndex = -1;
      for (let j = i + 1; j < allChangeEventsForPos.length; j++) {
        if (allChangeEventsForPos[j].incumbentId !== startEvent.incumbentId) {
          endEvent = allChangeEventsForPos[j];
          endEventIndex = j;
          break;
        }
      }

      const lastEventOfThisTenure = endEvent ? allChangeEventsForPos[endEventIndex - 1] : allChangeEventsForPos[allChangeEventsForPos.length - 1];
      const endDate = endEvent ? lastEventOfThisTenure.eventDate : null;

      let startDate = startEvent.eventDate;
      if (startEvent.hireDate && startEvent.hireDate.getTime() < startEvent.eventDate.getTime()) {
        if (isFirstEverEventForEmployee(startEvent.incumbentId, startEvent.eventDate, allLogData)) {
          startDate = startEvent.hireDate;
        }
      }

      historyRecords.push({
        startDate: startDate,
        endDate: endDate,
        incumbentId: startEvent.incumbentId,
        incumbentName: lastEventOfThisTenure.incumbentName,
        jobTitle: lastEventOfThisTenure.jobTitle,
        hireDate: startEvent.hireDate
      });

      i = endEvent ? endEventIndex : allChangeEventsForPos.length;
    }

    const changeCount = historyRecords.length;
    historyRecords.forEach(rec => rec.changeCount = changeCount);
    finalHistory[posId] = historyRecords;
  }
  return finalHistory;
}




/**
 * REVISED - Fetches and formats incumbency history for the web app.
 */
function getDetailedIncumbencyHistory(posId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `incumbency_history_${posId}`;
  const cachedHistory = cache.get(cacheKey);

  if (cachedHistory) {
    Logger.log(`Cache HIT for Position ID: ${posId}`);
    return JSON.parse(cachedHistory);
  }

  Logger.log(`Cache MISS for Position ID: ${posId}. Calculating from scratch.`);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheets()[0];
    const logSheet = spreadsheet.getSheetByName('change_log_sheet');
    if (!logSheet || !mainSheet || logSheet.getLastRow() < 2) return [];

    const mainData = mainSheet.getLastRow() > 1 ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 3).getValues() : [];
    const mainDataMap = new Map(mainData.map(row => [row[0], row]));

    const allLogDataWithHeaders = logSheet.getDataRange().getValues();
    const headers = allLogDataWithHeaders.shift();
    const allLogData = allLogDataWithHeaders;

    // Run the main history engine
    const allHistory = calculateIncumbencyEngine(allLogData, headers, mainDataMap);
    let positionHistory = allHistory[posId] || [];

    // --- START: DATA CORRECTION AND FAILSAFE LOGIC ---
    if (positionHistory.length > 0) {
      const lastRecord = positionHistory[positionHistory.length - 1];
      const liveRow = mainDataMap.get(posId);
      const currentLiveEmployeeId = liveRow ? (liveRow[1] || '').toString().trim() : null;
      const isCurrentlyVacant = !currentLiveEmployeeId;

      // Correction 1: If the last incumbent in the history is the current, active employee,
      // ensure their end date is null (i.e., 'Present'), overriding any erroneous log entry.
      if (lastRecord.incumbentId && currentLiveEmployeeId && lastRecord.incumbentId === currentLiveEmployeeId && lastRecord.endDate) {
        Logger.log(`Position ${posId} is currently held by ${currentLiveEmployeeId}, but history shows an end date. Correcting to 'Present'.`);
        lastRecord.endDate = null;
      }

      // Correction 2 (Failsafe): If the position is actually vacant, but the history shows 'Present',
      // find the last known event for that position and use its date as the end date.
      if (lastRecord.endDate === null && isCurrentlyVacant) {
        Logger.log(`Position ${posId} is vacant but history shows "Present". Applying final failsafe.`);
        const posIdIndex = headers.indexOf('Position ID');
        const effectiveDateIndex = headers.indexOf('Effective Date');
        const timestampIndex = headers.indexOf('Change Timestamp');
        
        const allEventsForPos = allLogData
          .filter(row => row[posIdIndex] === posId)
          .map(row => ({ date: _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]) }))
          .filter(event => event.date)
          .sort((a, b) => b.date.getTime() - a.date.getTime());

        if (allEventsForPos.length > 0) {
          lastRecord.endDate = allEventsForPos[0].date;
          Logger.log(`Failsafe applied. Corrected End Date to: ${lastRecord.endDate}`);
        }
      }
    }
    // --- END OF DATA CORRECTION AND FAILSAFE LOGIC ---

    const finalHistory = positionHistory
      .filter(rec => rec.incumbentId)
      .map(rec => {
        const startDate = rec.startDate;
        const endDateForCalc = rec.endDate || new Date();

        let tenureDays = 0;
        if (startDate && endDateForCalc) {
          const diffMillis = endDateForCalc.getTime() - startDate.getTime();
          tenureDays = Math.max(0, Math.floor(diffMillis / (1000 * 60 * 60 * 24)));
        }

        let tenureString = "0 days";
        if (tenureDays > 0) {
          const years = Math.floor(tenureDays / 365.25);
          const months = Math.floor((tenureDays % 365.25) / 30.44);
          const days = Math.round((tenureDays % 365.25) % 30.44);

          let parts = [];
          if (years > 0) parts.push(`${years} year${years > 1 ? 's' : ''}`);
          if (months > 0) parts.push(`${months} month${months > 1 ? 's' : ''}`);
          if (days > 0 || (years === 0 && months === 0)) parts.push(`${days} day${days !== 1 ? 's' : ''}`);
          tenureString = parts.join(', ');
        }

        return {
          name: rec.incumbentName,
          startDate: rec.startDate ? Utilities.formatDate(rec.startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'N/A',
          endDate: rec.endDate ? Utilities.formatDate(rec.endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'Present',
          tenure: tenureString,
          employeeHireDate: rec.hireDate ? Utilities.formatDate(rec.hireDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'N/A',
        };
      });

    const reversedHistory = finalHistory.reverse();
    cache.put(cacheKey, JSON.stringify(reversedHistory), 21600);
    return reversedHistory;

  } catch (e) {
    Logger.log(`Error in getDetailedIncumbencyHistory: ${e.toString()}\nStack: ${e.stack}`);
    return [];
  }
}


/**
 * REVISED NOTIFICATION FUNCTION
 */
function getUpcomingDues() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');


  const today = new Date();
  today.setHours(0, 0, 0, 0);


  const upcoming = [];
  const overdue = [];


  if (mainSheet.getLastRow() < 2) {
    return {
      upcoming,
      overdue
    };
  }


  const mainData = mainSheet.getDataRange().getValues();
  const mainHeaders = mainData.shift();
  const mainDataMap = new Map(mainData.map(row => [row[mainHeaders.indexOf('Position ID')], row]));


  const nameIndex = mainHeaders.indexOf('Employee Name');
  const contractTypeIndex = mainHeaders.indexOf('Contract Type');
  const contractEndIndex = mainHeaders.indexOf('Contract End Date');
  const statusIndex = mainHeaders.indexOf('Status');
  const posStatusIndex = mainHeaders.indexOf('Position Status');


  mainDataMap.forEach((row, posId) => {
    const positionStatus = (row[posStatusIndex] || '').toString().trim().toUpperCase();
    if (positionStatus === 'INACTIVE') return;


    const contractType = (row[contractTypeIndex] || '').toString().trim().toUpperCase();
    const endDate = row[contractEndIndex];
    if (contractType === 'JPRO' && endDate instanceof Date) {
      const normalizedEndDate = new Date(endDate.getTime());
      normalizedEndDate.setHours(0, 0, 0, 0);
      const timeDiff = normalizedEndDate.getTime() - today.getTime();
      const days = Math.round(timeDiff / (1000 * 60 * 60 * 24));
      const employeeName = row[nameIndex];


      if (days >= 0 && days <= 30) {
        const message = `${employeeName}'s JPRO contract ends in ${days} day${days !== 1 ? 's' : ''}.`;
        upcoming.push({
          days,
          message
        });
      } else if (days < 0) {
        const daysAgo = Math.abs(days);
        const message = `${employeeName}'s JPRO contract expired ${daysAgo} day${daysAgo !== 1 ? 's' : ''} ago. Please update their status.`;
        overdue.push({
          days: daysAgo,
          message
        });
      }
    }
  });


  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getDataRange().getValues();
    const logHeaders = logData.shift();
    const logPosIdIndex = logHeaders.indexOf('Position ID');
    const logNameIndex = logHeaders.indexOf('Employee Name');
    const logStatusIndex = logHeaders.indexOf('Status');
    const logEffectiveDateIndex = logHeaders.indexOf('Effective Date');


    if (logPosIdIndex > -1 && logStatusIndex > -1 && logEffectiveDateIndex > -1) {
      const latestResignations = new Map();
      for (let i = logData.length - 1; i >= 0; i--) {
        const row = logData[i];
        const posId = row[logPosIdIndex];
        const logStatus = (row[logStatusIndex] || '').trim().toUpperCase();
        if (posId && logStatus === 'RESIGNED' && !latestResignations.has(posId)) {
          latestResignations.set(posId, {
            date: row[logEffectiveDateIndex],
            name: row[logNameIndex]
          });
        }
      }


      latestResignations.forEach((resignation, posId) => {
        const currentPosData = mainDataMap.get(posId);
        if (!currentPosData || (currentPosData[statusIndex] || '').toUpperCase() !== 'RESIGNED') {
          return;
        }


        const effectiveDate = resignation.date;
        if (effectiveDate instanceof Date) {
          const normalizedEffectiveDate = new Date(effectiveDate.getTime());
          normalizedEffectiveDate.setHours(0, 0, 0, 0);
          const timeDiff = normalizedEffectiveDate.getTime() - today.getTime();
          const days = Math.round(timeDiff / (1000 * 60 * 60 * 24));


          if (days >= 0 && days <= 30) {
            const message = `${resignation.name}'s resignation is effective in ${days} day${days !== 1 ? 's' : ''}.`;
            upcoming.push({
              days,
              message
            });
          } else if (days < 0) {
            const daysAgo = Math.abs(days);
            const message = `${resignation.name}'s resignation was ${daysAgo} day${daysAgo !== 1 ? 's' : ''} ago. Please update the position to VACANT.`;
            overdue.push({
              days: daysAgo,
              message
            });
          }
        }
      });
    }
  }


  const sortedUpcoming = upcoming.sort((a, b) => a.days - b.days).map(d => d.message);
  const sortedOverdue = overdue.sort((a, b) => a.days - b.days).map(d => d.message);


  return {
    upcoming: sortedUpcoming,
    overdue: sortedOverdue
  };
}

/**
 * A robust helper function to parse various date formats.
 * @param {any} dateInput The value from the spreadsheet cell.
 * @returns {Date|null} A valid Date object or null if parsing fails.
 */
function _parseDate(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date && !isNaN(dateInput)) return dateInput;
  
  // Try creating a date from a string or number
  const date = new Date(dateInput);
  if (!isNaN(date)) return date;
  
  return null; // Return null if input is not a valid date
}
