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
    .addItem('Initialize Change Log Sheet', 'initializeChangeLogSheet')
    .addSeparator()
    .addItem('Update Headcount Summary & Create Approval Records', 'takeHeadcountSnapshotWithAlert')
    .addItem('Generate Incumbency History Report', 'generateIncumbencyReport')
    .addItem('Generate Masterlist Export', 'generateMasterlistSheetWithPrompt')
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


function initializeChangeLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'change_log_sheet';
  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert(`A sheet named "${sheetName}" already exists.`);
    return;
  }
  
  const sheet = ss.insertSheet(sheetName);
  const headers = [
    "Position ID", "Employee ID", "Employee Name", "Reporting to ID", "Reporting to", 
    "Job Title", "Division", "Group", "Department", "Section", "Gender", 
    "Level", "Payroll Type", "Job Level", "Contract Type", "Competency", 
    "Status", "Position Status", "Date Hired", "Contract End Date",
    "Change Timestamp", "Change Type", "Transfer Note", "Effective Date",
    "Division Headcount", "Department Headcount", "Section Headcount"
  ];
  
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert(`Successfully created the "${sheetName}" sheet.`);
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
        // Standardize the key: lowercase, remove spaces and special chars.
        // This handles "Reason for Leaving" -> "reasonforleaving"
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
    const keyMap = {};
    headers.forEach((header, i) => {
      const key = header.toLowerCase().replace(/\s+/g, '').replace(/[^a-z0-9]/gi, '');
      keyMap[key] = i;
    });

    for (const key in dataObject) {
      if (typeof dataObject[key] === 'string') {
        dataObject[key] = dataObject[key].toUpperCase();
      }
    }

    // --- Flags and variables for secondary actions ---
    let isTransfer = false;
    let oldPositionIdForTransfer = null;
    let transferredEmployeeId = null;

    let isManualVacate = false;
    let vacatingEmployeeId = null;
    let vacatedPositionId = null;

    let isFillingVacancy = false;
    let filledVacancyPositionId = null;
    let newEmployeeIdForVacancy = null;
    let newEmployeeNameForVacancy = null;

    // --- Promotion/Transfer Logic (Primary Change) ---
    if (dataObject.employeeid && (dataObject.status.toUpperCase() === 'PROMOTION' || dataObject.status.toUpperCase() === 'INTERNAL TRANSFER' || dataObject.status.toUpperCase() === 'LATERAL TRANSFER')) {
      const allData = mainSheet.getDataRange().getValues();
      const posIdIndex = headers.indexOf('Position ID');
      const empIdIndex = headers.indexOf('Employee ID');
      for (let i = 1; i < allData.length; i++) {
        const row = allData[i];
        if ((String(row[empIdIndex]) || '').toUpperCase() === dataObject.employeeid.toUpperCase() && (String(row[posIdIndex]) || '').toUpperCase() !== dataObject.positionid.toUpperCase()) {
          const oldRowIndex = i + 1;
          isTransfer = true;
          oldPositionIdForTransfer = row[posIdIndex];
          transferredEmployeeId = dataObject.employeeid.toUpperCase();

          if (dataObject.startdateinposition) {
            scriptProperties.setProperties({
              'pendingResignationPosId': oldPositionIdForTransfer.toUpperCase(),
              'pendingResignationDate': dataObject.startdateinposition
            });
          }
          mainSheet.getRange(oldRowIndex, keyMap['employeeid'] + 1).clearContent();
          mainSheet.getRange(oldRowIndex, keyMap['employeename'] + 1).clearContent();
          mainSheet.getRange(oldRowIndex, keyMap['gender'] + 1).clearContent();
          mainSheet.getRange(oldRowIndex, keyMap['datehired'] + 1).clearContent();
          mainSheet.getRange(oldRowIndex, keyMap['contractenddate'] + 1).clearContent();
          mainSheet.getRange(oldRowIndex, keyMap['status'] + 1).setValue('VACANT');
          break;
        }
      }
    }

    if ((dataObject.status === 'VACANT' || dataObject.status === 'RESIGNED') && dataObject.effectivedate) {
      PropertiesService.getScriptProperties().setProperties({
        'pendingEffectiveDatePosId': dataObject.positionid.toUpperCase(),
        'pendingEffectiveDate': dataObject.effectivedate
      });
    }
    if (dataObject.startdateinposition) {
      PropertiesService.getScriptProperties().setProperty('overrideTimestamp', dataObject.startdateinposition);
    }

    // --- Main Add/Edit Logic (Primary Change) ---
    if (mode === 'add') {
      const newRowData = Array(headers.length).fill('');
      for (const key in dataObject) {
        if (keyMap.hasOwnProperty(key)) newRowData[keyMap[key]] = dataObject[key];
      }
      const newPositionId = dataObject.positionid;
      let insertRowIndex = -1;
      if (newPositionId) {
        const positionIdPrefix = newPositionId.substring(0, newPositionId.lastIndexOf('-'));
        if (positionIdPrefix) {
          const positionIdColValues = mainSheet.getRange("A1:A").getValues().flat();
          for (let i = positionIdColValues.length - 1; i > 0; i--) {
            const currentPositionId = positionIdColValues[i];
            if (currentPositionId && currentPositionId.startsWith(positionIdPrefix)) {
              insertRowIndex = i + 1;
              break;
            }
          }
        }
      }
      if (insertRowIndex !== -1) {
        mainSheet.insertRowAfter(insertRowIndex);
        mainSheet.getRange(insertRowIndex + 1, 1, 1, newRowData.length).setValues([newRowData]);
      } else {
        mainSheet.appendRow(newRowData);
      }
    } else if (mode === 'edit') {
      const positionId = dataObject.positionid;
      const positionIdColValues = mainSheet.getRange("A2:A").getValues();
      const rowIndex = positionIdColValues.findIndex(r => r[0] == positionId) + 2;
      if (rowIndex === 1) throw new Error(`Position ID ${positionId} not found for editing.`);

      const rangeToUpdate = mainSheet.getRange(rowIndex, 1, 1, headers.length);
      const existingRowData = rangeToUpdate.getValues()[0];
      const statusColIndex = keyMap['status'];
      const originalStatus = existingRowData[statusColIndex];
      const originalEmployeeId = existingRowData[keyMap['employeeid']];

      // Check for manual vacate scenario
      if (dataObject.status.toUpperCase() === 'VACANT' && originalEmployeeId) {
        isManualVacate = true;
        vacatingEmployeeId = originalEmployeeId.toUpperCase();
        vacatedPositionId = positionId;
      }

      // Check for filling vacancy scenario
      if (originalStatus && originalStatus.toUpperCase() === 'VACANT' && dataObject.status && dataObject.status.toUpperCase() !== 'VACANT') {
        isFillingVacancy = true;
        filledVacancyPositionId = positionId;
        newEmployeeIdForVacancy = dataObject.employeeid;
        newEmployeeNameForVacancy = dataObject.employeename;
      }

      // Ghost tenure fix
      if (originalStatus && originalStatus.toUpperCase() === 'VACANT') {
        const isFillingAction = dataObject.status && dataObject.status.toUpperCase() !== 'VACANT' && dataObject.employeeid;
        if (!isFillingAction) {
          dataObject.employeeid = ''; dataObject.employeename = ''; dataObject.gender = '';
          dataObject.datehired = ''; dataObject.contractenddate = ''; dataObject.status = 'VACANT';
        }
      }

      // Apply primary data changes from the form
      for (const key in dataObject) {
        if (keyMap.hasOwnProperty(key)) {
          existingRowData[keyMap[key]] = dataObject[key];
        }
      }
      rangeToUpdate.setValues([existingRowData]);

      // --- NEW: Resignation Data Logging ---
      if (dataObject.status === 'RESIGNED') {
        let resignationSheet = ss.getSheetByName('Resignation Data');
        if (!resignationSheet) {
          resignationSheet = ss.insertSheet('Resignation Data');
          const resignationHeaders = [
            'Timestamp', 'Position ID', 'Employee ID', 'Employee Name',
            'Division', 'Group', 'Department', 'Section', 'Job Title',
            'Level', 'Job Level', 'Gender', 'Contract Type', 'Date Hired', 'Resignation Date', 'Reason for Leaving'
          ];
          resignationSheet.appendRow(resignationHeaders);
          resignationSheet.setFrozenRows(1);
        }

        const resignedEmployeeData = [
          new Date(),
          dataObject.positionid,
          dataObject.employeeid,
          dataObject.employeename,
          dataObject.division,
          dataObject.group,
          dataObject.department,
          dataObject.section,
          dataObject.jobtitle,
          dataObject.level,
          dataObject.joblevel,
          dataObject.gender,
          dataObject.contracttype,
          dataObject.datehired,
          dataObject.effectivedate,
          dataObject.reasonforleaving || ''
        ];
        resignationSheet.appendRow(resignedEmployeeData);
      }
    }

    // --- STEP 1: Log only the primary changes ---
    SpreadsheetApp.flush();
    logDataChanges();

    // --- STEP 2: Perform secondary, un-logged actions ---
    let secondaryChangesMade = false;
    const allDataForSecondary = mainSheet.getDataRange().getValues(); // Get fresh data
    const reportingToIdIndex = keyMap['reportingtoid'];
    const reportingToNameIndex = keyMap['reportingto'];

    // Action A: Re-assign subordinates from a transfer
    if (isTransfer && reportingToIdIndex !== undefined) {
      for (let i = 1; i < allDataForSecondary.length; i++) {
        if ((String(allDataForSecondary[i][reportingToIdIndex]) || '').toUpperCase() === transferredEmployeeId) {
          mainSheet.getRange(i + 1, reportingToIdIndex + 1).setValue(oldPositionIdForTransfer);
          mainSheet.getRange(i + 1, reportingToNameIndex + 1).clearContent();
          secondaryChangesMade = true;
        }
      }
    }

    // Action B: Re-assign subordinates from a manual vacate
    if (isManualVacate && reportingToIdIndex !== undefined) {
      for (let i = 1; i < allDataForSecondary.length; i++) {
        if ((String(allDataForSecondary[i][reportingToIdIndex]) || '').toUpperCase() === vacatingEmployeeId) {
          mainSheet.getRange(i + 1, reportingToIdIndex + 1).setValue(vacatedPositionId);
          mainSheet.getRange(i + 1, reportingToNameIndex + 1).clearContent();
          secondaryChangesMade = true;
        }
      }
    }

    // Action C: Re-assign subordinates when filling a vacancy
    if (isFillingVacancy && reportingToIdIndex !== undefined) {
      for (let i = 1; i < allDataForSecondary.length; i++) {
        if (allDataForSecondary[i][reportingToIdIndex] === filledVacancyPositionId) {
          mainSheet.getRange(i + 1, reportingToIdIndex + 1).setValue(newEmployeeIdForVacancy);
          mainSheet.getRange(i + 1, reportingToNameIndex + 1).setValue(newEmployeeNameForVacancy);
          secondaryChangesMade = true;
        }
      }
    }

    // --- STEP 3: Manually sync the data state if secondary changes were made ---
    if (secondaryChangesMade) {
      SpreadsheetApp.flush();
      const finalData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
      scriptProperties.setProperty('lastKnownData', JSON.stringify(finalData));
    }

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


  const allLogDataNoHeaders = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();

  for (const posId of sortedPosIds) {
    const records = allHistory[posId];
    if (!records || records.length === 0) continue;

    const lastRecord = records[records.length - 1];
    const currentLiveRow = mainDataMap.get(posId);
    const currentLiveEmployeeId = currentLiveRow ? (currentLiveRow[1] || '').toString().trim() : null;
    const isCurrentlyVacant = !currentLiveEmployeeId;

    // Correction 1: If the last incumbent is the current employee, ensure end date is 'Present'.
    if (lastRecord.incumbentId && currentLiveEmployeeId && lastRecord.incumbentId === currentLiveEmployeeId && lastRecord.endDate) {
      lastRecord.endDate = null;
    }

    // Correction 2 (Failsafe): If history says 'Present' but the position is vacant, find the true end date.
    if (lastRecord.endDate === null && isCurrentlyVacant) {
      const posIdIndex = headers.indexOf('Position ID');
      const effectiveDateIndex = headers.indexOf('Effective Date');
      const timestampIndex = headers.indexOf('Change Timestamp');

      const allEventsForPos = allLogDataNoHeaders
        .filter(row => row[posIdIndex] === posId)
        .map(row => ({ date: _parseDate(row[effectiveDateIndex]) || _parseDate(row[timestampIndex]) }))
        .filter(event => event.date)
        .sort((a, b) => b.date.getTime() - a.date.getTime());

      if (allEventsForPos.length > 0) {
        lastRecord.endDate = allEventsForPos[0].date;
      }
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
 * FINAL REWRITE v9 - calculateIncumbencyEngine
 * =================================================================================================
 * This version correctly identifies the end of a tenure by recognizing "effective-dated" events
 * (like promotions/resignations) as definitive termination points. It also correctly handles
 * subsequent "ghost" log entries that might occur for an employee after their tenure has
 * officially ended, preventing these from creating incorrect history records.
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
        hireDate: _parseDate(row[hireDateIndex]),
        isEffective: !!_parseDate(row[effectiveDateIndex])
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

      let startDate = startEvent.eventDate;
      if (startEvent.hireDate && startEvent.hireDate.getTime() < startEvent.eventDate.getTime()) {
        if (isFirstEverEventForEmployee(startEvent.incumbentId, startEvent.eventDate, allLogData)) {
          startDate = startEvent.hireDate;
        }
      }

      let endDate = null;
      let tenureEndingEvent = null;
      let nextEventIndex = i + 1;

      for (let j = i; j < allChangeEventsForPos.length; j++) {
        const currentEvent = allChangeEventsForPos[j];

        if (currentEvent.incumbentId !== startEvent.incumbentId) {
          endDate = currentEvent.eventDate;
          tenureEndingEvent = currentEvent;
          nextEventIndex = j;
          break;
        }

        if (currentEvent.isEffective && currentEvent.incumbentId === startEvent.incumbentId) {
          endDate = currentEvent.eventDate;
          tenureEndingEvent = currentEvent;
          let k = j + 1;
          while (k < allChangeEventsForPos.length && allChangeEventsForPos[k].incumbentId === startEvent.incumbentId) {
            k++;
          }
          nextEventIndex = k;
          break;
        }
      }

      if (!tenureEndingEvent) {
        nextEventIndex = allChangeEventsForPos.length;
      }

      const lastEventOfThisTenure = allChangeEventsForPos[nextEventIndex - 1];

      historyRecords.push({
        startDate: startDate,
        endDate: endDate,
        incumbentId: startEvent.incumbentId,
        incumbentName: lastEventOfThisTenure.incumbentName,
        jobTitle: lastEventOfThisTenure.jobTitle,
        hireDate: startEvent.hireDate
      });

      i = nextEventIndex;
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

function getJobDescription(positionId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const jdSheetName = 'Job Descriptions';
    let jdSheet = spreadsheet.getSheetByName(jdSheetName);

    if (!jdSheet) {
      jdSheet = spreadsheet.insertSheet(jdSheetName);
      jdSheet.appendRow(['Position ID', 'Description']);
      jdSheet.setFrozenRows(1);
      return null;
    }

    if (jdSheet.getLastRow() < 2) {
      return null;
    }

    const data = jdSheet.getRange(2, 1, jdSheet.getLastRow() - 1, 2).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === positionId.toString().trim()) {
        return data[i][1] || 'No description provided.';
      }
    }
    
    return null;

  } catch (e) {
    Logger.log(`Error in getJobDescription: ${e.toString()}`);
    throw new Error(`Could not retrieve job description. Error: ${e.message}`);
  }
}

function getResignationData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const resignationSheet = ss.getSheetByName('Resignation Data');

  const emptyResult = { reasonCounts: {}, resignationGenderCounts: {}, resignationContractCounts: {}, resignationDivisionCounts: {}, resignationJobGroupCounts: {}, monthlyTurnover: [], yearlyHiresLeavers: { hires: 0, leavers: 0 }, ytdTurnover: 0, overallHeadcount: 0, filteredResignationsCount: 0 };

  if (!mainSheet || mainSheet.getLastRow() < 2) return emptyResult;

  const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const dateHiredIndex = mainHeaders.indexOf('Date Hired');
  const empIdIndex = mainHeaders.indexOf('Employee ID');
  const overallHeadcount = mainData.filter(row => row[empIdIndex]).length;

  if (!resignationSheet || resignationSheet.getLastRow() < 2) {
    return { ...emptyResult, overallHeadcount };
  }

  const resignationData = resignationSheet.getRange(2, 1, resignationSheet.getLastRow() - 1, resignationSheet.getLastColumn()).getValues();
  const resignationHeaders = resignationSheet.getRange(1, 1, 1, resignationSheet.getLastColumn()).getValues()[0];
  const resHeaderMap = new Map(resignationHeaders.map((h, i) => [h.trim(), i]));

  const jobGroupMapping = { 1: 'Executives', 2: 'Director', 3: 'Managerial', 4: 'Supervisory', 5: 'Rank & File', 6: 'Jobcon' };
  const selectedYear = (filters.year && filters.year !== 'All Years') ? parseInt(filters.year) : new Date().getFullYear();

  const filteredResignations = resignationData.filter(row => {
    const resDate = new Date(row[resHeaderMap.get('Resignation Date')]);
    if (filters.year && !String(filters.year).toLowerCase().startsWith('all') && resDate.getFullYear() != filters.year) return false;
    if (filters.month && !String(filters.month).toLowerCase().startsWith('all')) {
        const monthIndex = new Date(Date.parse(filters.month +" 1, 2012")).getMonth();
        if (resDate.getMonth() != monthIndex) return false;
    }
    // Apply other filters
    if (filters.division && !String(filters.division).toLowerCase().startsWith('all') && row[resHeaderMap.get('Division')] !== filters.division) return false;
    if (filters.group && !String(filters.group).toLowerCase().startsWith('all') && row[resHeaderMap.get('Group')] !== filters.group) return false;
    if (filters.department && !String(filters.department).toLowerCase().startsWith('all') && row[resHeaderMap.get('Department')] !== filters.department) return false;
    if (filters.section && !String(filters.section).toLowerCase().startsWith('all') && row[resHeaderMap.get('Section')] !== filters.section) return false;
    if (filters.jobLevel && !String(filters.jobLevel).toLowerCase().startsWith('all') && row[resHeaderMap.get('Job Level')] !== filters.jobLevel) return false;
    if (filters.gender && !String(filters.gender).toLowerCase().startsWith('all') && row[resHeaderMap.get('Gender')] !== filters.gender) return false;
    if (filters.jobTitle && !String(filters.jobTitle).toLowerCase().startsWith('all') && row[resHeaderMap.get('Job Title')] !== filters.jobTitle) return false;
    return true;
  });

  const reasonCounts = {}, resignationGenderCounts = {}, resignationContractCounts = {}, resignationDivisionCounts = {}, resignationJobGroupCounts = {};
  filteredResignations.forEach(row => {
    reasonCounts[row[resHeaderMap.get('Reason for Leaving')] || 'Unknown'] = (reasonCounts[row[resHeaderMap.get('Reason for Leaving')] || 'Unknown'] || 0) + 1;
    resignationGenderCounts[row[resHeaderMap.get('Gender')] || 'Unknown'] = (resignationGenderCounts[row[resHeaderMap.get('Gender')] || 'Unknown'] || 0) + 1;
    resignationContractCounts[row[resHeaderMap.get('Contract Type')] || 'Unknown'] = (resignationContractCounts[row[resHeaderMap.get('Contract Type')] || 'Unknown'] || 0) + 1;
    resignationDivisionCounts[row[resHeaderMap.get('Division')] || 'Unknown'] = (resignationDivisionCounts[row[resHeaderMap.get('Division')] || 'Unknown'] || 0) + 1;
    resignationJobGroupCounts[jobGroupMapping[row[resHeaderMap.get('Level')]] || 'Unknown'] = (resignationJobGroupCounts[jobGroupMapping[row[resHeaderMap.get('Level')]] || 'Unknown'] || 0) + 1;
  });

  const leaversThisYear = resignationData.filter(r => new Date(r[resHeaderMap.get('Resignation Date')]).getFullYear() === selectedYear);
  const hiresThisYear = mainData.filter(r => r[dateHiredIndex] && new Date(r[dateHiredIndex]).getFullYear() === selectedYear);

  const endOfYearHeadcount = mainData.length;
  const startOfYearHeadcount = endOfYearHeadcount - hiresThisYear.length + leaversThisYear.length;
  const averageHeadcount = (startOfYearHeadcount + endOfYearHeadcount) / 2;

  const ytdTurnover = averageHeadcount > 0 ? (leaversThisYear.length / averageHeadcount) * 100 : 0;

  const monthlyTurnover = Array(12).fill(null).map((_, month) => {
    const monthlySeparations = leaversThisYear.filter(r => new Date(r[resHeaderMap.get('Resignation Date')]).getMonth() === month).length;
    const monthlyRate = averageHeadcount > 0 ? (monthlySeparations / averageHeadcount) * 100 : 0;
    return {
      month: new Date(selectedYear, month, 1).toLocaleString('default', { month: 'short' }),
      separations: monthlySeparations,
      rate: parseFloat(monthlyRate.toFixed(2))
    };
  });

  return {
    reasonCounts,
    resignationGenderCounts,
    resignationContractCounts,
    resignationDivisionCounts,
    resignationJobGroupCounts,
    monthlyTurnover,
    yearlyHiresLeavers: { hires: hiresThisYear.length, leavers: leaversThisYear.length },
    ytdTurnover: parseFloat(ytdTurnover.toFixed(2)),
    overallHeadcount,
    filteredResignationsCount: filteredResignations.length
  };
}

function getAnalyticsData(filters) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = spreadsheet.getSheets()[0];
  const previousSheet = spreadsheet.getSheetByName('Previous Headcount');

  let headers = [];
  let mainData = [];
  if (mainSheet.getLastRow() >= 1) {
    headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    if (mainSheet.getLastRow() > 1) {
      mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    }
  } else {
    return { statusCounts: {}, contractCounts: {}, trendData: [], totalHeadcount: 0, overallHeadcount: 0 };
  }

  const empIdIndex = headers.indexOf('Employee ID');
  const overallHeadcount = mainData.filter(row => row[empIdIndex]).length;

  let filteredData = [];
  const statusCounts = {};
  const contractCounts = {};
  let totalHeadcount = 0;

  // 1. Filter main data based on user's selections
  if (mainData.length > 0) {
    const headerMap = {
      division: headers.indexOf('Division'),
      group: headers.indexOf('Group'),
      department: headers.indexOf('Department'),
      section: headers.indexOf('Section'),
      jobLevel: headers.indexOf('Job Level'),
      positionStatus: headers.indexOf('Position Status'),
      gender: headers.indexOf('Gender'),
      jobTitle: headers.indexOf('Job Title'),
      status: headers.indexOf('Status'),
      contractType: headers.indexOf('Contract Type')
    };

    filteredData = mainData.filter(row => {
      return Object.keys(filters).every(key => {
        if (!filters[key] || String(filters[key]).toLowerCase().startsWith('all')) return true;
        const colIndex = headerMap[key];
        return colIndex !== -1 && (row[colIndex] || '').toString().trim() === filters[key];
      });
    });
  }

  // 2. Calculate all analytics metrics from filtered data
  const genderCounts = {};
  const jobGroupCounts = {};
  const losCounts = {
    '< 1 Year': 0, '1-3 Years': 0, '3-5 Years': 0, '5-10 Years': 0, '10+ Years': 0
  };
  
  const jobGroupMapping = {
    1: 'Executives', 2: 'Director', 3: 'Managerial',
    4: 'Supervisory', 5: 'Rank & File', 6: 'Jobcon'
  };

  const statusIndex = headers.indexOf('Status');
  const contractIndex = headers.indexOf('Contract Type');
  const genderIndex = headers.indexOf('Gender');
  const levelIndex = headers.indexOf('Level');
  const hiredIndex = headers.indexOf('Date Hired');
  const today = new Date();

  filteredData.forEach(row => {
    // Only count as "headcount" if the position is filled
    if (row[empIdIndex]) {
      totalHeadcount++;
      
      const status = (row[statusIndex] || 'Unknown').toString().trim();
      statusCounts[status] = (statusCounts[status] || 0) + 1;

      const contract = (row[contractIndex] || 'Unknown').toString().trim();
      contractCounts[contract] = (contractCounts[contract] || 0) + 1;
      
      const gender = (row[genderIndex] || 'Unknown').toString().trim();
      genderCounts[gender] = (genderCounts[gender] || 0) + 1;
      
      const level = row[levelIndex];
      const jobGroup = jobGroupMapping[level] || 'Unknown';
      jobGroupCounts[jobGroup] = (jobGroupCounts[jobGroup] || 0) + 1;

      const hiredDate = row[hiredIndex] ? new Date(row[hiredIndex]) : null;
      if (hiredDate && !isNaN(hiredDate)) {
        const years = (today - hiredDate) / (1000 * 60 * 60 * 24 * 365.25);
        if (years < 1) losCounts['< 1 Year']++;
        else if (years < 3) losCounts['1-3 Years']++;
        else if (years < 5) losCounts['3-5 Years']++;
        else if (years < 10) losCounts['5-10 Years']++;
        else losCounts['10+ Years']++;
      }
    }
  });


  // 3. Headcount Trend Data
  // This section reads the 'Previous Headcount' sheet to build a historical trend.
  // It expects a specific format generated by the 'takeHeadcountSnapshot' function:
  // - Columns are added dynamically for each snapshot, e.g., "Jan 2023 Filled", "Jan 2023 Vacant".
  // - It identifies these columns by looking for the " Filled" suffix.
  // - To avoid double-counting, it only aggregates rows that represent a division total.
  //   These are identified as rows where the 'Group', 'Department', and 'Section' columns (2, 3, 4) are blank.
  const trendData = [];
  if (previousSheet && previousSheet.getLastRow() > 1) {
    const prevData = previousSheet.getDataRange().getValues();
    const prevHeaders = prevData.shift();
    const divIdx = 0, grpIdx = 1, dptIdx = 2, secIdx = 3;

    for (let i = 0; i < prevHeaders.length; i++) {
      const header = prevHeaders[i];
      if (header.includes(' Filled')) {
        const month = header.replace(' Filled', '').trim();
        const vacantHeader = `${month} Vacant`;
        const vacantIndex = prevHeaders.indexOf(vacantHeader);
        
        if (vacantIndex !== -1) {
          let totalFilled = 0;
          let totalVacant = 0;
          
          const sectionFilter = (filters.section && !String(filters.section).toLowerCase().startsWith('all')) ? filters.section : null;
          const departmentFilter = (filters.department && !String(filters.department).toLowerCase().startsWith('all')) ? filters.department : null;
          const groupFilter = (filters.group && !String(filters.group).toLowerCase().startsWith('all')) ? filters.group : null;
          const divisionFilter = (filters.division && !String(filters.division).toLowerCase().startsWith('all')) ? filters.division : null;

          let targetRow = null;

          if (sectionFilter) {
              targetRow = prevData.find(r => r[secIdx] === sectionFilter && r[dptIdx] === departmentFilter && r[grpIdx] === groupFilter && r[divIdx] === divisionFilter);
          } else if (departmentFilter) {
              targetRow = prevData.find(r => r[dptIdx] === departmentFilter && r[grpIdx] === groupFilter && r[divIdx] === divisionFilter && !r[secIdx]);
          } else if (groupFilter) {
              targetRow = prevData.find(r => r[grpIdx] === groupFilter && r[divIdx] === divisionFilter && !r[dptIdx] && !r[secIdx]);
          } else if (divisionFilter) {
              targetRow = prevData.find(r => r[divIdx] === divisionFilter && !r[grpIdx] && !r[dptIdx] && !r[secIdx]);
          }

          if (targetRow) {
              totalFilled = parseInt(targetRow[i] || 0);
              totalVacant = parseInt(targetRow[vacantIndex] || 0);
          } else if (!sectionFilter && !departmentFilter && !groupFilter && !divisionFilter) {
              // If no filters are active, sum up all the division-level totals for a grand total.
              const divisionTotalRows = prevData.filter(r => r[divIdx] && !r[grpIdx] && !r[dptIdx] && !r[secIdx]);
              divisionTotalRows.forEach(row => {
                  totalFilled += parseInt(row[i] || 0);
                  totalVacant += parseInt(row[vacantIndex] || 0);
              });
          }
          
          trendData.push({ month: month, filled: totalFilled, vacant: totalVacant });
        }
      }
    }
  }

  // --- NEW: Calculate New Hires data from Log Sheet ---
  const newHiresByMonth = {};
  const logSheet = spreadsheet.getSheetByName('change_log_sheet');
  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const logHeaderMap = new Map(logHeaders.map((h, i) => [h.trim(), i]));
    
    const logStatusIndex = logHeaderMap.get('Status');
    const logEffectiveDateIndex = logHeaderMap.get('Effective Date');
    const logTimestampIndex = logHeaderMap.get('Change Timestamp');
    
    const twelveMonthsAgo = new Date();
    twelveMonthsAgo.setMonth(twelveMonthsAgo.getMonth() - 12);

    logData.forEach(row => {
      if ((row[logStatusIndex] || '').toUpperCase() === 'NEW HIRE') {
        const eventDate = row[logEffectiveDateIndex] || row[logTimestampIndex];
        if (eventDate && new Date(eventDate) >= twelveMonthsAgo) {
          const date = new Date(eventDate);
          const monthYear = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM yyyy');
          newHiresByMonth[monthYear] = (newHiresByMonth[monthYear] || 0) + 1;
        }
      }
    });
  }

  return {
    statusCounts: statusCounts,
    contractCounts: contractCounts,
    genderCounts: genderCounts,
    jobGroupCounts: jobGroupCounts,
    losCounts: losCounts,
    trendData: trendData,
    totalHeadcount: totalHeadcount,
    overallHeadcount: overallHeadcount,
    newHiresByMonth: newHiresByMonth
  };
}

function getEmployeeMovementData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheets()[0];
  const logSheet = ss.getSheetByName('change_log_sheet');

  const emptyResult = {
    promotionCount: 0,
    transferCount: 0,
    promotionRate: 0,
    transferRate: 0,
    promotionsByDept: {},
    transfersByDept: {}
  };

  if (!logSheet) {
    Logger.log('Warning: change_log_sheet not found. Returning empty data for Employee Movement.');
    return emptyResult;
  }

  if (!mainSheet || mainSheet.getLastRow() < 2 || logSheet.getLastRow() < 2) {
    return emptyResult;
  }

  // Get main data for headcount calculation
  const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  const mainHeaderMap = new Map(mainHeaders.map((h, i) => [h.trim(), i]));

  // Get log data for movement calculation
  const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
  const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const logHeaderMap = new Map(logHeaders.map((h, i) => [h.trim(), i]));

  const statusIndex = logHeaderMap.get('Status');
  const effectiveDateIndex = logHeaderMap.get('Effective Date');
  const timestampIndex = logHeaderMap.get('Change Timestamp');
  const departmentIndex = logHeaderMap.get('Department');

  const selectedYear = (filters.year && !String(filters.year).toLowerCase().startsWith('all')) ? parseInt(filters.year) : new Date().getFullYear();

  // Filter log data based on filters
  const filteredLogData = logData.filter(row => {
    const eventDate = row[effectiveDateIndex] || row[timestampIndex];
    if (!eventDate) return false;

    const date = new Date(eventDate);
    if (filters.year && !String(filters.year).toLowerCase().startsWith('all') && date.getFullYear() != filters.year) return false;
    if (filters.month && !String(filters.month).toLowerCase().startsWith('all')) {
        const monthIndex = new Date(Date.parse(filters.month +" 1, 2012")).getMonth();
        if (date.getMonth() != monthIndex) return false;
    }
    
    // Apply other location filters
    if (filters.division && !String(filters.division).toLowerCase().startsWith('all') && row[logHeaderMap.get('Division')] !== filters.division) return false;
    if (filters.group && !String(filters.group).toLowerCase().startsWith('all') && row[logHeaderMap.get('Group')] !== filters.group) return false;
    if (filters.department && !String(filters.department).toLowerCase().startsWith('all') && row[logHeaderMap.get('Department')] !== filters.department) return false;
    if (filters.section && !String(filters.section).toLowerCase().startsWith('all') && row[logHeaderMap.get('Section')] !== filters.section) return false;
    if (filters.jobLevel && !String(filters.jobLevel).toLowerCase().startsWith('all') && row[logHeaderMap.get('Job Level')] !== filters.jobLevel) return false;
    if (filters.gender && !String(filters.gender).toLowerCase().startsWith('all') && row[logHeaderMap.get('Gender')] !== filters.gender) return false;
    if (filters.jobTitle && !String(filters.jobTitle).toLowerCase().startsWith('all') && row[logHeaderMap.get('Job Title')] !== filters.jobTitle) return false;
    
    return true;
  });

  let promotionCount = 0;
  let transferCount = 0;
  const promotionsByDept = {};
  const transfersByDept = {};

  filteredLogData.forEach(row => {
    const status = (row[statusIndex] || '').toUpperCase();
    const department = row[departmentIndex] || 'Unknown';

    if (status === 'PROMOTION') {
      promotionCount++;
      promotionsByDept[department] = (promotionsByDept[department] || 0) + 1;
    } else if (status === 'INTERNAL TRANSFER' || status === 'LATERAL TRANSFER') {
      transferCount++;
      transfersByDept[department] = (transfersByDept[department] || 0) + 1;
    }
  });

  // Calculate Rates (similar to resignation rate)
  const leaversThisYear = logData.filter(r => (r[logHeaderMap.get('Status')] || '').toUpperCase() === 'RESIGNED' && new Date(r[effectiveDateIndex] || r[timestampIndex]).getFullYear() === selectedYear);
  const hiresThisYear = mainData.filter(r => r[mainHeaderMap.get('Date Hired')] && new Date(r[mainHeaderMap.get('Date Hired')]).getFullYear() === selectedYear);

  const endOfYearHeadcount = mainData.filter(r => r[mainHeaderMap.get('Employee ID')]).length;
  const startOfYearHeadcount = endOfYearHeadcount - hiresThisYear.length + leaversThisYear.length;
  const averageHeadcount = (startOfYearHeadcount + endOfYearHeadcount) / 2;

  const promotionRate = averageHeadcount > 0 ? (promotionCount / averageHeadcount) * 100 : 0;
  const transferRate = averageHeadcount > 0 ? (transferCount / averageHeadcount) * 100 : 0;

  return {
    promotionCount,
    transferCount,
    promotionRate: parseFloat(promotionRate.toFixed(2)),
    transferRate: parseFloat(transferRate.toFixed(2)),
    promotionsByDept,
    transfersByDept
  };
}


function generateMasterlistSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheets()[0];

    if (mainSheet.getLastRow() < 2) {
      throw new Error("The source data sheet is empty.");
    }

    const data = mainSheet.getDataRange().getValues();
    
    const newSpreadsheet = SpreadsheetApp.create(`Employee Masterlist - ${new Date().toLocaleDateString()}`);
    const newSheet = newSpreadsheet.getSheets()[0];
    
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    newSheet.setFrozenRows(1);
    newSheet.autoResizeColumns(1, data[0].length);
    
    return newSpreadsheet.getUrl();

  } catch (e) {
    Logger.log(`Error in generateMasterlistSheet: ${e.toString()}`);
    throw new Error(`Failed to generate masterlist. Error: ${e.message}`);
  }
}

function generateMasterlistSheetWithPrompt() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = generateMasterlistSheet();
    const htmlOutput = HtmlService.createHtmlOutput(`<p>Masterlist generated successfully. <a href="${url}" target="_blank">Click here to open the new spreadsheet.</a></p>`)
        .setWidth(400)
        .setHeight(100);
    ui.showModalDialog(htmlOutput, 'Masterlist Generated');
  } catch (e) {
    ui.alert('Error', `Failed to generate masterlist: ${e.message}`, ui.ButtonSet.OK);
  }
}

function saveJobDescription(positionId, description) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const jdSheetName = 'Job Descriptions';
    let jdSheet = spreadsheet.getSheetByName(jdSheetName);

    if (!jdSheet) {
      jdSheet = spreadsheet.insertSheet(jdSheetName);
      jdSheet.appendRow(['Position ID', 'Description']);
      jdSheet.setFrozenRows(1);
    }
    
    const data = jdSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === positionId.toString().trim()) {
        jdSheet.getRange(i + 1, 2).setValue(description);
        found = true;
        break;
      }
    }

    if (!found) {
      jdSheet.appendRow([positionId, description]);
    }
    
    return "Job description saved successfully.";

  } catch (e) {
    Logger.log(`Error in saveJobDescription: ${e.toString()}`);
    throw new Error(`Could not save job description. Error: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}
