/**
 * 📌 On open: Build dynamic custom menu from Settings
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menuName = getSetting('MenuName') || 'Bulk Admin Tools';

  ui.createMenu(menuName)
    .addItem('Clear Data', 'clearData')
    .addItem('Apply Header Formatting', 'applyHeaderFormatting')
    .addItem('Duplicate Template', 'duplicateTemplate')
    .addItem('Archive Data', 'archiveData')
    .addItem('Protect Header Row', 'protectHeaderRow')
    .addToUi();
}

/**
 * 📌 Get value from Settings sheet by Parameter name
 */
function getSetting(param) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) throw new Error('Settings sheet not found.');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === param) {
      return data[i][1];
    }
  }
  return '';
}

/**
 * 📌 Clear all data except header rows in active sheet
 */
function clearData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headerRows = parseInt(getSetting('ProtectRangeRows') || '1');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow > headerRows) {
    sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol).clearContent();
  }
  SpreadsheetApp.getUi().alert('Data cleared (except header rows).');
}

/**
 * 📌 Apply standard header formatting in active sheet
 */
function applyHeaderFormatting() {
  const sheet = SpreadsheetApp.getActiveSheet();
  let headerRows = parseInt(getSetting('ProtectRangeRows') || '0');

  if (!headerRows || headerRows < 1) {
    const response = SpreadsheetApp.getUi().prompt(
      'Header Rows',
      'Enter the number of header rows to format:',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      SpreadsheetApp.getUi().alert('Operation cancelled.');
      return;
    }
    headerRows = parseInt(response.getResponseText()) || 1;
  }

  const header = sheet.getRange(1, 1, headerRows, sheet.getLastColumn());
  header.setFontWeight('bold').setBackground('#f1f1f1');

  SpreadsheetApp.getUi().alert(`Formatting applied to top ${headerRows} row(s).`);
}

/**
 * 📌 Duplicate the template sheet with timestamp
 */
function duplicateTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let templateName = getSetting('TemplateSheetName');

  if (!templateName) {
    const sheets = ss.getSheets().map(sheet => sheet.getName()).join('\n');
    const response = SpreadsheetApp.getUi().prompt(
      'Template Sheet Name',
      `Enter the name of the sheet to duplicate.\n\nAvailable sheets:\n${sheets}`,
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      SpreadsheetApp.getUi().alert('Operation cancelled.');
      return;
    }

    templateName = response.getResponseText().trim();
    if (!templateName) {
      SpreadsheetApp.getUi().alert('No sheet name entered. Operation cancelled.');
      return;
    }
  }

  const template = ss.getSheetByName(templateName);
  if (!template) {
    SpreadsheetApp.getUi().alert(`Template sheet "${templateName}" not found.`);
    return;
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const newSheetName = `${templateName}_${timestamp}`;
  template.copyTo(ss).setName(newSheetName);

  SpreadsheetApp.getUi().alert(`Template duplicated as "${newSheetName}".`);
}

/**
 * 📌 Archive current sheet data to archive sheet
 */
function archiveData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  let archiveSheetName = getSetting('ArchiveSheetName');

  if (!archiveSheetName) {
    const response = SpreadsheetApp.getUi().prompt(
      'Archive Sheet Name',
      'Enter the name of the archive sheet:',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      SpreadsheetApp.getUi().alert('Operation cancelled.');
      return;
    }
    archiveSheetName = response.getResponseText().trim();
    if (!archiveSheetName) {
      SpreadsheetApp.getUi().alert('No name entered. Operation cancelled.');
      return;
    }
  }

  let archiveSheet = ss.getSheetByName(archiveSheetName);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveSheetName);
  }

  const activeData = activeSheet.getDataRange().getValues();
  if (activeData.length < 2) {
    SpreadsheetApp.getUi().alert('No data to archive.');
    return;
  }

  const lastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(lastRow + 1, 1, activeData.length, activeData[0].length).setValues(activeData);

  SpreadsheetApp.getUi().alert(`Archived ${activeData.length} rows to "${archiveSheetName}".`);
}

/**
 * 📌 Protect header rows from editing
 */
function protectHeaderRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  let headerRows = parseInt(getSetting('ProtectRangeRows') || '0');

  if (!headerRows || headerRows < 1) {
    const response = SpreadsheetApp.getUi().prompt(
      'Protect Header Rows',
      'Enter the number of header rows to protect:',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
      SpreadsheetApp.getUi().alert('Operation cancelled.');
      return;
    }
    headerRows = parseInt(response.getResponseText()) || 1;
  }

  const protection = sheet.getRange(1, 1, headerRows, sheet.getLastColumn()).protect();
  protection.setDescription('Protected header rows');

  SpreadsheetApp.getUi().alert(`Protected top ${headerRows} row(s).`);
}

