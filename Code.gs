/**
 * === Google Drive File Organizer and Sharer ===
 * Author: [Your Name]
 * License: MIT
 *
 * Moves or copies files from SourceFolder to DestinationFolder with:
 * - Advanced renaming options
 * - MIME type/category filtering
 * - Filename contains filtering
 * - Batch limit and delay
 * - Date-stamped subfolders
 * - Sharing permissions from Settings
 * - Dry run / preview mode
 * - Logging to a "Logs" sheet
 *
 * Reads all configuration from a single "Settings" tab in Google Sheets.
 */

function organizeFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName('Settings');
  const logSheet = getOrCreateLogSheet(ss);

  if (!settingsSheet) {
    ui.alert('Settings sheet not found!');
    return;
  }

  // === Parse settings ===
  const settings = getGeneralSettings(settingsSheet);
  const shareRules = getSharingRules(settingsSheet);

  // === Interactive prompts for missing settings ===
  settings.SourceFolderID = settings.SourceFolderID || promptUser('Source Folder ID');
  settings.DestinationFolderID = settings.DestinationFolderID || promptUser('Destination Folder ID');
  settings.MoveMode = settings.MoveMode || promptUser('Move Mode (move/copy)');
  settings.RenameMode = settings.RenameMode || 'none';
  settings.MaxFiles = parseInt(settings.MaxFiles || '10');
  settings.DryRunMode = (settings.DryRunMode || 'FALSE').toUpperCase() === 'TRUE';
  settings.DelaySeconds = parseInt(settings.DelaySeconds || '0');
  settings.UseDateSubfolder = (settings.UseDateSubfolder || 'FALSE').toUpperCase() === 'TRUE';

  // === Validate folders ===
  const sourceFolder = getDriveFolderById(settings.SourceFolderID);
  const destinationFolder = getOrCreateDestinationFolder(settings);

  if (!sourceFolder || !destinationFolder) {
    ui.alert('Error accessing Source or Destination folder!');
    return;
  }

  // === Get files to process ===
  let files = sourceFolder.getFiles();
  let processed = 0;
  let errors = 0;

  while (files.hasNext() && processed < settings.MaxFiles) {
    try {
      let file = files.next();

      // === Filter: AllowedMimeTypes / AllowedCategory ===
      if (!isAllowedMimeType(file, settings)) continue;

      // === Filter: Filename Contains Text ===
      if (settings.MoveIfContainsText && !file.getName().toLowerCase().includes(settings.MoveIfContainsText.toLowerCase())) {
        continue;
      }

      // === Determine new name ===
      const newName = getRenamedFileName(file.getName(), settings);

      // === Determine subfolder ===
      let targetFolder = destinationFolder;
      if (settings.UseDateSubfolder) {
        targetFolder = getOrCreateDateSubfolder(destinationFolder);
      }

      // === Dry Run ===
      if (settings.DryRunMode) {
        logSheet.appendRow([
          new Date(), file.getName(), newName, 'DRY RUN', settings.SourceFolderID, settings.DestinationFolderID, '', ''
        ]);
        processed++;
        continue;
      }

      // === Move or Copy ===
      let newFile;
      if (settings.MoveMode.toLowerCase() === 'copy') {
        newFile = file.makeCopy(newName, targetFolder);
      } else {
        file.setName(newName);
        targetFolder.addFile(file);
        sourceFolder.removeFile(file);
        newFile = file;
      }

      // === Apply Sharing Rules ===
      applySharingRules(newFile, shareRules);

      // === Log success ===
      logSheet.appendRow([
        new Date(), file.getName(), newName, 'SUCCESS', settings.SourceFolderID, settings.DestinationFolderID, '', ''
      ]);

      processed++;

      // === Delay if needed ===
      if (settings.DelaySeconds > 0) Utilities.sleep(settings.DelaySeconds * 1000);

    } catch (e) {
      errors++;
      Logger.log(e);
      logSheet.appendRow([
        new Date(), '', '', 'ERROR', settings.SourceFolderID, settings.DestinationFolderID, e.toString(), ''
      ]);
    }
  }

  ui.alert(`Processing complete.\nProcessed: ${processed}\nErrors: ${errors}`);
}

/**
 * === Helpers ===
 */

function getGeneralSettings(sheet) {
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || !data[i][1]) continue;
    if (data[i][0] === 'ShareType') break; // stop at Sharing Rules
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function getSharingRules(sheet) {
  const data = sheet.getDataRange().getValues();
  const rules = [];
  let inRules = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'ShareType') {
      inRules = true;
      continue;
    }
    if (inRules && data[i][0]) {
      rules.push({ type: data[i][0], emails: parseEmailList(data[i][1]) });
    }
  }
  return rules;
}

function parseEmailList(emailString) {
  if (!emailString) return [];
  return emailString.split(',').map(e => e.trim()).filter(e => e);
}

function promptUser(label) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(`Enter value for ${label}:`);
  return result.getSelectedButton() === ui.Button.OK ? result.getResponseText().trim() : '';
}

function getDriveFolderById(folderId) {
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('Invalid folder ID: ' + folderId);
    return null;
  }
}

function getOrCreateDestinationFolder(settings) {
  const destination = getDriveFolderById(settings.DestinationFolderID);
  if (!destination) return null;
  return destination;
}

function getOrCreateDateSubfolder(destinationFolder) {
  const dateString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let subfolders = destinationFolder.getFoldersByName(dateString);
  if (subfolders.hasNext()) return subfolders.next();
  return destinationFolder.createFolder(dateString);
}

function getRenamedFileName(originalName, settings) {
  let name = originalName;
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  switch ((settings.RenameMode || '').toLowerCase()) {
    case 'prefix':
      name = settings.RenameText + name;
      break;
    case 'suffix':
      name = name + settings.RenameText;
      break;
    case 'timestamp':
      name = name + '_' + timestamp;
      break;
    case 'replace':
      if (settings.ReplaceFrom && settings.ReplaceTo) {
        name = name.replace(settings.ReplaceFrom, settings.ReplaceTo);
      }
      break;
  }
  return name;
}

function isAllowedMimeType(file, settings) {
  const mime = file.getMimeType();
  if (settings.AllowedMimeTypes) {
    const allowedList = settings.AllowedMimeTypes.split(',').map(e => e.trim());
    if (!allowedList.includes(mime)) return false;
  }
  if (settings.AllowedCategory) {
    if (!isMimeInCategory(mime, settings.AllowedCategory)) return false;
  }
  return true;
}

function isMimeInCategory(mime, category) {
  const categories = {
    'Documents': [
      'application/pdf',
      'application/vnd.google-apps.document',
      'application/msword'
    ],
    'Images': [
      'image/png',
      'image/jpeg',
      'image/gif'
    ],
    'Videos': [
      'video/mp4',
      'video/quicktime'
    ]
  };
  return categories[category] && categories[category].includes(mime);
}

function applySharingRules(file, shareRules) {
  for (let rule of shareRules) {
    for (let email of rule.emails) {
      try {
        if (rule.type.toLowerCase() === 'reader') {
          file.addViewer(email);
        } else if (rule.type.toLowerCase() === 'writer') {
          file.addEditor(email);
        } else if (rule.type.toLowerCase() === 'commenter') {
          file.addCommenter(email);
        }
      } catch (e) {
        Logger.log(`Error sharing with ${email}: ${e}`);
      }
    }
  }
}

function getOrCreateLogSheet(ss) {
  let logSheet = ss.getSheetByName('Logs');
  if (!logSheet) {
    logSheet = ss.insertSheet('Logs');
    logSheet.appendRow([
      'Timestamp', 'Original File Name', 'New File Name', 'Status',
      'SourceFolderID', 'DestinationFolderID', 'Error', 'Notes'
    ]);
  }
  return logSheet;
}
