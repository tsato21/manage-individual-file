/**
 * Triggers the authorization flow for Google services used in the script.
 * This function is useful for manually initiating the authorization process.
 */
function showAuthorizationDialog() {
  SpreadsheetApp;
  DriveApp;
  GmailApp;
}
/**
 * Creates a custom menu in the Google Sheets UI when the spreadsheet is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu with items to run specific functions.
  ui.createMenu('Custom Menu')
    .addItem('Copy and Name Files', 'copyAndNameFile')
    .addSeparator()
    .addItem('Output File Information', 'outputFileInfo')
    .addSeparator()
    .addItem('Share Files Without Notification', 'shareFilesWithoutNotification')
    .addSeparator()
    .addItem('Reset All Sharing Status', 'resetAllSharingStaus')
    .addSeparator()
    .addItem('Create Email Drafts', 'createDrafts')
    .addToUi();
}
/**
 * Copies a sample file multiple times based on a list of names from a Google Sheet.
 * Names are fetched from a predefined sheet and each copied file is stored in a specified folder.
 */
function copyAndNameFile() {
  let folderUrl = Browser.inputBox('Input the URL of the folder to store newly created files', Browser.Buttons.OK_CANCEL);
  if (!inputIsValid_(folderUrl)) return;
  let folderId = extractIdFromUrl_(folderUrl);
  let storedFolder = DriveApp.getFolderById(folderId);
  if (!existsValidation_(storedFolder, 'Folder')) return;

  if (!inputIsValid_(CREATE_SHEET_NAME)) return;
  let createSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CREATE_SHEET_NAME);
  if (!existsValidation_(createSheet, 'Create Sheet')) return;
  let lastRow = createSheet.getLastRow() - 1;
  if (lastRow === 0) {
    Browser.msgBox(`Application File Names are not listed in ${CREATE_SHEET_NAME}.`);
    return;
  }
  let applicationSheetNames = createSheet.getRange(2, 1, lastRow, 1).getValues();

  let sampleSheetUrl = Browser.inputBox('Input the URL of the sample sheet', Browser.Buttons.OK_CANCEL);
  if (!inputIsValid_(sampleSheetUrl)) return;
  let sampleSheetId = extractIdFromUrl_(sampleSheetUrl);
  let sampleSheet = DriveApp.getFileById(sampleSheetId);
  if (!existsValidation_(sampleSheet, 'Sample Sheet')) return;

  for (let eachSheetName of applicationSheetNames) {
    sampleSheet.makeCopy(eachSheetName[0], storedFolder);
  }

  Browser.msgBox('All of the application sheet files were successfully created and stored in the designated folder.');
}
/**
 * Validates the user input from a prompt.
 * @param {string} input - The user input to validate.
 * @return {boolean} True if input is valid, false otherwise.
 */
function inputIsValid_(input) {
  if (input === 'cancel' || input === '') {
    Browser.msgBox('Input was cancelled or empty.');
    return false;
  }
  return true;
}
/**
 * Checks if a specific object (like a file or folder) exists.
 * @param {Object} object - The object to validate.
 * @param {string} objectType - The type of the object (e.g., 'Folder').
 * @return {boolean} True if the object exists, false otherwise.
 */
function existsValidation_(object, objectType) {
  if (!object) {
    Browser.msgBox('The designated ' + objectType.toLowerCase() + ' does not exist.');
    return false;
  }
  return true;
}

/**
 * Extracts the ID from a given URL.
 * @param {string} url - The URL to extract the ID from.
 * @return {string|null} The extracted ID, or null if no ID is found.
 */
function extractIdFromUrl_(url) {
  let match = url.match(/[-\w]{25,}/);
  if (!match) {
    Browser.msgBox('Invalid URL. Please check the URL and try again.');
    return null;
  }
  return match[0];
}
/**
 * Gathers and displays information about files stored in a specified folder.
 * Information such as file name, URL, ID, editors' emails, and custom identifiers are fetched and displayed.
 */
function outputFileInfo() {
  let outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHARE_SHEET_NAME);
  let preLastRow = outputSheet.getRange("A3:A").getValues().filter(String).length;
  if (preLastRow > 0) {
    outputSheet.getRange(3, 1, preLastRow, 4).clearContent();
  }

  let folderUrl = Browser.inputBox('Input the URL of the folder to store target files', Browser.Buttons.OK_CANCEL);
  if (!inputIsValid_(folderUrl)) return;
  let folderId = extractIdFromUrl_(folderUrl);
  let storedFolder = DriveApp.getFolderById(folderId);
  if (!existsValidation_(storedFolder, 'Folder')) return;

  let targetFiles = storedFolder.getFiles();

  let fileData = [];
  while (targetFiles.hasNext()) {
    let file = targetFiles.next();
    let fileName = file.getName();
    let fileUrl = file.getUrl();
    let fileId = file.getId();
    let studentId = fileName.match(/【(.*?)_/);
    studentId = studentId ? studentId[1] : "Unknown";

    let editors = file.getEditors();
    let editorEmails = editors.map(editor => editor.getEmail()).join(", ");

    fileData.push([fileName, fileUrl, fileId, editorEmails, studentId]);
  }

  if (fileData.length > 0) {
    let fileNum = fileData.length;
    let colNum = fileData[0].length;
    let range = outputSheet.getRange(3, 1, fileNum, colNum);
    range.setValues(fileData);
  }

  Browser.msgBox(`Information of all of the target application sheets were successfully displayed in ${SHARE_SHEET_NAME}.`);
}
/**
 * Shares files without sending notification emails.
 * It sets sharing permissions for students and instructors based on email addresses provided in a Google Sheet.
 */
function shareFilesWithoutNotification() {
  let targets = getFileAndRecipientData_();

  console.log(`targets:\n${JSON.stringify(targets)}`);

  targets.forEach((target, index) => {
    console.log(target);
    try {
      Drive.Permissions.insert(
        {
          value: target['Student Email'],
          type: 'user',
          role: 'writer'
        },
        target['File ID'],
        { sendNotificationEmails: false }
      );

      Drive.Permissions.insert(
        {
          value: target['Instructor Email'],
          type: 'user',
          role: 'writer'
        },
        target['File ID'],
        { sendNotificationEmails: false }
      );

      // Get file by ID and retrieve current editors
      let file = DriveApp.getFileById(target['File ID']);
      let editors = file.getEditors();
      let editorEmails = editors.map(editor => editor.getEmail()).join(", ");

      // Update the spreadsheet with the editor emails
      refSheet.getRange(3 + index, 4).setValue(editorEmails);

    } catch (e) {
      console.error("Error processing file, " + target['Sheet Name'] + ": " + e.message);
      Browser.msgBox("Error processing file, " + target['Sheet Name'] + "Contact the owner of the script.");
    }
  });
}
/**
 * Resets all sharing permissions on files within a specified folder.
 * It removes all editors from each file in the folder.
 */
function resetAllSharingStaus() {
  let folderUrl = Browser.inputBox('Input the URL of the folder to reset sharing status', Browser.Buttons.OK_CANCEL);
  if (!inputIsValid_(folderUrl)) return;
  let folderId = extractIdFromUrl_(folderUrl);
  let resetFolder = DriveApp.getFolderById(folderId);
  if (!existsValidation_(resetFolder, 'Reset Folder')) return;

  let targetFiles = DriveApp.getFolderById(targetFolderId).getFiles();
  while (targetFiles.hasNext()) {
    let targetFile = targetFiles.next();
    let targetFileName = targetFile.getName();
    let editors = targetFile.getEditors();
    editors.forEach((editor) => {
      let editorEmail = editor.getEmail();
      targetFile.removeEditor(editorEmail);
    })
  }

  Browser.msgBox(`Reset sharing all of the target files were succeeded.`);
}
/**
 * Retrieves data from a Google Sheet for sharing and recipient information.
 * @return {Object[]} An array of objects, each representing a row from the sheet with key-value pairs.
 */
function getFileAndRecipientData_(){
  let refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHARE_SHEET_NAME);
  let dataRange = refSheet.getRange("A2:A");
  let lastRow = dataRange.getValues().filter(String).length;
  console.log(`last row is ${lastRow}`);

  if (lastRow === 0) {
    Browser.msgBox(`Any Spreadsheet Files are not listed from column A to D in ${SHARE_SHEET_NAME}.`);
    return;
  }

  let data = refSheet.getRange(2, 1, lastRow, 9).getValues();
  let header = data.shift(); // Assuming the first row of your data range contains the headers
  // Mapping each row of the data into an object based on the headers
  // This section transforms the data from a 2D array into an array of objects
  // Each object represents a row in the sheet, with key-value pairs corresponding to header and cell data
  let targets = data.map(row => {
    // The 'reduce' method is used to accumulate values into a single object
    // 'obj' is our accumulator object, 'key' is the current header, and 'index' is the current index
    return header.reduce((obj, key, index) => {
      // For each header (key), assign the corresponding value from the current row
      // row[index] refers to the cell in the current row under the current header
      obj[key] = row[index];

      // console.log(obj);
      // Return the accumulator object for the next iteration
      // On each iteration, this builds up the object with more key-value pairs
      return obj;
    }, {}); // The initial value of the accumulator 'obj' is an empty object
  });
  return targets;
}
/**
 * Creates draft emails in Gmail for each record in a dataset.
 * Drafts are personalized for students and instructors with specific content and links.
 */
function createDrafts() {
  let fileAndRecipientData = getFileAndRecipientData_(); // Ensure this function is defined

  for (let eachRecord of fileAndRecipientData) {
    let to = eachRecord["Student Email"];
    let cc = eachRecord["Instructor Email"];

    let subject = "Please Access the Application Sheet";
    let studentName = eachRecord["Student Name(ENG)"];
    let instructorName = eachRecord["Instructor Name(ENG)"];
    let fileLink = eachRecord["File Link"];
    let body = 
      `Dear ${studentName},<br>(CC: Prof. ${instructorName})<br><br>Please access your Application Sheet below.<br>${fileLink}<br><br>ONLY if you have any questions, please reply back to us.</p><br><br>Sincerely,<br><br>Division Name`;

    let options = {
      "htmlBody": body,
      "cc": cc
    };

    GmailApp.createDraft(to, subject, "", options); // Use 'to' for the recipient and an empty string for plain text body
  }

  Browser.msgBox(`Drafts were successfully created in Gmail.`);
}

