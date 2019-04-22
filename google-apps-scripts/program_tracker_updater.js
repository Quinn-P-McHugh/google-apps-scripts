/* The following document uses JSDoc 3 standard for JavaScript documentation: http://usejsdoc.org/ */

/**
 * STATUS: Unfinished because I realized that this would save me virtually no time during
 * the semester, as I already sunk a couple hours into writing this program. In addition,
 * automating this portion of my responsibilities as an Assistant Resident Director (ARD) is inadvisable
 * because using this program would enable me to complete a portion of my responsibilities
 * without actually having to know what kinds of programs my RAs are hosting, which is something an ARD
 * absolutely SHOULD know off the top of their head.
 *
 * That said, it was nice to brush up on some JavaScript and continue developing my knowledge of Google's
 * scripting platform. Overall, the project was a good learning experience, but certainly not
 * a productive one.
 * _____________________
 *
 * Automatically updates resident assistant (RA) engagement model trackers by reading from program
 * proposal and program evaluation emails saved as text files by extracting necessary data from each
 * file and inputting it into each RAs corresponding engagement model tracker spreadsheet on Google
 * Sheets.
 *
 * @version 1.0
 */
function updateAllTrackers() {
  var files = getFiles("Program Proposal Text Files");
  for (var i = 0; i < files.length; i++) {

    /* Extract all text in the file and split it line by line */
    var text = DocumentApp.openById(files[i].getId()).getBody().getText();
    var lines = text.split(/\r?\n/);

    var fullName = extractFieldValue(lines, "What is your full name (First Last)?").split(" ");
    var firstName = fullName[0];
    Logger.log(firstName);
    if (getTrackerSheet(firstName) != null) {
      var trackerSheet = getTrackerSheet(firstName);
      Logger.log(trackerSheet.getName());
    }
  }
};

/*
 * Extracts data from program proposal and program evaluation text files and inputs
 * it into the user-specified program tracker.
 *
 * @param {Array} lines        An array of text, split line-by-line.
 * @param {Sheet} trackerSheet The program tracker where data will be inputted.
 */
function inputTrackerData(lines, trackerSheet) {
  var programName = extractFieldValue(lines, "Title of event:");
  var proposalDate = extractFieldValue(lines, "Today's Date:");
  var programDate = extractFieldValue(lines, "Date of event:");

  var eventType = extractFieldValue(lines, "Event type:");
  if (eventType.localCompare("Community Builder (CB)")) {
    /* Input event data under community builder table in program tracker */
  }
  else if (eventType.localCompare("Campus and Community Connection (C3)")) {
    /* Input event data under C3 table in program tracker */
  }
}

/**
 * Extracts a text field value from the list of text using a user-specified text field.
 *
 * @param {Array}   lines An array of text, split line-by-line.
 * @return {String} field The text field that corresponds to the desired text field value.
 */
function extractFieldValue(lines, field) {
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].indexOf(field) > -1) {
      return lines[i+1].trim();
    }
  }
};

/**
 * Returns the program tracker sheet associated with an RA's last name.
 *
 * @param {String} lastName The last name of the RA.
 * @return {Sheet} sheet    The sheet whose name contains the specified last name.
 */
function getTrackerSheet(lastName) {
  var filesIter = DriveApp.getFilesByName("HPC 1&2: Fall 2018 - Program Tracker & Ledger");
  var programTracker = filesIter.next();
  programTracker = SpreadsheetApp.open(programTracker);
  var sheets = programTracker.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName().indexOf(lastName) > -1) {
      return sheet;
    }
  }
};

/**
 * Returns a collection of all files in a specified Google Drive folder.
 *
 * @param {String} folderName The name of the folder that contains the desired files.
 * @return {Array} files      A collection of files that are in the specified folder.
 */
function getFiles(folderName) {
  var folderIter = DriveApp.getFoldersByName(folderName);
  var folder = folderIter.next();
  var filesIter = folder.getFiles();
  var files = [];
  while (filesIter.hasNext()) {
    var file = filesIter.next();
    files.push(file);
    Logger.log(file);
  }
  return files;
};
