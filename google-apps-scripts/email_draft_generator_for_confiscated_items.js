/**
 * Generates Gmail drafts to remind students to pick up their confiscated items based on data that was
 * previously inputted into a confiscated items spreadsheet.
 *
 * @version 1.0
 */
function generateGmailDrafts() {
  // -=-=-=- USER-SPECIFIED VARIABLES -=-=-=-
  // The URL of the Google spreadsheet containing the list of confiscated items and all associated student data
  var sheetURL = "https://docs.google.com/spreadsheets/d/1PggZQl15XqiHlCKz-RRMHVs_VRIZTUntGhMnfT1dEeY/edit?usp=sharing"
  
  // Spreadsheet column numbers
  var roomNumCol = letterToColumn('E');            // The index of the column containing the students' room number
  var confiscatedItemCol = letterToColumn('K');    // The index of the column containing the confiscated item
  var emailAddressCol = letterToColumn('H');       // The index of the column containing the student email addresses that the reminder will be sent to

  // Other variables
  var startRow = 2;    // The spreadsheet row that the list of confiscated items starts on

  // -=-=-=- MAIN PROGRAM -=-=-=-
  var sheet = SpreadsheetApp.openByUrl(sheetURL).getActiveSheet();    // The sheet containing the confiscated items list

  // Extract desired data range from confiscated items spreadsheet
  var numRows = sheet.getMaxRows() - startRow;
  var numColumns = sheet.getMaxColumns();
  var dataRange = sheet.getRange(startRow, 1, numRows, numColumns);
  var data = dataRange.getValues();

  // Iterate through the data, row by row
  var rowNum = 1;
  while (rowNum <= numRows) {
    var row = data[rowNum-1];

    var pickedUp = row[0];    // Whether or not the item has been picked up by the student or not
    if (pickedUp === "No") {
      var roomNum = row[roomNumCol-1];
      var subject = roomNum + " Confiscated Items Pickup";
      var confiscatedItem = row[confiscatedItemCol-1];
      var emailAddress = row[emailAddressCol-1];

      /* If the confiscation involves multiple students, add all their email addresses to the recipients
       * of the email reminder (all the email addresses listed within the merged row)
       */
      var cell = dataRange.getCell(rowNum, 2);
      if (cell.isPartOfMerge()) {
        var mergedRange = cell.getMergedRanges()[0];
        var numMergedRows = mergedRange.getNumRows();
        for (var i = 1; i < numMergedRows; i++) {
          rowNum++;
          var row = data[rowNum-1];
          var emailAddress = row[emailAddressCol-1];
          var recipientsTO = recipientsTO + "," + emailAddress;
        }
      }
      else {    // Only one student is involved with the confiscation
        var recipientsTO = emailAddress;
      }

      var htmlMessage = "Hi," +
        "<p>This is a reminder that you had an item confiscated by RLUH staff during a room inspection that occurred " +
          "in the past two months. The item that was confiscated was:" +
            "<p>" + confiscatedItem +
              "<p>If you haven't already done so, please pick up this item at the Holly Pointe Commons front desk in E-pod, 1st floor " +
                "between the hours of 8AM - 8PM Monday through Friday. <b>Please note that when you pick up your confiscated item, " +
                  "it is expected that you intend to leave campus and bring the item home with you immediately afterward.</b>" +

                    "<p>After you pick up the item, please respond to this email to stop receiving these reminders." +
                      "<p>Thank you." +
                        "<p>Sincerely,\n" +
                          "<br>Quinn McHugh\n" +
                            "<br>Assistant Resident Director, Holly Pointe Floors 1 & 2";

      GmailApp.createDraft(
        recipientsTO,
        subject,
        htmlMessage,
        {htmlBody: htmlMessage}
      );
    }
    rowNum++;
  }
}

/*
 * Converts a spreadsheet column number to its corresponding column letter.
 *
 * @param {int} colNum The column number to be converted
 * @return The corresponding column letter.
 */
function columnToLetter(colNum) {
  var temp = '';
  var colLetter = '';
  while (colNum > 0) {
    temp = (colNum - 1) % 26;
    colLetter = String.fromCharCode(temp + 65) + colLetter;
    colNum = (colNum - temp - 1) / 26;
  }
  return colLetter;
}

/*
 * Converts a spreadsheet column letter to its corresponding column number.
 *
 * @param {char} colLetter The column letter to be converted.
 * @return colNum The corresponding column number.
 */
function letterToColumn(colLetter) {
  var colNum = 0;
  var length = colLetter.length;
  for (var i = 0; i < length; i++) {
    colNum += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return colNum;
}
