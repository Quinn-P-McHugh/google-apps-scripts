/**
 * Sends custom duty swap confirmation emails to each RA staff member once the duty swap Google Form is submitted
 *
 * For this script to function correctly, an "On Form Submit" trigger must be set up through Google Forms. For more
 * information, please visit the following help page: https://developers.google.com/apps-script/guides/triggers/events#form-submit_4
 */

/**
 * Input RA/ARD email addresses here using the syntax shown. 
 *
 * Please note: RA/ARD names listed here MUST match the RA/ARD names 
 * in the dropdown field called "RA/ARD you are swapping with:" on the 
 * Google Form.
 */
var emails = {
  "Aleeyah Oliphant": "oliphanta4@students.rowan.edu",
  "Benjamin Taylor": "taylorb9@students.rowan.edu",
  "Colin Yost": "yostc8@students.rowan.edu",
  "Destiny Hall": "halld5@students.rowan.edu",
  "Erin Flynn": "flynne7@students.rowan.edu",
  "Errin Edwards": "edwardse4@students.rowan.edu",
  "Ethan Ellis": "ellise5@students.rowan.edu",
  "Frank Versace": "versacef6@students.rowan.edu",
  "Gatha Adhikari": "adhikarig1@students.rowan.edu",
  "Hesham Nassar": "nassarh4@students.rowan.edu",
  "James Vega": "vegaj7@students.rowan.edu",
  "James Matinog": "james matinogj4@students.rowan.edu",
  "Jasmine Dennis": "dennisj6@students.rowan.edu",
  "Jenna Greenlee": "greenleej6@students.rowan.edu",
  "Jonathan Maturano": "maturanoj7@students.rowan.edu",
  "Josiah Bell": "bellj5@students.rowan.edu",
  "Justin Roldan": "roldanj0@students.rowan.edu",
  "Katherine Villacis": "villac89@students.rowan.edu",
  "Kimberly Gould": "gouldk2@students.rowan.edu",
  "Laura Colandrea": "colandrel4@students.rowan.edu",
  "Marian Suganob": "suganobm6@students.rowan.edu",
  "Matthew Acquarola": "acquarolm4@students.rowan.edu",
  "Mia Sclafani": "sclafanim6@students.rowan.edu",
  "Mitchell McDaniels": "mcdanielm2@students.rowan.edu",
  "Nah'Ja Washington": "washingtn6@students.rowan.edu",
  "Quinn McHugh": "mchughq5@students.rowan.edu",
  "Ravi Dhruv": "dhruvr4@students.rowan.edu",
  "RJ Smith": "smithr37@students.rowan.edu",
  "Robin Purtell": "purtellr6@students.rowan.edu",
  "Serwah Danquah": "danquahs7@students.rowan.edu",
  "Stephanie Ibe": "ibes7@students.rowan.edu",
  "Stephanie Scott": "scotts6@students.rowan.edu",
  "Thomas Helbock": "helbockt4@students.rowan.edu",
  "Tsion Abay": "abayt7@students.rowan.edu",
  "Zach Schoellig": "schoelliz9@students.rowan.edu",
}

/* Specify columns in form responses spreadsheet here: */
var StaffMember1_Name_Column = "C";    // The column containing the name of staff member 1
var StaffMember2_Name_Column = "E";    // The column containing the name of staff member 2
var StaffMember1_NewDutyDate_Column = "I";
var StaffMember2_NewDutyDate_Column = "F";

var StaffMember = function(name, newDutyDate) {
  this.name = name;
  this.email = emails[name];
  if (newDutyDate == "") {
    this.newDutyDate = "N/A";
  }
  else {
    this.newDutyDate = newDutyDate;
  }
}

/* 
 * Converts a spreadsheet column letter to its corresponding column number. 
 
 * @param  {character} colLetter The column letter to be converted.
 * @return {integer}   colNum    The corresponding column number.
 */
function letterToColumn(colLetter) {
  var colNum = 0;
  var length = colLetter.length;
  for (var i = 0; i < length; i++) {
    colNum += (colLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return colNum;
}

function onFormSubmit(e) {
  var formValues = e.values;
  
  var StaffMember1_Name = formValues[letterToColumn(StaffMember1_Name_Column) - 1];
  var StaffMember2_Name = formValues[letterToColumn(StaffMember2_Name_Column) - 1];
  var StaffMember1_NewDutyDate = formValues[letterToColumn(StaffMember1_NewDutyDate_Column) - 1];
  var StaffMember2_NewDutyDate = formValues[letterToColumn(StaffMember2_NewDutyDate_Column) - 1];
  
  /* Get specific form values from form responses spreadsheet */
  var staffMember1 = new StaffMember(StaffMember1_Name, StaffMember1_NewDutyDate);
  var staffMember2 = new StaffMember(StaffMember2_Name, StaffMember2_NewDutyDate);
  
  staffMember1.message = "Hi," +
    "<p>Thank you for submitting an RA duty swap request. Your request will be reviewed by RLUH administrative staff " +
    "and you will be notified if your request is approved or not. Here's the swap you requested: " +
    "<p>" + staffMember1.name + " <i><u>was</u></i> on duty on: <b>" + staffMember2.newDutyDate + "</b>" +
    "<br>" + staffMember1.name + " <i><u>will now be<u></i> on duty on: <b>" + staffMember1.newDutyDate + "</b>" +
    "<p>" + staffMember2.name + " <i><u>was</u></i> on duty on: <b>" + staffMember1.newDutyDate + "</b>" +
    "<br>" + staffMember2.name + " <i><u>will now be<u></i> on duty on: <b>" + staffMember2.newDutyDate + "</b>" +
    "<p>" + staffMember2.name + " has also also been notified about this request." +
    "<p>If you have any questions or noticed a mistake in your swap request, please contact your RD. Thank you!";
   
  staffMember2.message = "Hi," +
    "<p>" + staffMember1.name + " recently requested to swap duty with you. Please review the duty " +
    "swap information below and verify that it is correct. If this request was a mistake, " +
    "please email " + staffMember1.name + " and your RD as soon as possible. Here's the swap that was requested: " +
    "<p>" + staffMember1.name + " <i><u>was</u></i> on duty on: <b>" + staffMember2.newDutyDate + "</b>" +
    "<br>" + staffMember1.name + " <i><u>will now be<u></i> on duty on: <b>" + staffMember1.newDutyDate + "</b>" +
    "<p>" + staffMember2.name + " <i><u>was</u></i> on duty on: <b>" + staffMember1.newDutyDate + "</b>" +
    "<br>" + staffMember2.name + " <i><u>will now be<u></i> on duty on: <b>" + staffMember2.newDutyDate + "</b>" +
    "<p>If you have any questions or noticed a mistake in this swap request, please contact your RD. Thank you!";
  
  var subject = "RA Duty Swap Request Confirmation";
  
  GmailApp.sendEmail(staffMember1.email, subject, "", {
    htmlBody: staffMember1.message,
    cc: ""
  });
  
  GmailApp.sendEmail(staffMember2.email, subject, "", {
    htmlBody: staffMember2.message,
    cc: ""
  });
}
