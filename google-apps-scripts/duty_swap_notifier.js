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
  "full name": "email",
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
