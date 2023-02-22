function main(TimeOfDay) {

/* 
Author: Sebastian Schneider, 20/01/23.

The script's purpose is to send an email reminder to specific individuals to fill out their "Mitarbeiterstunden". 
It uses a dictionary, myDictionary, to determine which employees should be considered and which cells in the spreadsheet correspond to each employee's sign-in and sign-out times. It will check the relevant cells in the spreadsheet to see if they are empty and if they are send an email reminder to the respective person. 
The script has two different modes of operation, "morning" and "evening", which correspond to the specific time of day when the script is run and, accordingly, check different cells in the spreadsheet.

It is important to note that the input for the script must be correct in order for it to function properly. Specifically, the script assumes that the data for the "Mitarbeiterstunden" form starts couting in row 3 on 01/01/23 of the active sheet and that myDictionary correctly maps employee names to the corresponding cells in the spreadsheet where their sign-in times are recorded. Furthermore, it assumes that a person's sign out column is to the right of the respective person's sign in column. 

If the data in the spreadsheet is not correctly aligned with the input in the script, it may not be able to find the correct cells to check or may send email reminders to the wrong individuals. Therefore, it is important to verify that the start day in row 3 and the employee information in myDictionary are accurate before running the script.


In short: If you want to make changes to the Mitarbeiter, adapt myDictionary.
        : If you want to start counting from another date and in another row, adapt startDate and relevant_row, respectively.
*/


  var sheet = SpreadsheetApp.getActiveSheet();
    var thisDocumentUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();


  // Dictonnary with { Key: ["Cell", "Email"]}
  var myDictionary = [
    { "Sebastian":  ["B", "sebastians-mail"] },
	{ "Christian":  ["D", "christians-mail"] },
	{ "Max":        ["F", "maxs-mail"] },
	{ "Mustermann": ["H", "mustermanns-mail"] }
  ];

var n = findElapsedDays();
// Logger.log("Number of days between jan 1st 2023 and today: " + n);

relevant_row = n + 3; // this is due to the fact that we start in row 3 with January 1st
// Logger.log("relevant_row" + relevant_row);

if (TimeOfDay == "morning") {
  Logger.log(TimeOfDay)
  for (var i = 0; i < myDictionary.length; i++) {
    // iterates through the dictionnary, checks if relevant cells are empty and sends an e-mail to the specific person if it is
    var key = Object.keys(myDictionary[i])[0];
    var relevant_col = myDictionary[i][key][0];
    var relevant_cell = relevant_col + relevant_row;
    var relevant_cell_content = sheet.getRange(relevant_cell).getValues()[0][0];

    // The following will be terribly inefficient, but was easy to implement, check the previous-evening cell.
    var relevant_row_evening_before      = relevant_row -1;
    var relevant_col_evening_before      = getNextLetter(relevant_col); // Note that the sign-out time is in the column on the right of the column specified in the dictionnary
    var relevant_cell_evening_before     = relevant_col_evening_before + relevant_row_evening_before;
    relevant_cell_content_evening_before = sheet.getRange(relevant_cell_evening_before).getValues()[0][0];

    // Logger.log("We consider cell " + relevant_cell + " and note that it's content is: " + relevant_cell_content);

    if (relevant_cell_content == "" || relevant_cell_content_evening_before == "") {
      var recipient = myDictionary[i][key][1];
      // Logger.log("Mail will be sent to" + recipient);
      var subject = "Liebe/r " + key + ". Bitte das heutige Mitarbeiterstundenformular ausfüllen";
      var body = "Liebe/r " + key + ". \nBitte fülle das Mitarbeiterstundenformular aus, im speziellen, Zelle " + relevant_cell + " (heute morgen) und " + relevant_cell_evening_before + " (gestern abend) in der Tabelle (s.u.). Das ist eine automatisch generierte E-Mail. Bei Anregungen oder Problemen, wende Dich bitte an mich. \n\nLiebe Grüße, \nSebastian\n" + thisDocumentUrl;
      MailApp.sendEmail(recipient, subject, body);
    }
  }
}

if (TimeOfDay == "evening") {
  Logger.log(TimeOfDay)
  for (var i = 0; i < myDictionary.length; i++) {
    // iterates through the dictionnary, checks if relevant cells are empty and sends an e-mail to the specific person if it is
    var key = Object.keys(myDictionary[i])[0];
    var relevant_col = myDictionary[i][key][0];
    relevant_col = getNextLetter(relevant_col); // Note that the sign-out time is in the column on the right of the column specified in the dictionnary
    var relevant_cell = relevant_col + relevant_row;
    var relevant_cell_content = sheet.getRange(relevant_cell).getValues()[0][0];
    // Logger.log("We consider cell " + relevant_cell + " and note that it's content is: " + relevant_cell_content);
    
    if (relevant_cell_content == "") {
      var recipient = myDictionary[i][key][1];
      // Logger.log("Mail will be sent to" + recipient);
      var subject = "Liebe/r, " + key + ". Bitte das Mitarbeiterstundenformular ausfüllen";
      var body = "Liebe/r, " + key + ". Bitte fülle das Mitarbeiterstundenformular aus, im Speziellen, Zelle " + relevant_cell + ". This is today's time of sign out cell.\nKind regards, \nSebastian";
      MailApp.sendEmail(recipient, subject, body);
    }
  }
}

function findElapsedDays() {
  // This function outputs the amount of days between today and our start date, January 1st 2023
  var startDate = new Date(2023, 0, 1); // 1st january 2023
  var today = new Date();
  var elapsed = (today.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24);
  return Math.floor(elapsed);
}

function getNextLetter(letter) {
  //  This function letter as input and returns the next letter in the alphabet. Note that it only works for letters A-Y.
  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var index = alphabet.indexOf(letter);
  if (index === -1) {
    return "Invalid input";
  } else if (index == alphabet.length  -1) {
    return alphabet[0];
  } else {
    return alphabet[index + 1];
  }
}}