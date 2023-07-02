function sendThankYouEmail() {
  const formResponsesSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = formResponsesSheet.getDataRange();
  const data = dataRange.getValues();

  const lastRowIndex = dataRange.getLastRow();
  const lastResponse = data[lastRowIndex - 1];

  const name = lastResponse[1];
  const email = lastResponse[4];

  const subject = "Thank you for submitting the form";
  const body = `Dear ${name},\n\nThank you for submitting the form. We appreciate your participation.\n\nBest regards,\nSuman Bhandari`;

  MailApp.sendEmail(email, subject, body);
  Logger.log(email, subject, body);
}
