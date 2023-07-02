function PayslipSender() {
    var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
  
    if (spreadsheet) {
      var lastRow = spreadsheet.getLastRow();
      var dataRange = spreadsheet.getRange("A2:C" + lastRow).getValues();
  
      for (var i = 0; i < dataRange.length; i++) {
        var employeeData = dataRange[i];
        var employeeName = employeeData[0];
        var email = employeeData[1];
        console.log(email)
        var salary = employeeData[2];
  
        if (email && email !== "") {
          var payslipMessageContent = PayslipMessage(employeeName, salary);
         MailApp.sendEmail(email, 'PaySlip', payslipMessageContent);
        } else {
          console.log("Skipping payslip email for employee " + employeeName + " due to missing or invalid email address.");
        }
      }
    } else {
      console.log("Unable to find the 'Sheet1' sheet or active spreadsheet.");
    }
  }
  
  function PayslipMessage(employeeName, salary) {
    var message = "Hi " + employeeName + "\n";
    message += "We are pleased to inform you that your salary for Jestha has been successfully credited to your account" + "\n";
    message += "Payable: " + salary + "\n";
    message += "Best regards," + "\n";
    message += "Suman Bhandari" + "\n";
    return message;
  }
  
  PayslipSender()
  