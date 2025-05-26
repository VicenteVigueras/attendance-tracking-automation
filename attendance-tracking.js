function sendAbsenceAlert() {
  var spreadsheet = SpreadsheetApp.openById("");
  var sheet = spreadsheet.getSheetByName("");
  var logSheet = spreadsheet.getSheetByName("") || spreadsheet.insertSheet("Notifications Log");

  var data = sheet.getDataRange().getValues();
  var ccEmails = "";

  for (var i = 1; i < data.length; i++) {
    var email = data[i][5]; // Column F (index 5) contains emails
    var absences = parseInt(data[i][3]) || 0; // Column D (index 3) contains absence counts
    var lates = parseInt(data[i][2]) || 0; // Column C (index 2) contains late counts
    var notifiedStatus = (data[i][0] || "").split(",").map(String); // Column A tracks "Notified" status
    var studentName = data[i][4]; // Column E (index 4) contains student name

    if (!email) {
      Logger.log(`Skipping row ${i + 1}: Missing email.`);
      continue;
    }

    var shouldNotifyAbsences = absences >= 3 && !notifiedStatus.includes(absences + "A");
    var shouldNotifyLates = lates >= 3 && !notifiedStatus.includes(lates + "L");

    if (shouldNotifyAbsences || shouldNotifyLates) {
      var subject = "Attendance Status Notification";
var emailBody = `
  <p>Dear <strong>${studentName}</strong>,</p>
  <p>This is a reminder about the attendance standards for our program and an update on your current attendance record.</p>
  <p>As outlined in our attendance policy, consistent attendance and punctuality are critical to your success in this Software Engineering training program. Our attendance practices reflect the expectations of employers in the industry and are designed to support your professional development.</p>
  
  <hr>
  <p>Sincerely,<br><strong>Program Support</strong></p>
`;

     GmailApp.sendEmail(email, subject, '', {
  htmlBody: emailBody,
  cc: ccEmails,
  from: ""
});


      var updatedNotified = [...notifiedStatus];
      if (shouldNotifyAbsences && !updatedNotified.includes(absences + "A")) updatedNotified.push(absences + "A");
      if (shouldNotifyLates && !updatedNotified.includes(lates + "L")) updatedNotified.push(lates + "L");

      sheet.getRange(i + 1, 1).setValue(updatedNotified.join(","));

      logSheet.appendRow([
        new Date().toLocaleString(),
        studentName,
        email,
        absences,
        lates,
        updatedNotified.join(",")
      ]);

      Logger.log(`Email sent to ${studentName} (${email}) for ${shouldNotifyAbsences ? absences + " absences" : ""} ${shouldNotifyLates ? lates + " lates" : ""}.`);
    }
  }
}
