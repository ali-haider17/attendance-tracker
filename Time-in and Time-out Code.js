function onEdit(e) {
  if(e && e.range){
    try {
    const range = e.range;
    const column = range.getColumn();
    const row = range.getRowIndex();
    const usaTime = getFormattedUSTime();

    if (column == 4 || column == 6) {
      range.setNote('Last modified: ' + usaTime);
      SpreadsheetApp.getActiveSheet().getRange(row, 14).setValue(usaTime);

      const update = SpreadsheetApp.getActiveSheet().getRange(row, 14).getValue();
      const new_update = update.split(',')[0]; // Extract the part before the comma
      const current = SpreadsheetApp.getActiveSheet().getRange(row, 3).getValue();
      SpreadsheetApp.getActiveSheet().getRange(row, 18).setValue(new_update);
      SpreadsheetApp.getActiveSheet().getRange(row, 19).setValue(current);

      const diff = calculateDateDifference(
        SpreadsheetApp.getActiveSheet().getRange(row, 18).getValue(),
        SpreadsheetApp.getActiveSheet().getRange(row, 19).getValue()
      );

      // Set "Late" status when delay in days is greater than 2
      if (diff > 0 && (diff < 2 || diff > 2)) {
        SpreadsheetApp.getActiveSheet().getRange(row, 16).setValue("Late");
      } else if (diff < 0 && SpreadsheetApp.getActiveSheet().getRange(row, 4).getValue() !== "" &&
        SpreadsheetApp.getActiveSheet().getRange(row, 6).getValue() !== "") {
        SpreadsheetApp.getActiveSheet().getRange(row, 16).setValue("Early");
        sendEmail(row);
        protectAndSetDomainEdit(false);
      } else if (diff === 0 && SpreadsheetApp.getActiveSheet().getRange(row, 4).getValue() !== "" &&
        SpreadsheetApp.getActiveSheet().getRange(row, 6).getValue() !== "") {
        SpreadsheetApp.getActiveSheet().getRange(row, 16).setValue("Marked");
        sendEmail(row);
      }
    } else if (column === 10 && SpreadsheetApp.getActiveSheet().getRange(row, 10).isChecked() !== null) {
      SpreadsheetApp.getActiveSheet().getRange(row, 12).insertCheckboxes();
      SpreadsheetApp.getActiveSheet().getRange(row, 11).setValue(usaTime);
    } else if (column === 12 && SpreadsheetApp.getActiveSheet().getRange(row, 12).isChecked() !== null) {
      SpreadsheetApp.getActiveSheet().getRange(row, 13).setValue(usaTime);
    }
  } catch (error) {
    // Handle any errors or log them as needed
    console.error(error);
  }
  }

  else {
    console.error("Event object or range is undefined. This script is designed to run on cell edits.");
  }
  
}

function sendEmail(row) {
  try {
    const name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    const currentdate = SpreadsheetApp.getActiveSheet().getRange(row, 3).getValue();
    const useremail = SpreadsheetApp.getActiveSheet().getRange('AB12').getValue();
    const supervisoremail = SpreadsheetApp.getActiveSheet().getRange('AC12').getValue();

    if (SpreadsheetApp.getActiveSheet().getRange(row, 16).getValue() === "Marked") {
      // Send email for "Marked" status
      sendStatusEmail(useremail, supervisoremail, "Attendance Marked On Time!");
    } else if (SpreadsheetApp.getActiveSheet().getRange(row, 16).getValue() === "Early") {
      // Send email for "Early" status
      sendStatusEmail(useremail, "ali.dymaxtech@gmail.com", "Fix Your Attendance!");
    } else if (SpreadsheetApp.getActiveSheet().getRange(row, 16).getValue() === "Late") {
      // Send email for "Late" status
      sendStatusEmail(useremail, "ali.dymaxtech@gmail.com", "Late Attendance!");
    }
  } catch (error) {
    console.error(error);
  }
}

function sendStatusEmail(to, cc, subject) {
  try {
    // Construct and send the email
    MailApp.sendEmail({
      to: to,
      cc: cc,
      subject: subject,
      htmlBody: getEmailContent(subject),
    });
  } catch (error) {
    console.error(error);
  }
}

function getEmailContent(status) {
  const greetings = "<h3>Greetings!</h3>";
  if (status === "Attendance Marked On Time!") {
    return greetings + "This is to inform you that your attendance has been marked!";
  } else if (status === "Fix Your Attendance!") {
    return greetings + "<p style='color:red'>This is to inform you that you are marking your attendance earlier than expected!</p>" +
      "If you don't fix it in time, it will be marked as ABSENT!";
  } else if (status === "Late Attendance!") {
    return greetings + "<p style='color:red'>This is to inform you that you are marking your attendance late on </p>" +
      "<p style='color:red'><strong>Note:</strong> If you don't update your attendance within a period of 2 Days, it will be marked as ABSENT.</p>";
  }
}

function getFormattedUSTime() {
  try {
    const date = new Date();
    return date.toLocaleString("en-US", { timeZone: "Asia/Karachi" });
  } catch (error) {
    console.error(error);
    return "";
  }
}

function calculateDateDifference(date1, date2) {
  try {
    return (date1 - date2) / (1000 * 60 * 60 * 24);
  } catch (error) {
    console.error(error);
    return 0;
  }
}

function protectAndSetDomainEdit(domainEdit) {
  try {
    const protection = e.range.protect();
    protection.setDescription('Sample protected range');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(domainEdit);
    }
  } catch (error) {
    console.error(error);
  }
}




