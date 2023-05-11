var formAccessTries = 0;
var form;
getForm();

function getForm() {
  try {
    form = FormApp.getActiveForm();
  } catch (exception) {
    formAccessTries++;
    if (formAccessTries > 3) throw exception;
    Logger.log(exception);
    setTimeout(getForm, 500);
  }
}

var studentsItem = form.getItemById("24462578"),
  summaryItem = form.getItemById("1821413772"),
  dateTimeItem = form.getItemById("940044187"),
  methodItem = form.getItemById("1701747489"),
  categoryItem = form.getItemById("1607620336"),
  connectItem = form.getItemById("1227903304"),
  notifyItem = form.getItemById("768900339"),
  otherStaffItem = form.getItemById("375842976"),
  staffItem = form.getItemById("1869311159");
var spreadsheet = SpreadsheetApp.openById(form.getDestinationId()),
  studentsSheet = spreadsheet.getSheetByName("Student List"),
  staffSheet = spreadsheet.getSheetByName("Staff List"),
  anecdotalsSheet = spreadsheet.getSheetByName("Anecdotals");

function onOpen() {
  var ui = FormApp.getUi();
  ui.createMenu('Anecdotals')
    .addItem('Update student list on form', 'updateStudentChoices')
    .addItem('Update staff list on form', 'updateStaffChoices')
    .addItem('Restore missing responses','restoreResponses')
    .addToUi();
}

function getItemIds() {
  var items = form.getItems();
  for (var item of items) {
    Logger.log(item.getId() + ": " + item.getTitle());
  }
}

function getEditableUrls() {
  var responses = form.getResponses(new Date("2022-11-10"));
  for (var response of responses) {
    Logger.log(response.getEditResponseUrl());
  }
}

function updateStudentChoices() {
  var studentChoices = studentsSheet.getRange(3, 8, studentsSheet.getLastRow() - 2, 1).getValues();
  studentsItem.asCheckboxItem().setChoiceValues(studentChoices);
}

function updateStaffChoices() {
  var staffChoices = staffSheet.getRange(2, 7, staffSheet.getLastRow() - 1, 1).getValues();
  staffItem.asCheckboxItem().setChoiceValues(staffChoices);
}

function processForm(event) {
  var response = event.response;
  var reporterEmail = response.getRespondentEmail(),
    reporterName = AdminDirectory.Users.get(reporterEmail).name.fullName,
    url = response.getEditResponseUrl(),
    id = response.getId(),
    timestamp = formatDate(response.getTimestamp(), true),
    students = response.getResponseForItem(studentsItem).getResponse(),
    summary = response.getResponseForItem(summaryItem).getResponse(),
    dateTime = formatDate(new Date(response.getResponseForItem(dateTimeItem).getResponse()), true),
    method = response.getResponseForItem(methodItem).getResponse().join(", "),
    category = response.getResponseForItem(categoryItem).getResponse(),
    connected = response.getResponseForItem(connectItem).getResponse(),
    notifyOtherStaff = response.getResponseForItem(otherStaffItem).getResponse() == "Yes";

  try {
    var notify = response.getResponseForItem(notifyItem).getResponse();
    var index = notify.indexOf("Yourself");
    if (index != -1) {
      notify[index] = "Self";
    }
    var notifyStudents = notify.includes("Student(s)"),
      notifyFamilies = notify.includes("Family members"),
      notifyAdvisors = notify.includes("Advisor(s)"),
      notifySelf = notify.includes("Self");
  } catch (error) {
    var notify = [],
      notifyStudents = false,
      notifyFamilies = false,
      notifyAdvisors = false,
      notifySelf = false;
  }
  try {
    var staffMembers = response.getResponseForItem(staffItem).getResponse();
  } catch (error) {
    var staffMembers = [];
  }

  var existingUrls = anecdotalsSheet.getRange("L3:L").getValues();
  var edited = false;
  var previousDates, previousDate;
  var rowsToDelete = [];
  for (var i = 0; i < existingUrls.length; i++) {
    if (existingUrls[i] == url) {
      previousDates = anecdotalsSheet.getRange(i + 3, 11).getNote();
      previousDate = anecdotalsSheet.getRange(i + 3, 11).getValue();
      rowsToDelete.push(i+3);
      edited = true;
    }
  }
  var deletedRows = 0;
  for (var row of rowsToDelete) {
    anecdotalsSheet.deleteRow(row - deletedRows);
    deletedRows++;
    SpreadsheetApp.flush();
  }

  if (edited) {
    if (previousDates == null || previousDates == "") {
      previousDates = "Originally submitted " + previousDate;
    }
    previousDates = "Edited " + timestamp + "\n" + previousDates;
  }

  for (var student of students) {
    var studentName = student.substring(0, student.indexOf("(") - 1),
      studentOSIS = student.substring(student.indexOf("(") + 6, student.indexOf(";")),
      studentGrade = student.substring(student.indexOf(";") + 8, student.indexOf(")"));
    var row = [dateTime, studentName, studentOSIS, studentGrade, category, reporterName, summary, method, connected, notify.join("; ") + (staffMembers.length > 0 ? (notify.length > 0 ? ", " : "") + staffMembers.join("; ") : ""), timestamp, url, reporterEmail];
    anecdotalsSheet.insertRowBefore(3);
    anecdotalsSheet.getRange(3, 1, 1, 13).setValues([row]);
    if (edited) {
      anecdotalsSheet.getRange(3, 11).setNote(previousDates);
    }
  }
  SpreadsheetApp.flush();

  if (!edited && (notifyAdvisors || notifyOtherStaff || notifyFamilies || notifyStudents || notifySelf)) {
    var studentData = studentsSheet.getRange(3, 8, studentsSheet.getLastRow() - 2, 7).getValues();
    var indices = [];
    for (var i = 0; i < studentData.length; i++) {
      if (students.includes(studentData[i][0])) {
        indices.push(i);
      }
    }

    var staffEmails = [];
    var studentEmails = [];
    var familyEmails = [];

    if (notifyStudents) {
      for (var i of indices) {
        studentEmails.push(studentData[i][1]);
      }
    }

    if (notifyAdvisors) {
      for (var i of indices) {
        staffEmails.push(studentData[i][5]);
      }
    }

    if (notifyFamilies) {
      for (var i of indices) {
        familyEmails.push(studentData[i][6]);
      }
    }

    if (notifyOtherStaff) {
      var staffData = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 7).getValues();
      for (var i = 0; i < staffData.length; i++) {
        if (staffMembers.includes(staffData[i][6]) && !staffEmails.includes(staffData[i][6])) {
          staffEmails.push(staffData[i][5]);
        }
      }
    }

    staffEmails.filter(function(el) {return el;} );
    // Email staff members
    if (staffEmails.length > 0) {
      var body = reporterName + " created a new anecdotal: " +
        "\nDate and Time: " + dateTime +
        "\nStudent(s): " + students.join(",") +
        "\nCategory: " + category +
        "\nMethod(s): " + method +
        "\nDescription: " + summary +
        (notify.length > 0 ? ("\nNotified: " + notify) : "") +
        (notifyOtherStaff ? ("\nOther Staff Notified: " + staffMembers.join(",")) : "");
      var htmlBody = "<p><b>" + reporterName + "</b> created a new anecdotal:</p>" +
        "<p><b>Date and Time: </b> " + dateTime + "</p>" +
        "<p><b>Students: </b> <ul><li>" + students.join("</li><li>") + "</li></ul>" +
        "<p><b>Category: </b> " + category + "</p>" +
        "<p><b>Method(s): </b> " + method + "</p>" +
        "<p><b>Description: </b> " + summary + "</p>" +
        (notify.length > 0 ? ("<p><b>Notified: </b> " + notify + "</p>") : "") +
        (notifyOtherStaff ? ("<p>Other Staff Notified: </b> <ul><li>" + staffMembers.join("</li><li>") + "</li></ul>") : "");
      GmailApp.createDraft(staffEmails.join(","), "New Anecdotal by " + reporterName, body, {
        "from": "webmaster@47aslhs.net",
        name: "Anecdotal System",
        replyTo: reporterEmail,
        htmlBody: htmlBody
      }).send();
    }
    // Email students
    if (studentEmails.length > 0) {
      var body = reporterName + " created a new anecdotal about you: " +
        "\nDate and Time: " + dateTime +
        "\nCategory: " + category +
        "\nMethod(s): " + method +
        "\nDescription: " + summary +
        (notify.length > 0 ? ("\nNotified: " + notify) : "") +
        (notifyOtherStaff ? ("\nOther Staff Notified: " + staffMembers.join(",")) : "");
      var htmlBody = "<p><b>" + reporterName + "</b> created a new anecdotal about you:</p>" +
        "<p><b>Date and Time: </b> " + dateTime + "</p>" +
        "<p><b>Category: </b> " + category + "</p>" +
        "<p><b>Method(s): </b> " + method + "</p>" +
        "<p><b>Description: </b> " + summary + "</p>" +
        (notify.length > 0 ? ("<p><b>Notified: </b> " + notify + "</p>") : "") +
        (notifyOtherStaff ? ("<p><b>Other Staff Notified: </b> <ul><li>" + staffMembers.join("</li><li>") + "</li></ul></p") : "");
      for (var email of studentEmails) {
        if (email != null && email != "") {
          GmailApp.createDraft(email, "New Anecdotal by " + reporterName, body, {
            "from": "webmaster@47aslhs.net",
            name: "Anecdotal System",
            replyTo: reporterEmail,
            htmlBody: htmlBody
          }).send();
        }
      }
    }
    // Email families
    for (var i = 0; i < indices.length; i++) {
      if (familyEmails[i] != null && familyEmails[i] != "") {
        var studentName = studentData[i][0].substring(0,studentData)
        var body = reporterName + " created a new anecdotal about " + studentName + ":" +
          "\nDate and Time: " + dateTime +
          "\nCategory: " + category +
          "\nMethod(s): " + method +
          "\nDescription: " + summary +
          (notify.length > 0 ? ("\nNotified: " + notify) : "") +
          (notifyOtherStaff ? ("\nOther Staff Notified: " + staffMembers.join(",")) : "");
        var htmlBody = "<p><b>" + reporterName + "</b> created a new anecdotal about " + studentName + ":</p>" +
          "<p><b>Date and Time: </b> " + dateTime + "</p>" +
          "<p><b>Category: </b> " + category + "</p>" +
          "<p><b>Method(s): </b> " + method + "</p>" +
          "<p><b>Description: </b> " + summary + "</p>" +
          (notify.length > 0 ? ("<p><b>Notified: </b> " + notify + "</p>") : "") +
          (notifyOtherStaff ? ("<p><b>Other Staff Notified: </b> <ul><li>" + staffMembers.join("</li><li>") + "</li></ul></p") : "");
        GmailApp.createDraft(familyEmails[i], "New Anecdotal by " + reporterName, body, {
          "from": "webmaster@47aslhs.net",
          name: "“47” Anecdotal System",
          replyTo: reporterEmail,
          htmlBody: htmlBody
        }).send();
      }
    }
    // Email self
    if (notifySelf) {
      var body = reporterName + " created a new anecdotal: " +
        "\nDate and Time: " + dateTime +
        "\nStudent(s): " + students.join(",") +
        "\nCategory: " + category +
        "\nMethod(s): " + method +
        "\nDescription: " + summary +
        (notify.length > 0 ? ("\nNotified:</b> " + notify) : "") +
        (notifyOtherStaff ? ("\nOther Staff Notified: " + staffMembers.join(",")) : "") +
        "\nEdit URL: " + url;
      var htmlBody = "<p><b>" + reporterName + "</b> created a new anecdotal:</p>" +
        "<p><b>Date and Time: </b> " + dateTime + "</p>" +
        "<p><b>Students: </b> <ul><li>" + students.join("</li><li>") + "</li></ul>" +
        "<p><b>Category: </b> " + category + "</p>" +
        "<p><b>Method(s): </b> " + method + "</p>" +
        "<p><b>Description: </b> " + summary + "</p>" +
        (notify.length > 0 ? ("<p><b>Notified:</b> " + notify + "</p>") : "") +
        (notifyOtherStaff ? ("<p>Other Staff Notified:</b> <ul><li>" + staffMembers.join("</li><li>") + "</li></ul>") : "")+
        "<p><b>Edit URL: </b> " + "<a href=\"" + url + "\">" + url + "</a></p>";
      GmailApp.createDraft(reporterEmail, "Your New Anecdotal on " + dateTime, body, {
        "from": "webmaster@47aslhs.net",
        name: "Anecdotal System",
        htmlBody: htmlBody
      }).send();
    }
  }
}

function restoreResponses() {
  var responses = form.getResponses();
  var studentCount = 0;
  var restoredCount = 0;
  var existingCount = 0;
  for (var response of responses) {
    var url = response.getEditResponseUrl();
    var existingUrls = anecdotalsSheet.getRange("L3:L").getValues();
    var exists = false;
    for (var i = 0; i < existingUrls.length; i++) {
      if (existingUrls[i] == url) {
        exists = true;
        break;
      }
    }
    if (exists) { // This anecdotal is already in the anecdotalsSheet, so don't add it again
      existingCount++;
      continue;
    }
    restoredCount++;

    var reporterEmail = response.getRespondentEmail(),
      reporterName = AdminDirectory.Users.get(reporterEmail).name.fullName,
      id = response.getId(),
      timestamp = formatDate(response.getTimestamp(), true),
      students = response.getResponseForItem(studentsItem).getResponse(),
      summary = response.getResponseForItem(summaryItem).getResponse(),
      dateTime = formatDate(new Date(response.getResponseForItem(dateTimeItem).getResponse()), true),
      method = response.getResponseForItem(methodItem).getResponse().join(", "),
      category = response.getResponseForItem(categoryItem).getResponse(),
      connected = response.getResponseForItem(connectItem).getResponse();

    try {
      var notify = response.getResponseForItem(notifyItem).getResponse();
      var index = notify.indexOf("Yourself");
      if (index != -1) {
        notify[index] = "Self";
      }
    } catch (error) {
      var notify = [];
    }
    try {
      var staffMembers = response.getResponseForItem(staffItem).getResponse();
    } catch (error) {
      var staffMembers = [];
    }

    for (var student of students) {
      var studentName = student.substring(0, student.indexOf("(") - 1),
        studentOSIS = student.substring(student.indexOf("(") + 6, student.indexOf(";")),
        studentGrade = student.substring(student.indexOf(";") + 8, student.indexOf(")"));
      var row = [dateTime, studentName, studentOSIS, studentGrade, category, reporterName, summary, method, connected, notify.join("; ") + (staffMembers.length > 0 ? (notify.length > 0 ? ", " : "") + staffMembers.join("; ") : ""), timestamp, url, reporterEmail];
      anecdotalsSheet.insertRowBefore(3);
      anecdotalsSheet.getRange(3, 1, 1, 13).setValues([row]);
      studentCount++;
    }
    SpreadsheetApp.flush();
  }
  Logger.log("%s anecdotals for %s students restored, %s anecdotals unchanged", restoredCount, studentCount, existingCount);
}

function formatDate(date, includeTime) {
  if (includeTime) return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy.MM.dd hh:mm:ss aaa");
  else return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy.MM.dd EEE");
}
