var EXCLUDE_EMAILS = [
    "friend1@example.com", 
    "friend2@example.com",
]; // Add your exclude emails here, must have comma at the end of the line

var EXCLUDE_DOMAINS = [
    "gmail.com", 
    "icloud.com", 
    "hotmail.com", 
    "outlook.com", 
    "yahoo.com",
]; // Add the domains of default email services here, must have comma at the end of the line

var LABEL_NAME = "Contact_Saved";
var SHEET_NAME = "Contacts";


function getOrCreateLabel() {
  var label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (label) {
    return label;
  } else {
    return GmailApp.createLabel(LABEL_NAME);
  }
}

function getOrCreateSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
        var sheet = spreadsheet.getSheets()[0];
        sheet.appendRow(['No', 'Name', 'Email', 'CompanyLink']);
        sheet.setName(SHEET_NAME);
        return sheet;
    } else if (!sheet.getFilter()) {
        var range = sheet.getRange("A1:D");
        var filter = range.createFilter();
    }
    return sheet;
}

function isBusinessEmail(email) {
    var domain = email.substring(email.indexOf('@') + 1);
    return !EXCLUDE_DOMAINS.includes(domain);
}

function isExceptEmail(email) {
    return EXCLUDE_EMAILS.includes(email);
  }

function getContactInfo(data) {
    var name = "";
    var email = "";
    
    var matches = data.match(/\s*"?([^"]*)"?\s+<(.+)>/);
    
    if (matches) {
      name = matches[1].trim(); 
      email = matches[2];
    }
    else {
      name = "N/A";
      email = data;
    }

    var companyEmail = isBusinessEmail(email) ? email.substring(email.indexOf('@') + 1) : 'N/A';
    
    return {name: name, email: email, companyEmail: companyEmail};
}

function getContactsFromGmail() {
  var label = getOrCreateLabel();
  var threads = GmailApp.search("-label:" + LABEL_NAME);
  var contactList = [];
  var sheet = getOrCreateSheet();
  var startRow = sheet.getLastRow() + 1;

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var fromContact = getContactInfo(messages[j].getFrom());
      var toContact = getContactInfo(messages[j].getTo());
      
      if (!isExceptEmail(fromContact.email) && !contactList.some((contact) => contact.email === fromContact.email)) {
        contactList.push(fromContact);
      }

      if (!isExceptEmail(toContact.email) && !contactList.some((contact) => contact.email === toContact.email)) {
        contactList.push(toContact);
      }
    }
    threads[i].addLabel(label);
  }

  for (var k = 0; k < contactList.length; k++) {
    var contact = contactList[k];
    sheet.getRange(startRow + k, 1).setValue(startRow + k - 1); // Set the No value
    sheet.getRange(startRow + k, 2).setValue(contact.name);
    sheet.getRange(startRow + k, 3).setValue(contact.email);
    sheet.getRange(startRow + k, 4).setValue(contact.companyEmail);
  }
}

function createButton() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Contacts')
        .addItem('Get Contacts', 'getContactsFromGmail')
        .addToUi();
}

function onOpen() {
    createButton();    
}
