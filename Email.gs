function onCreateFirstEmailDraft() {
  trace("onCreateFirstEmailDraft");
  let email = new Email("ClientEmail1");
  trace("NEW " + email.trace);
  email.createDraft();
}

function onSendFirstEmail() {
  trace("onSendFirstEmail");
  let email = new Email("ClientEmail1");
  trace("NEW " + email.trace);
  email.send();
}

class Email {

  static init(emailListName) {
    Email.databaseSheetId = "1UqPLV5754d5I4DNCJ65ab6mo4bFQEMC6XgyJYcpoR1g";
//  Email.databaseSheet = SpreadsheetApp.openById(Email.databaseSheetId);
    Email.databaseRange = Range.getByName(emailListName);
  }
  
  static populateMenu(menu, emailListName) {
    trace("Email.populateMenu " + emailListName);
    Email.init(emailListName);
    Email.databaseRange.rewind();
    for (var rowOffset = 0; rowOffset < Email.databaseRange.height; rowOffset++) {
      let emailData = Email.databaseRange.getNextRowValues();
      console.log(emailData);
//      trace(" - " + emailData[1]);
//      menu.addItem(emailData[1], globalLibName + ".Email.init");
//      let rowRange = this.databaseRange.range.offset(rowOffset, 0, 1);
//      let name = new EventRow(this.data[rowOffset], rowOffset, rowRange);
    }      
    menu.addItem("Refresh", globalLibName + ".Email.init");
  }
  
  constructor(name = "") {
    if (name != "") {
      this.recipient = "monicacoumbe@gmail.com";
      this.subject = "First Email From Template";
      this.emailTemplate = HtmlService.createTemplateFromFile(name);
      
    }
    trace("NEW " + this.trace);
  }
  
  prepareMessage() {
    trace(`Email.prepareMessage ${this.trace}`);
    this.emailTemplate.firstName = "Monica"; // this.receiver.firstName;
    this.htmlMessage = this.emailTemplate.evaluate().getContent();
    this.options = {name: "Hour Weddings", htmlBody: this.htmlMessage }
  }
  
  createDraft() {
    trace(`Email.createDraft ${this.trace}`);
    this.prepareMessage();
    GmailApp.createDraft(this.recipient, this.subject, "HTML message only", this.options);
  }
  
  send() {
    trace(`Email.createDraft ${this.trace}`);
    this.prepareMessage();
    //GmailApp.send(this.recipient, this.subject, "HTML message only", this.options);
  }
  
  get trace() {
    return `Email ${this.to} ${this.subject}`;
  }
  
}
