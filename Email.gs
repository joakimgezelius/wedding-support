emailTemplateDatabaseSheetId = "1UqPLV5754d5I4DNCJ65ab6mo4bFQEMC6XgyJYcpoR1g";
emailTemplateFolderId = "1v3v7dFW10fc-T-LAEPsR1qkZwOAFUytS";
defaultEmailTemplateListName = "EnquiryEmails";

function onCreatetEmailDraft0() { EmailTemplateList.createtEmailDraft(0); }
function onCreatetEmailDraft1() { EmailTemplateList.createtEmailDraft(1); }
function onCreatetEmailDraft2() { EmailTemplateList.createtEmailDraft(2); }
function onCreatetEmailDraft3() { EmailTemplateList.createtEmailDraft(3); }
function onCreatetEmailDraft4() { EmailTemplateList.createtEmailDraft(4); }
function onCreatetEmailDraft5() { EmailTemplateList.createtEmailDraft(5); }
function onCreatetEmailDraft6() { EmailTemplateList.createtEmailDraft(6); }
function onCreatetEmailDraft7() { EmailTemplateList.createtEmailDraft(7); }
function onCreatetEmailDraft8() { EmailTemplateList.createtEmailDraft(8); }
function onCreatetEmailDraft9() { EmailTemplateList.createtEmailDraft(9); }


function onSendFirstEmail() {
  trace("onSendFirstEmail");
  let email = new Email("ClientEmail1");
  trace("NEW " + email.trace);
  email.send();
}


//================================================================================================

class EmailTemplateList {

  constructor(emailTemplateListName = defaultEmailTemplateListName) {
//  this.databaseSheet = SpreadsheetApp.openById(Email.databaseSheetId);
    this.emailTemplates = [];
    this.listRange = Range.getByName(emailTemplateListName).loadColumnNames();
    for (let rowOffset = 0; rowOffset < this.listRange.height; rowOffset++) {
      let row = new EmailTemplateRow(this.listRange, rowOffset);
      this.emailTemplates[row.id] = row;
    }
    trace("NEW " + this.trace);
  }
  
  populateMenu(menu) {
    trace(`${this.trace}.populateMenu`);
    this.emailTemplates.forEach(item => menu.addItem(item.name, globalLibName + ".onCreatetEmailDraft" + item.id));
    return menu;
  }

  static get singleton() {
    if (EmailTemplateList._singleton === undefined) {
      trace("EmailTemplateList.singleton, initialize");
      EmailTemplateList._singleton = new EmailTemplateList();
    }
    return EmailTemplateList._singleton;
  }
  
  
  static createtEmailDraft(emailId) {
    trace("createtEmailDraft");
    let email = new Email("ClientEmail1");
    trace("NEW " + email.trace);
    email.createDraft();
  }

  get trace() {
    return `{EmailTemplateList ${this.listRange.trace}}`;
  }
  
} // EmailTemplateList


//================================================================================================

class EmailTemplateRow extends EventRow {

  constructor(containerRange, rowOffset) {
    super(containerRange, rowOffset);
    this._trace = `{EmailTemplateRow ${this.id} ${this.name}}`;
    trace("NEW " + this.trace);
  }

  get id()      { return this.get("Id", "string"); }
  get name()    { return this.get("Name", "string"); }
  get file()    { return this.get("File", "string"); }
  get trace()   { return this._trace; }
  
} // EmailTemplateRow


//================================================================================================

class Email {
  
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
  
} // Email

class EmailCompositionMatrix {

  constructor(rangeName =  defaultEmailCompositionMatrix) {
  }

  
}
