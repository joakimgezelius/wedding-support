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
