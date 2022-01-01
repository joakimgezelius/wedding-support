//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/base/user
//
class User {

  constructor(nativeUser) {
    this._nativeUser = nativeUser;
    this._email = nativeUser?.getEmail() ?? "-";
    this._isDeveloper = User.developers.includes(this.email.toLowerCase());
    this._trace = `{User ${this.email} (${this.isDeveloper ? "" : "not "}deveoloper)}`;
    trace(`NEW ${this.trace}`);
  }

  static get active() {
    return User._active ?? (User._active =  new User(Session.getActiveUser()));
  }

  equals(otherUser) {
    return this.nativeUser === otherUser.nativeUser;
  }

  get nativeUser()  { return this._nativeUser; }
  get email()       { return this._email; }
  get isDeveloper() { return this._isDeveloper; }
  get trace()       { return this._trace; }

} // User

User._active = null; // Static property

User.developers = [ 
  "joakim.gezelius@gmail.com", 
  "iamfurkanshaikh@gmail.com", 
  "monica@hour.events",
  "it@hour.events",
  ];

