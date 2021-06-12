//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/base/user
//
class User {

  constructor(nativeUser) {
    this._nativeUser = nativeUser;
    this._email = nativeUser.getEmail();
    this._isDeveloper = User.developers.includes(this.email.toLowerCase());
    this._trace = `{User ${this.email} (${this.isDeveloper ? "" : "not "}deveoloper)}`;
    trace(`NEW ${this.trace}`);
  }

  static get active() {
    if (User._active === null) {
      User._active = new User(Session.getActiveUser());
    }
    let user = User._active;
    trace(`User.getActive --> ${user.trace}`);
    return user;
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
  ];

