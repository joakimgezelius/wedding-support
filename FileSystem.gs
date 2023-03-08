//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/drive/folder
//
class Folder {
  
  constructor(nativeFolder) {
    this._nativeFolder = nativeFolder;
    this._name = this.nativeFolder.getName();
    if (!this.parents.hasNext() && this.name === "Drive") { // This is the root in a shared drive)
      this._name = SharedDrive.GetNameById(parent.getId());
    }
    this._path = this.getPath();
    this._trace = `{Folder ${this.id} "${this.path}"}`;
    trace(`NEW ${this.trace}`);
  } 

  getPath() {
    let fullName = "/" + this.name;
    let parents = this.parents;
    while (parents.hasNext()) { // We're not yet at the root, keep going up the hierarchy (down the tree)
      let parent = parents.next();
      let name = parent.getName();
      parents = parent.getParents();
      if (!parents.hasNext() && name === "Drive") { // We've reached the root in a shared drive
        name = SharedDrive.GetNameById(parent.getId());
      }
      fullName = "/" + name + fullName;
    }
    return fullName;
  }

  static getById(folderId) {
    trace(`> Folder.getById(${folderId})`);
    let nativeFolder = DriveApp.getFolderById(folderId);
    let folder = new Folder(nativeFolder);
    trace(`< Folder.getById(${folderId}) --> ${folder.trace}`);
    return folder;
  }  

  static getByUrl(url) {
    trace(`> Folder.getByUrl(${url})`);
    let folder = null;
    try {
      let folderId = Folder.getIdFromUrl(url);
      folder = Folder.getById(folderId);
    } catch (error) {
      trace(`Folder.getByUrl(${url}), caught exception: "${error}"`);     
    }
    trace(`< Folder.getByUrl(${url}) --> ${folder !== null ? folder.trace : "null"}`);
    return folder;
  }

  static getIdFromUrl(url) {
    // Extract the file id with regular expression, as follows:
    // [-\w]  This matches hyphen and the w character class (A-Za-z0-9_)
    // {19,}  Quantifier matching any sequence of 19+ occurrances of [-\w]
    // Reference: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_Expressions
    return url.match(/[-\w]{19,}/);   
  }
  
  getSubfolder(name) { // NOTE: we assume there is at most one subfolder with the given name in the folder
    // https://developers.google.com/apps-script/reference/drive/folder#getFoldersByName(String)
    let nativeFolderIterator = this.nativeFolder.getFoldersByName(name);
    let subfolder = (nativeFolderIterator.hasNext()) ? new Folder(nativeFolderIterator.next()) : null;
    trace(`getSubfolder(${name}) in ${this.trace} --> ${subfolder === null ? "null (not found)" : subfolder.trace}`);
    return  subfolder;
  }
  
  getFile(name) { // NOTE: we assume there is at most one file with the given name in the folder
    // https://developers.google.com/apps-script/reference/drive/folder#getFilesByName(String)
    let nativeFileIterator = this.nativeFolder.getFilesByName(name);
    let file = (nativeFileIterator.hasNext()) ? new File(nativeFileIterator.next()) : null;
    trace(`getFile(${name}) in ${this.trace} --> ${file === null ? "null (not found)" : file.trace}`);
    return file;
  }

  fileExists(name) {
    // https://developers.google.com/apps-script/reference/drive/folder#getFilesByName(String)
    let files = this.nativeFolder.getFilesByName(name);
    let result = (files.hasNext()) ? true : false;
    trace(`fileExists(${name}) in ${this.trace} --> ${result}`);
    return result;
  }

  folderExists(name) {    //Check whether folder exists or not
    // https://developers.google.com/apps-script/reference/drive/folder#getFoldersByName(String)
    let folders = this.nativeFolder.getFoldersByName(name);
    let result = (folders.hasNext()) ? true : false;
    trace(`folderExists(${name}) in ${this.trace} --> ${result}`);
    return result;
  }

  copyTo(destination, name = this.name) {
    trace(`> Folder.copyTo(${destination.trace}) this=${this.trace}`);
    let copy = destination.createFolder(name); // Create a copy of this folder inside the destination folder
    // here, loop over files, and copy them over
    // https://developers.google.com/apps-script/reference/drive/folder#getFiles()
    let files = this.nativeFolder.getFiles();
    while (files.hasNext()) {                  // Determines whether calling next() will return an item.
      let file = new File(files.next());       // Gets the next item in the collection of files.      
      file.copyTo(copy);                       // Copy to the newly created folder
    }
    // next, loop over folders, and recursively copy them over
    // https://developers.google.com/apps-script/reference/drive/folder#getfolders
    let folders =  this.nativeFolder.getFolders();
    while (folders.hasNext()) {                     
      let folder = new Folder(folders.next()); 
      folder.copyTo(copy);
    }
    trace(`< Folder.copyTo --> ${copy.trace}`);
    return copy;
  }

  // Generic recursive folder tree walker, pass in a context object that defines the actions taken
  //
  recursiveWalk(context, itemCount = 0, level = 0) {
    trace(`> Folder.recursiveWalk ${this.trace} level ${level}`);
    // https://developers.google.com/apps-script/reference/drive/folder#getfolders
    let subFolders =  this.nativeFolder.getFolders();
    while (subFolders.hasNext()) {                    // Determines whether calling next() will return an item                 
      let subFolder = new Folder(subFolders.next());  // Gets the next item in the collection of folders
      trace(`Folder.recursiveWalk (level ${level}) found subfolder (item ${itemCount}): ${subFolder.name}`);
      context.action(subFolder, itemCount++, level);
      itemCount = subFolder.recursiveWalk(context, itemCount, level + 1);
    }
    // https://developers.google.com/apps-script/reference/drive/folder#getFiles()
    let files = this.nativeFolder.getFiles();
    while (files.hasNext()) {                 // Determines whether calling next() will return an item.
      let file = new File(files.next());      // Gets the next item in the collection of files.
      trace(`Folder.recursiveWalk (level ${level}) found file (item ${itemCount}): ${file.name}`);
      context.action(file, itemCount++, level);
    }
    trace(`< Folder.recursiveWalk ${this.trace}  level ${level} -> itemCount: ${itemCount}`);
    return itemCount;
  }

  createFolder(name) {
    trace(`createFolder("${name}) in ${this.trace}`);
    // https://developers.google.com/apps-script/reference/drive/folder#createFolder(String)
    let newFolder = this._nativeFolder.createFolder(name);
    return new Folder(newFolder);
  }
  
  // Create a shortcut in the folder to either a file or a folder, optionally giving it a new name
  // usage: myFolder.createShortcut(target[, newName]);
  //
  createShortcut(target, shortcutName = target.name) {
    trace(`createShortcut to ${target.trace}, name: ${shortcutName} in ${this.trace}`);
    if (this.fileExists(shortcutName)) { // Avoid creating multiple shortcuts with the same name
      Error.fatal(`File "${shortcutName}" already exists in folder "${this.path}"`);
    }
    // https://developers.google.com/apps-script/reference/drive/folder#createshortcuttargetid
    let shortcut = new File(this._nativeFolder.createShortcut(target.id)); // The shortcut is a file, even if it is a shortcut to a folder
    shortcut.name = shortcutName; // Rename the shortcut if a different name was given
    return shortcut;
  }

  get nativeFolder() { return this._nativeFolder; }
  get parents()      { return this.nativeFolder.getParents(); }
  get parent()       { return new Folder(this.parents.next()); }  
  get id()           { return this.nativeFolder.getId(); }
  get url()          { return this.nativeFolder.getUrl(); }
  get name()         { return this._name; }
  get path()         { return this._path; }
  get owner()        { return this.nativeFolder.getOwner()?.getEmail() ?? null; } // Owner email address
  get trace()        { return this._trace; }

  set owner(emailAddress) {
    trace(`set owner of ${this.trace} to ${emailAddress}`);
    this.nativeFolder.setOwner(emailAddress);
  }

} // Folder


//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/drive/file
//
class File {

  constructor(nativeFile) {
    this._nativeFile = nativeFile;
    trace("NEW " + this.trace);
  }
  
  static getById(fileId) {
    trace(`> File.getById(${fileId})`);
    let nativeFile = DriveApp.getFileById(fileId);
    let file = new File(nativeFile);
    trace(`< File.getById(${fileId}) --> ${file.trace}`);
    return file;
  }

  static getByUrl(url) { 
    trace(`> File.getByUrl(${url})`);
    let file = null;
    try {
      let fileId = File.getIdFromUrl(url);
      file = File.getById(fileId);
    } catch (error) {
      trace(`File.getByUrl(${url}), caught exception: "${error}"`);     
    }
    trace(`< File.getByUrl(${url}) --> ${file !== null ? file.trace : "null"}`);
    return file;
  }

  static getIdFromUrl(url) {
    // Extract the file id with regular expression, as follows:
    // [-\w]  This matches hyphen and the w character class (A-Za-z0-9_)
    // {25,}  Quantifier matching any sequence of 25+ occurrances of [-\w]
    // Reference: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_Expressions
    return url.match(/[-\w]{25,}/);   
  }  
  
  copyTo(folder, newName = this.name) {
    trace(`Making copy of ${this.trace} in ${folder.name}, new name: "${newName}"`);
    let newFile = this._nativeFile.makeCopy(newName, folder.nativeFolder);
    return new File(newFile);
  }
  
  get nativeFile() { return this._nativeFile; }
  get parent()     { return new Folder(this.parents.next()); }  
  get parents()    { return this.nativeFile.getParents(); }
  get id()         { return this.nativeFile.getId(); }
  get url()        { return this.nativeFile.getUrl(); }
  get name()       { return this.nativeFile.getName(); }
  get owner()      { return this.nativeFile.getOwner()?.getEmail() ?? null; } // Owner email address
  get trace()      { return `{File ${this.id} "${this.name}"}`; }

  set owner(emailAddress) {
    trace(`set owner of ${this.trace} to ${emailAddress}`);
    this.nativeFile.setOwner(emailAddress);
  }

  // name property setter function, renames the file
  // usage: myFile.name = "new name";
  set name(newName) {
    // https://developers.google.com/apps-script/reference/drive/file#setName(String)
    trace(`set name of  ${this.trace} to ${newName}`);
    this.nativeFile.setName(newName);
  }

  grantImportAccess() {
    // From the forums:
    //  - https://stackoverflow.com/questions/28038768/how-to-allow-access-for-importrange-function-via-apps-script/32474761#32474761
    //  - https://stackoverflow.com/questions/25178205/how-can-i-make-an-apps-script-to-instantly-allow-access-to-all-imported-elements
    // https://developers.google.com/apps-script/reference/drive/file#setSharing(Access,Permission)
    trace(`file.grantImportAccess ${this.trace}`);
    this.nativeFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }

} // File

//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/drive/api/v3/reference/drives
//

class SharedDrive {

  constructor(id, name) {
    this._id = id;
    this._name = name;
    this._trace = `{SharedDrive ${name}, id=${id}}`;
    trace("NEW " + this.trace);
  }

  get id()    { return this._id; }
  get name()  { return this._name; }
  get trace() { return this._trace; }

  static get list() {
    if (typeof SharedDrive._list === "undefined") {
      // List of shared drives, as an array of { id, name } elements:
      //  inspiration: https://yagisanatode.com/2021/07/26/get-a-list-of-google-shared-drives-by-id-and-name-in-google-apps-script/
      SharedDrive._list = Drive.Drives.list({ maxResults: 50 }).items.map(drive => ({id:drive.id,name:drive.name}));
      trace("Initiate SharedDrive.list, entries:");
      for (const [id, drive] of SharedDrive._list.entries()) {
        trace(`  id: ${id} => name: ${drive.name}`);
      }
    }
    return SharedDrive._list;
  }

  static GetById(lookup) {
    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/find
    const drive = SharedDrive.list.find( ({ id }) => id === lookup );
    return new SharedDrive(drive.id, drive.name);
  }

  static GetByName(lookup) {
    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/find
    const drive = SharedDrive.list.find( ({ name }) => name === lookup );
    return new SharedDrive(drive.id, drive.name);
  }

  static GetNameById(lookup) {
    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/find
    const drive = SharedDrive.list.find( ({ id }) => id === lookup );
    return (drive === null ? "[unknown drive]" : drive.name);
  }
}


