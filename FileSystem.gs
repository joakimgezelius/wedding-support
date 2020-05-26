//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/drive/folder
//
class Folder {
  
  constructor (nativeFolder) {
    this._nativeFolder = nativeFolder;
    this._trace = `{Folder ${this.id} ${this.name}}`;
    trace("NEW " + this.trace);
  }
  
  fileExists(name) {
    var files = this.nativeFolder.getFilesByName(name);
    result = (files.hasNext()) ? true : false;
    trace(`fileExists(${name}) in ${this.trace} --> ${result}`);
    return result;
  }

  static getById(folderId) {
    return new Folder(DriveApp.getFolderById(folderId));
  }
  
  get nativeFolder() { return this._nativeFolder; }
  get parents()      { return this.nativeFolder.getParents(); }
  get parent()       { return new Folder(this.parents.next()); }  
  get id()           { return this.nativeFolder.getId(); }
  get name()         { return this.nativeFolder.getName() };
  get trace()        { return this._trace; }

} // Folder


//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/drive/file
//
class File {

  constructor(nativeFile) {
    this._nativeFile = nativeFile;
    this._trace = `{File ${this.id} ${this.name}}`;
    trace("NEW " + this.trace);
  }
  
  static getById(fileId) {
    return new File(DriveApp.getFileById(fileId));
  }
  
  makeCopy(newName, folder) {
    trace(`Making copy of ${this.trace} in ${folder.name}, new name: "${newName}"`);
    let newFile = this._nativeFile.makeCopy(newName, folder.nativeFolder);
    return new File(newFile);
  }
  
  get nativeFile() { return this._nativeFile; }
  get parent()     { return new Folder(this.parents.next()); }  
  get parents()    { return this.nativeFile.getParents(); }
  get id()         { return this.nativeFile.getId(); }
  get name()       { return this.nativeFile.getName(); }
  get trace()      { return this._trace; }
     
} // File

