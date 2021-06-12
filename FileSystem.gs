//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/drive/folder
//
class Folder {
  
  constructor(nativeFolder) {
    this._nativeFolder = nativeFolder;
    this._trace = `{Folder ${this.id} "${this.name}"}`;
    trace(`NEW ${this.trace}`);
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
    let folderId = File.getIdFromUrl(url);
    let folder = Folder.getById(folderId);
    trace(`< Folder.getByUrl(${url}) --> ${folder.trace}`);
    return folder;
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
    let files = this.nativeFolder.getFiles();
    while (files.hasNext()) {                  // Determines whether calling next() will return an item.
      let file = new File(files.next());       // Gets the next item in the collection of files.      
      file.copyTo(copy);                       // Copy to the newly created folder
    }
    // next, loop over folders, and recursively copy them over
    let folders =  this.nativeFolder.getFolders();
    while (folders.hasNext()) {                     
      let folder = new Folder(folders.next()); 
      folder.copyTo(copy);
    }
    trace(`< Folder.copyTo --> ${copy.trace}`);
    return copy;
  }

  recursiveWalk() {
    trace(`> Folder.recursiveWalk ${this.trace}`);    
    // https://developers.google.com/apps-script/reference/drive/folder#getFiles()
    let files = this.nativeFolder.getFiles();
    while (files.hasNext()) {                 // Determines whether calling next() will return an item.
      let file = new File(files.next());      // Gets the next item in the collection of files.
      trace(`  found file: ${file.name}`);
    } 
    // https://developers.google.com/apps-script/reference/drive/folder#getfolders
    let subFolders =  this.nativeFolder.getFolders();
    while (subFolders.hasNext()) {                     
      let subFolder = new Folder(subFolders.next());    
      trace(`  found subfolder: ${subFolder.name}`);
      subFolder.recursiveWalk();
    }
    trace(`< Folder.recursiveWalk ${this.trace}`);   
  }

  createFolder(name) {
    trace(`createFolder("${name}) in ${this.trace}`);
    // https://developers.google.com/apps-script/reference/drive/folder#createFolder(String)
    let newFolder = this._nativeFolder.createFolder(name);
    return new Folder(newFolder);
  }
  
  get nativeFolder() { return this._nativeFolder; }
  get parents()      { return this.nativeFolder.getParents(); }
  get parent()       { return new Folder(this.parents.next()); }  
  get id()           { return this.nativeFolder.getId(); }
  get name()         { return this.nativeFolder.getName(); }
  get trace()        { return this._trace; }

} // Folder


//----------------------------------------------------------------------------------------
// Wrapper for https://developers.google.com/apps-script/reference/drive/file
//
class File {

  constructor(nativeFile) {
    this._nativeFile = nativeFile;
    this._trace = `{File ${this.id} "${this.name}"}`;
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
    let fileId = File.getIdFromUrl(url);
    let file = File.getById(fileId);
    trace(`< File.getByUrl(${url}) --> ${file.trace}`);
    return file;
  }

  static getIdFromUrl(url) {
    return url.match(/[-\w]{25,}/);   //returns the file/folder id with regular expression
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
  get name()       { return this.nativeFile.getName(); }
  get trace()      { return this._trace; }
     
} // File

