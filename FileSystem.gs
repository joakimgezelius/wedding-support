class Folder {
  
  constructor (folder) {
    this._folder = folder;
    this._trace = `{Folder ${this.id} ${this.name}}`;
    trace("NEW " + this.trace);
  }
  
  fileExists(name) {
    var files = this.folder.getFilesByName(name);
    result = (files.hasNext()) ? true : false;
    trace(`fileExists(${name}) in ${folder.trace} --> ${result}`);
    return result;
  }
  
  get folder()  { return this._folder; }
  get parents() { return this._folder.getParents(); }
  get parent()  { return new Folder(this.parents.next()); }  
  get id()      { return this._file.getId(); }
  get name()    { return this._folder.getName() };
  get trace()   { return this._trace; }

} // Folder


class File {

  constructor(file) {
    this._file = file;
    this._trace = `{File ${this.id} ${this.name}}`;
    trace("NEW " + this.trace);
  }
  
  get file()    { return this._file; }
  get parent()  { return new Folder(this.parents.next()); }  
  get parents() { return this._file.getParents(); }
  get id()      { return this._file.getId(); }
  get name()    { return this._file.getName(); }
  get trace()   { return this._trace; }
     
} // File

