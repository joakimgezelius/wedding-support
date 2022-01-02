
function onGetFileInfo() {
  trace("onGetFileInfo");
  let url = Range.getByName("FileUrl").value;
  let file = File.getByUrl(url);
  if (file !== null) {
    Dialog.notify("File Info", file.trace);      
  } else {
    Dialog.notify("File Info", `${url} is not a valid file URL`);
  }
}

function onGetFolderInfo() {
  trace("onGetFolderInfo");
  applyFolderAction();
}

function onTransferFileOwnership() {
  trace("onTransferFileOwnership");
  let newOwner = Range.getByName("NewOwner").value;
  applyFolderAction(newOwner);
}

function applyFolderAction(newOwner = null) {
  trace("onGetFolderInfo");
  let folderInfoRange = Range.getByName("FolderInfo");
  folderInfoRange.clear();
  // Get a two-dimensional array representing the values of the range:
  // https://developers.google.com/apps-script/reference/spreadsheet/range?hl=en#getvalues
  let folderInfo = folderInfoRange.values; 
  let row = 0;
  let url = Range.getByName("SourceFolderUrl").value;
  let folder = Folder.getByUrl(url);
  if (folder !== null) {
    folderInfo[row++][0] = folder.path;
    if (newOwner !== null && Dialog.confirm("Transfer Ownership", `Are you sure you want to recursively transfer ownership of ${folder.path} to ${newOwner}?`)) {
      let context = new RecursiveWalkerActionContext(folderInfo, row, 2, newOwner);
      folder.recursiveWalk(context);
    }
  } else {
    //folderInfo[row++][0] = `${url} is not a valid folder URL`;
    Error.fatal(`${url} is not a valid folder URL`);
  }
  // https://developers.google.com/apps-script/reference/spreadsheet/range?hl=en#setvaluesvalues
  folderInfoRange.values = folderInfo;
}

function onMoveToSharedDrive() {
  trace("onMoveToSharedDrive");
  Dialog.notify("onMoveToSharedDrive", "onMoveToSharedDrive")
}

class RecursiveWalkerActionContext {
  constructor(infoArray, rowOffset = 0, columnOffset = 0, newOwner = null) {
    this.infoArray = infoArray;
    this.newOwner = newOwner;
    this.rowOffset = rowOffset;
    this.columnOffset = columnOffset;
    this.me = User.active.email;
    trace(`NEW ${this.trace}`);
  }

  get trace() { return `RecursiveWalkerActionContext offsets: ${this.rowOffset} ${this.columnOffset} me: ${this.me} newOwner: ${this.newOwner ?? "-"}` }

  action(object, itemCount, level) {
    let row = itemCount + this.rowOffset;
    if (object.owner == this.me) { // We're the owner of the file, meaning operations can be applied...
      if (this.newOwner != null) {
        //object.owner = newOwner;
        this.infoArray[row][0] = `me -> ${this.newOwner}`;
      } else {
        this.infoArray[row][0] = "me";
      }
    } else {
      this.infoArray[row][0] = object.owner ?? "-";
    }
    let column = level + this.columnOffset;
    if (object instanceof Folder) {
      this.infoArray[row][column] = object.name + "/";
    } else { // Else it is a file
      this.infoArray[row][column] = object.name;
    }
  }

}

class TypeUtils {

  /** 
   * Returns true if the given test value is an object; false otherwise.
   */
  static isObject(test) {
    return Object.prototype.toString.call(test) === '[object Object]';
  }

  /** 
   * Returns true if the given test value is an array containing at least one object; false otherwise.
   */
  static isObjectArray(test) {
    for (var i = 0; i < test.length; i++) {
      if (TypeUtils.isObject(test[i])) {
        return true; 
      }
    }  
    return false;
  }

}

class StringUtils {

  /** 
   * Encodes the given value to use within a URL.
   *
   * @param {value} the value to be encoded
   * 
   * @return the value encoded using URL percent-encoding
   */
  static urlEncode(value) {
    return encodeURIComponent(value.toString());  
  }

  /** 
   * Locates the index where the two strings values stop being equal, stopping automatically at the stopAt index.
   */
  static findEqualityEndpoint(string1, string2, stopAt) {
    if (!string1 || !string2) {
      return -1; 
    }
    
    var maxEndpoint = Math.min(stopAt, string1.length, string2.length);
    
    for (var i = 0; i < maxEndpoint; i++) {
      if (string1.charAt(i) != string2.charAt(i)) {
        return i;
      }
    }
    
    return maxEndpoint;
  }

  /** 
   * Converts the text to title case.
   */
  static toTitleCase(text) {
    if (text == null) {
      return null;
    }   
    return text.replace(/\w\S*/g, function(word) { return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); });
  }

}
