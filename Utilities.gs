function WalkFolder(folderUrl) {
  trace(`WalkFolder(${folderUrl})`);
  let folder = Folder.getByUrl(folderUrl);
  return folder.trace;
}

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
  let drives = SharedDrive.list;
  SharedDrive.GetById("0AIvL-1_7EwaoUk9PVA");
  let url = Range.getByName("SourceFolderUrl").value;
  let folder = Folder.getByUrl(url);
  if (folder !== null) {
    Dialog.notify("Folder Info", folder.trace);
  } else {
    Dialog.notify("Folder Info", `${url} is not a valid file URL`);     
  }
}

function onMoveToSharedDrive() {
  trace("onMoveToSharedDrive");
  Dialog.notify("onMoveToSharedDrive", "onMoveToSharedDrive")
}

function onTransferFileOwnership() {
  trace("onTransferFileOwnership");
  Dialog.notify("onTransferFileOwnership", "onTransferFileOwnership")
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
