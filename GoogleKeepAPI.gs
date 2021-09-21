class GoogleKeep {

  static create() {

    // https://developers.google.com/keep/api/reference/rest/v1/notes#resource:-note

    const date = Utilities.formatDate(new Date(),"GMT+2","yyyy-MM-dd HH:mm");
    let theAccessToken = ScriptApp.getOAuthToken();
    //let apiKey = 'AIzaSyDgDwYOt_P0nJFtq01vF4XS92ko-4rPUrI';
    let body = {
      "name": "From Client Sheet",                                // The resource name of this note
      "createTime": date,                                         // When this note was created.
      "updateTime": date,                                         // When this note was last modified.
      //"trashTime": string,                                      // When this note was trashed.
      "trashed": false,
      "attachments": [
        {
          //object (Attachment)
        }
      ],
      "permissions": [
        {
          //object (Permission)
        }
      ],
      "title": "Test - Keep Note 1",                                // The title of the note. Length must be less than 1,000 characters
      "body": {
        //object (Section)
      }      
    }
    let options = {
      "method" : "POST",
      "muteHttpExceptions": true,
      "payload": JSON.stringify(body),
      "headers": {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " +  theAccessToken
      }
    }
    let url = 'https://keep.googleapis.com/v1/notes';
    let response = UrlFetchApp.fetch(url,options);
    trace(`Response : ${response.getContentText()}`);
  }

}