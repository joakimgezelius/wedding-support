//----------------------------------------------------------------------------------------
// Wrapper for https://developers.asana.com/docs/projects
// Project creation : https://developers.asana.com/docs/create-a-project

ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";
WORKSPACE_ID = "1200711902496585";            // Hour Productions Testing

class Asana {

  static getUrl(method) {
    return `${Asana.baseUrl}/${method}`;
  }

  static createProject() {
    let body = {
     "data": {
       "archived": false,
       "color": "dark-red",
       "current_status": { 
         "author": {
           "name": "Furkan Shaikh"
         },
         "color": "green",
         "created_by": {
           "name": "Furkan Shaikh"
         },
         "html_text": "<body>The project is moving forward according to plan...</body>",
         "text": "The project is moving forward according to plan...",
         "title": "Status Update - Aug 04"
       },
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": "Large Weddings",
       "notes": "These are things we need to purchase.",
       "public": true,
      } 
    };
    let options = {
      "method" : "POST",
      "payload": JSON.stringify(body),
      "headers": {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      },
    };      
    let url = Asana.getUrl("?workspace=1200711902496585");
    let response = UrlFetchApp.fetch(url,options);
    trace(`Asana.createProject --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }

}

Asana.baseUrl = "https://app.asana.com/api/1.0/projects";














