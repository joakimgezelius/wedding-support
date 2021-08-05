ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";  // Personal access token
WORKSPACE_ID = "1200711902496585";                                     // Hour Productions Testing

class Asana {

  static getProjectUrl(method) {
      return `${Asana.projectUrl}/${method}`;      
  }

  static getTaskUrl(method) {
      return `${Asana.taskUrl}/${method}`;      
  }
  
}

Asana.projectUrl = "https://app.asana.com/api/1.0/projects";     // For basic project operations
Asana.taskUrl    = "https://app.asana.com/api/1.0/tasks";        // For creating the task in Asana


//=========================================================================================================
// Wrapper for project - https://developers.asana.com/docs/projects

class Project {

  static create() {
    let body = {
     "data": {
       "archived": false,
       "color": "dark-red",
       "current_status": { 
          "author": {
            "name": "Furkan Shaikh"
          },
          "color": "green",           // Statuses : Green - On Track, Red - Off Track, Blue - On Hold, Yellow - At Risk 
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
      }
    };      
    let url = Asana.getProjectUrl("?workspace=1200711902496585");
    let response = UrlFetchApp.fetch(url,options);
    trace(`Project.create --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }

}


//=========================================================================================================
// Wrapper for Task - https://developers.asana.com/docs/tasks

class Task {

static create() {
    let body = {
      "data": {
        "approval_status": "pending",     // approved, rejected, changes_requested
        "assignee": "me",
        "assignee_status": "upcoming",    // today, later, new, inbox
        "completed": false,
        "due_on": "2021-08-04",
        "html_notes": "<body>Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are <em>Projects and Tasks</em></body>",
        "name": " Add classes for Project & Task entity",
        "notes": "Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are Projects and Tasks",
        "projects": [
          "1200714734875880"        // project_gid for creating task under it
        ],
        "resource_subtype": "default_task",   //milestone, approval, section
      }
    };
    let options = {
      "method" : "POST",
      "payload": JSON.stringify(body),
      "headers": {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let url = Asana.getTaskUrl("?workspace=1200711902496585");
    let response = UrlFetchApp.fetch(url,options);
    trace(`Task.create --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }

}















