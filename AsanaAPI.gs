ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";  // Personal access token
WORKSPACE_ID = "1200711902496585";                                     // Hour Productions Testing

class Asana {

  static getProjectUrl(method) {
      return `${Asana.projectUrl}/${method}`;      
  }

  static getTaskUrl(method) {
      return `${Asana.taskUrl}/${method}`;      
  } 

  static getSubtaskUrl(method) {
      return `${Asana.taskUrl}/${method}/subtasks`; 
  }
}

Asana.projectUrl = "https://app.asana.com/api/1.0/projects";     // For basic project operations
Asana.taskUrl    = "https://app.asana.com/api/1.0/tasks";        // For creating the task in Asana

//=========================================================================================================
// Wrapper for project - https://developers.asana.com/docs/projects
// Class Project

class Project {

  static create() {
    let body = {
     "data": {
       "archived": false,
       "color": "dark-red",           // Color for icon
       "current_status": { 
          "author": {
            "name": "Furkan Shaikh"
          },
          "color": "green",           // Status : Green - On Track, Red - Off Track, Blue - On Hold, Yellow - At Risk 
          "created_by": {
            "name": "Furkan Shaikh"
          },
          "html_text": "<body>The project is for the tesing the APIs...</body>",  // Description
          "text": "The project is moving forward according to plan...",
          "title": "Status Update - Aug 05 : Project to test APIs"                // Status title
        },
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": "Large Weddings",                         // name for the project
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
    let url = Asana.getProjectUrl("?workspace=1200711902496585");   //  Workspace_ID where project lies
    let response = UrlFetchApp.fetch(url,options);
    trace(`Project.create --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }

  static update() {
    let body = {
     "data": {
       "archived": false,
       "color": "dark-teal",
       "current_status": { 
          "author": {
            "name": "Furkan Shaikh"
          },
          "color": "blue",
          "created_by": {
            "name": "Furkan Shaikh"
          },
          "html_text": "<body>The project <strong>Status</strong> is updated...</body>",
          "text": "The project status is updated",
          "title": "Status Update - Aug 05"
        },
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": "Hour Weddings",
       "notes": "These are things we need to purchase.",
       "public": true,
      } 
    };
    let options = {
      "method" : "PUT",
      "payload": JSON.stringify(body),
      "headers": {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };      
    let url = Asana.getProjectUrl("1200714734875880");
    let response = UrlFetchApp.fetch(url,options);
    trace(`Project.update --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }

  static destroy() {
    let options = {
      "method" : "DELETE",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };      
    let url = Asana.getProjectUrl("1234567890");         // Use desired {project_gid} to delete project
    UrlFetchApp.fetch(url,options);
  }

}


//=========================================================================================================
// Wrapper for Task - https://developers.asana.com/docs/tasks
// Class Task

class Task {

static create() {
  let body = {
    "data": {
      "approval_status": "pending",         // approved, rejected, changes_requested, pending
      "assignee": "me",
      "assignee_status": "upcoming",        // today, later, new, inbox, upcoming
      "completed": false,
      "due_on": "2021-08-09",
      "html_notes": "<body>Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are <em>Projects and Tasks</em></body>",                         // description
      "name": "Shopping in XYZ for ABC",    // task title
      "notes": "Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are Projects and Tasks",
      "projects": [
        "1200714734875880"                  // project_gid for creating task under it
      ],
      "resource_subtype": "default_task",   //milestone, approval, section, default_task
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

  static update() { 
    let body = {
      "data": {
        "approval_status": "approved",         // approved, rejected, changes_requested, pending
        "assignee": "me",
        "assignee_status": "upcoming",         // today, later, new, inbox, upcoming
        "due_on": "2021-08-06",
        "html_notes": "<body>Add support for the maintenance of <em>Projects and Tasks</em> in Asana</body>",
        "name": "Implement API to support maintenance for Project & Task",
        "notes": "Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are Projects and Tasks",
        "resource_subtype": "default_task",    // milestone, approval, section, default_task
      }
    };
  let options = {
    "method" : "PUT",
    "payload": JSON.stringify(body),
    "headers": {
      "Content-Type": "application/json",
      "Accept": "application/json",
      "Authorization": "Bearer " + ACCESS_TOKEN
      }
   }; 
    let url = Asana.getTaskUrl("1200720964791004");         // task_gid to update the task
    let response = UrlFetchApp.fetch(url,options);
    trace(`Task.update --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }  

  static destroy() {
    let options = {
      "method" : "DELETE",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };      
    let url = Asana.getTaskUrl("1234567890");         // Use desired {task_gid} to delete task
    UrlFetchApp.fetch(url,options);
  }
}


//=========================================================================================================
// Wrapper for Subtask - https://developers.asana.com/docs/get-subtasks-from-a-task
// Class Subtask

class Subtask {

  static create() {
    let body = {
      "data": {
        "approval_status": "pending",
        "assignee": "me",
        "assignee_status": "upcoming",
        "completed": false,
        "due_on": "2021-08-07",
        "html_notes": "<body>Integration of<strong>client sheets</strong> with Asana.</body>",
        "name": "Integration of client sheets with Asana",
        "notes": "Subtasks created in the project.",
        "resource_subtype": "default_task",
        "workspace": WORKSPACE_ID             // Hour Productions Testing
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
    let url = Asana.getSubtaskUrl("1200723005936586");        // {task_gid} to add subtask under that
    let response = UrlFetchApp.fetch(url,options);
    trace(`Subtask.create --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
  }

}














