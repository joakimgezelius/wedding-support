ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";  // Personal access token
WORKSPACE_ID = "1200711902496585";                                     // Hour Productions Testing
//ASSIGNEE     = User.active.email;                                          

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

  static getProjectName() {                     // Returns active spreadsheet name for project
    return Spreadsheet.active.name;
  }

  static getProjectDueDate() {
    let dueDate = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('WeddingDate').getValue();
    return Utilities.formatDate(dueDate, "GMT+1", "YYYY-MM-dd");
  }

  static getProjectGid() {                                      // Gets all projects details under the workspace  
    let options = {
    "method" : "GET",
    "headers": {
      "Accept": "application/json",
      "Authorization": "Bearer " + ACCESS_TOKEN
    }
    };
    let url = Asana.getProjectUrl("?workspace=1200711902496585");
    //let response = UrlFetchApp.fetch(url,options);
    let response = {"data":[{"gid":"1200711906671987","name":"Test API","resource_type":"project"},{"gid":"1200756401999433","name":"2021 Wedding Process ASANA upload (work in progress)","resource_type":"project"}]}
    //trace(`response = ${response}`);
    //let data = JSON.parse(response.getContentText());
    //trace(`data = ${data}`);                            // returns data = [object Object]

    let searchVal = Asana.getProjectName();
    trace(`Project Name = ${searchVal}`);
    for(let i = 0; i < response.data.length; i++) {
      if(response.data[i].name == searchVal)
      {
        return response.data[i].gid;
      }
      else {
        Dialog.notify("Try again!","Loop fails");
      }
    }

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
       "color": "dark-red",           // Color for project icon
       "current_status": { 
          "author": {
            "name": "Furkan Shaikh"
          },
          "color": "green",           // Status : Green - On Track, Red - Off Track, Blue - On Hold, Yellow - At Risk 
          "created_by": {
            "name": "Furkan Shaikh"
          },
          "html_text": "<body>Project created using the active spreadsheet name.</body>",  // Description
          "text": "The project is moving forward according to plan...",
          "title": "Project created using the active spreadsheet name"                     // Status title
        },
       //"start_on" : "",                                                                  // Project Start Date (Premium Only)
       "due_on": Asana.getProjectDueDate(),                                                // Due Date for Project
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": Asana.getProjectName(),                                                     // name for the project
       "notes": "These are things we need to purchase.",
       "owner": "joakim@gezelius.org",                                                     // owner of the project
       //"team": "",                                                                       // team_gid 
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
    let url = Asana.getProjectUrl("?workspace=1200711902496585");   // Hour Productions Testing Workspace_ID where project sits
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
          "color": "green",
          "created_by": {
            "name": "Furkan Shaikh"
          },
          "html_text": "<body>The project <strong>Status</strong> is updated...</body>",
          "text": "The project status is updated",
          "title": "Status Update - Aug 05"
        },        
       //"start_on" : "",                                                                  // Project Start Date (Premium Only)
       "due_on": Asana.getProjectDueDate(),
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": Asana.getProjectName(),
       "notes": "These are things we need to purchase.",       
       "owner": "joakim@gezelius.org",                                                     // owner of the project
       //"team": "",                                                                       // team_gid 
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
    let url = Asana.getProjectUrl("1200714734875880");                                     // set the project_gid param
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
    let url = Asana.getProjectUrl("");         // Use {project_gid} to delete desired project
    UrlFetchApp.fetch(url,options);
  }

}


//=========================================================================================================
// Wrapper for Task - https://developers.asana.com/docs/tasks
// Class Task

class Task {

static create() {
  let taskList = Range.getByName("AsanaTaskList","To Asana API");

  let newTask = {
    "data": {
      "approval_status": "pending",         // approved, rejected, changes_requested, pending
      "assignee": "me",
      "assignee_status": "upcoming",        // today, later, new, inbox, upcoming
      "completed": false,
      "due_on": "2021-08-09",               // due date
       //"start_on": "",                    // start date (Premium)
      "html_notes": "<body>Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are <em>Projects and Tasks</em></body>",                         // description
      "name": "Shopping in XYZ for ABC",    // task title
      "notes": "Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are Projects and Tasks",
      "projects": [
      //Asana.getProjectGid()               // project_gid for creating task under it
      ],
      "resource_subtype": "default_task",   //milestone, approval, section, default_task
    }
  };
  let options = {
    "method" : "POST",
    "payload": JSON.stringify(newTask),
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














