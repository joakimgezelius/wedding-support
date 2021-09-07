ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";  // Personal access token (furkan)
WORKSPACE_ID = "1200711902496585";                                     // Hour Productions Testing
//ASSIGNEE   = User.active.email;                                        

function onCreateTask() {
  trace("onCreateTask");
  if(Asana.checkAsanaProjectNames()) {
    Dialog.notify("Uploading the tasks to Asana...","Please wait this may take few minutes to upload the tasks in Asana workspace!");
    let taskList = new TaskList();
    let taskCreator = new TaskCreator();
    taskList.apply(taskCreator);
  }
  else {
    Dialog.notify("Project not found!", PROJECT_NAME + " project does not exist in the Asana, first make the project and try again!");
  }
}

function onUpdateTask() {
  trace("onUpdateTask");
  let taskList = new TaskList();
  let taskUpdater = new TaskUpdater();
  taskList.apply(taskUpdater);
}

function onCreateProject () {
  trace(`onCreateProject`);
  Project.create();
}

function onUpdateProject () {
  trace(`onUpdateProject`);  
  if(Asana.checkAsanaProjectNames()) {
    Project.update();
  }
  else {
    Dialog.notify("Project not found!", PROJECT_NAME + " project does not exist in the Asana, first make the project and try again!");
  }
}

function onDestroyProject () {
  trace(`onDestroyProject`);
  if(Asana.checkAsanaProjectNames()) {
    if (Dialog.confirm("To Delete Asana Project - Confirmation Required", "Are you sure you want to delete "+ PROJECT_NAME +" project from Asana workspace?") == true) {
    Project.destroy();
    }
  }
  else {
    Dialog.notify("Project not found!", PROJECT_NAME + " project does not exist in the Asana, first make the project and try again!");
  }

}

//=========================================================================================================

class Asana { 

  static get asanaTaskListRange() { 
    return Range.getByName("AsanaTaskList", "To Asana API").loadColumnNames(); 
  }  

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

  static getProjectGid() {                      // Gets all projects details under the workspace & returns active project_gid
    let options = {
      "method" : "GET",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let url = Asana.getProjectUrl("?workspace=1200711902496585");
    let response = UrlFetchApp.fetch(url,options);  
    let data = JSON.parse(response.getContentText());   
    let projects = Array.from(data['data']);         // Creates a new shallow-copied Array instance from an array-like or iterable object.
    //trace(projects);
    let searchVal = Asana.getProjectName();           
    let project_gid;
    for(let i = 0; i < projects.length; i++) {       // Loop through the array and find the match
      if( projects[i].name == searchVal)
      {
        project_gid = projects[i].gid;
      }
    }
    return project_gid;
  }

  static checkAsanaProjectNames() {
    let options = {
      "method" : "GET",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let url = Asana.getProjectUrl("?workspace=1200711902496585");
    let response = UrlFetchApp.fetch(url,options);
    let data = JSON.parse(response.getContentText());
    let projects = Array.from(data['data']);
    //trace(projects);
    let searchVal = Asana.getProjectName(); 
    let asanaProject_name;
    for(let i = 0; i < projects.length; i++) {       // Loop through the array and find the match
      if( projects[i].name == searchVal)
      {
        asanaProject_name = projects[i].name;
      }
    }
    return asanaProject_name;
  }
}

Asana.projectUrl = "https://app.asana.com/api/1.0/projects";     // For basic project operations
Asana.taskUrl    = "https://app.asana.com/api/1.0/tasks";        // For creating the task in Asana

//PROJECT_GID  = Asana.getProjectGid();                          // Returns active project gid
PROJECT_NAME = Asana.getProjectName();                           // Returns active project name
PROJECT_DUE  = Asana.getProjectDueDate();                        // Returns active project due date


//=========================================================================================================
// Wrapper for project - https://developers.asana.com/docs/projects
// Class Project

class Project {
  
  static create() {
    let project_gid = Asana.getProjectGid();
    if(Asana.checkAsanaProjectNames()) {
      Dialog.notify("Project Already Exist!","Couldn't make another "+ PROJECT_NAME +" project, Please click Ok for more details!");
      Browser.newTab("https://app.asana.com/0/"+project_gid+"/list");
    }
    else {
      Dialog.notify("Creating the Project...","Please wait creating the "+ PROJECT_NAME +" project in Asana workspace!");
      let body = {
      "data": {
        "archived": false,
        "color": "dark-red",           // Color for project icon
        "current_status": { 
            "author": {
              "name": "Hour Productions"
            },
            "color": "green",           // Status : Green - On Track, Red - Off Track, Blue - On Hold, Yellow - At Risk 
            "created_by": {
              "name": "Hour Productions"
            },
            "html_text": "<body>Project created using the active spreadsheet name.</body>",  // Description
            "text": "The project is moving forward according to plan...",
            "title": "Project created using the active spreadsheet name"                     // Status title
          },
        //"start_on" : "",                                                                  // Project Start Date (Premium Only)
        "due_on": PROJECT_DUE,                                                              // Due Date for Project
        "html_notes": "<body>These are things we need to purchase.</body>",
        "name": PROJECT_NAME,                                                               // name for the project
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
      UrlFetchApp.fetch(url,options);
      //trace(`Project.create --> ${response.getContentText()}`);
      Dialog.notify("Project is created!","Please check Asana workspace for more details of project " +  PROJECT_NAME);
    }
  }

  static update() {
    let project_gid = Asana.getProjectGid();
    Dialog.notify("Updating the Project...","Please wait updating the "+ PROJECT_NAME +"!");
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
       "due_on": PROJECT_DUE,
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": PROJECT_NAME,
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
    let url = Asana.getProjectUrl(project_gid);                                     // set the project_gid param
    UrlFetchApp.fetch(url,options);
    //trace(`Project.update --> ${response.getContentText()}`);
    Dialog.notify("Project is updated!","Please check Asana workspace for more details of project " +  PROJECT_NAME);
  }

  static destroy() {
    let project_gid = Asana.getProjectGid();
    let options = {
      "method" : "DELETE",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };             
    let url = Asana.getProjectUrl(project_gid);         
    UrlFetchApp.fetch(url,options);
    Dialog.notify("Project Removed from Asana", PROJECT_NAME + " project is deleted from Asana workspace");
  }

}

//=========================================================================================================
// Wrapper for Task - https://developers.asana.com/docs/tasks
// Class Task

class Task {

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
// Class TaskCreator, TaskUpdater
//

class TaskCreator {
  
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }
  
  onEnd() {
    trace("TaskCreator.onEnd - no-op");
  }

  onRow(row) {
  let newTask = {
    "data": {
      "approval_status": "pending",       // approved, rejected, changes_requested, pending
      "assignee": "me",                   // add monica and team to workspace
      //"assignee_section": {
      //  "name" : "Onboarding"
      //},
      "assignee_status": "upcoming",      // today, later, new, inbox, upcoming
      "completed": false,
      "due_on": row.dueDate,              // due date
    //"start_on": "",                     // start date (Premium)
      "html_notes": "<body>" + row.description + "</body>",      // description
      "name": row.name,                   // task name
      "notes": row.notes,                 // notes
      "projects": [
        Asana.getProjectGid()               
      ],
      "resource_subtype": "default_task",       // Premium feature - milestone, approval, section, default_task
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
    UrlFetchApp.fetch(url,options);
    //let response = UrlFetchApp.fetch(url,options);
    //trace(`Task.create --> ${response.getContentText()}`);    // Exceeds execution time
  }

  get trace() {
    return `{TaskCreator forced=${this.forced}}`;
  }
}

//---------------------------------------------------------------------------------------------------------

class TaskUpdater {
  
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }
  
  onEnd() {
    trace("TaskUpdater.onEnd - no-op");
  }

  onRow(row) {
  let body = {
    "data": {
      "approval_status": "approved",         // approved, rejected, changes_requested, pending
      "assignee": "me",
      "assignee_status": "upcoming",         // today, later, new, inbox, upcoming
      "due_on": row.dueDate,
      "html_notes": "<body>"+ row.notes +"</body>",
      "name": row.name,
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
    UrlFetchApp.fetch(url,options);
  }

  get trace() {
    return `{TaskUpdater forced=${this.forced}}`;
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


//=========================================================================================================
// Wrapper for Subtask - https://developers.asana.com/docs/sections
// Class Section

class Section {
  
}
















