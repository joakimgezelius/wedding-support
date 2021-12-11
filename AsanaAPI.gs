ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";  // Personal access token (furkan)
WORKSPACE_ID = "1200711902496585";                                     // Hour Productions Testing                               

function onCreateTask() {
  trace("onCreateTask");
  if(Asana.checkAsanaProjectNames()) {
    Dialog.notify("Uploading the tasks to Asana...","Please wait this may take few minutes to upload the tasks in Asana workspace!");
    let taskList = new TaskList();
    let taskCreator = new TaskCreator();
    taskList.apply(taskCreator);
    //onCreateSubTask();
  }
  else {
    Dialog.notify("Project not found!", PROJECT_NAME + " project does not exist in the Asana, first make the project and try again!");
  }
}

function onCreateSubTask() {
  let taskList  = new TaskList();
  let subTaskCreator = new SubTaskCreator();
  taskList.apply(subTaskCreator);
}

function onUpdateTask() {
  trace("onUpdateTask");
  let tasks = Asana.getTaskGid();
  if (tasks !== null) {
    Dialog.notify("Updating the Tasks...","Please wait it may take few minutes to update the tasks in Asana for " + PROJECT_NAME);
    let taskList = new TaskList();
    let taskUpdater = new TaskUpdater();
    taskList.apply(taskUpdater);
    Dialog.notify("Tasks are Updated!","Please check tasks are updated in Asana for " + PROJECT_NAME);
  }
  else {
    Dialog.notify("No Tasks Found!", PROJECT_NAME + " project does not contain any task in the Asana yet!");
  }
}

function onDestroyTask() {
  trace("onDestroyTask");
  let tasks = Asana.getTaskGid();
  if (tasks !== null) {
    if (Dialog.confirm("To Delete Tasks from Asana Project - Confirmation Required", "Are you sure you want to delete?") == true) {
      let taskList = new TaskList();
      let taskDestroyer = new TaskDestroyer();
      taskList.apply(taskDestroyer);  
    }
  }
  else {
    Dialog.notify("No Tasks Found!", PROJECT_NAME + " project does not contain any task in the Asana yet!");
  }
}

function onCreateProject () {
  trace(`onCreateProject`);
  if(Asana.checkAsanaProjectNames()) {
    let project_gid = Asana.getProjectGid();
    Dialog.notify("Project Already Exist!","Couldn't make another "+ PROJECT_NAME +" project, Please click Ok for more details!");
    Browser.newTab("https://app.asana.com/0/"+project_gid+"/list");
  }
  else {
    Dialog.notify("Creating the Project...","Please wait creating the "+ PROJECT_NAME +" project in Asana workspace!");
    Project.create();
    Dialog.notify("Adding the Sections...","Please wait it will take few seconds to add sections in the Project");
    let phaseList = new WeddingPhaseList();
    let sectionCreator = new SectionCreator();
    phaseList.apply(sectionCreator);
  }
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
    return Range.getByName("AsanaTaskList", "To Asana").loadColumnNames(); 
  }  

  static get asanaWeddingPhaseRange() {
    return Range.getByName("WeddingPhases", "Wedding Params").loadColumnNames();
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
    return "n/a" ; // Spreadsheet.active.name;  NOTE: the way this is coded it causes a static object to be created at load time, let's try to avoid that
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
    let project_gid, i = 0;
    let pl = projects.length;
    for(i; i < pl; ++i) {       // Loop through the array and find the match
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
    let asanaProject_name, i = 0;
    let pl = projects.length;
    for(i; i < pl; ++i) {                       // Loop through the array and find the match
      if( projects[i].name == searchVal)
      {
        asanaProject_name = projects[i].name;
      }
    }
    return asanaProject_name;
  }

  static getProjectSectionGid(taskSection) {
    let options = {
      "method" : "GET",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let projectGid = Asana.getProjectGid();
    let url = Asana.getProjectUrl(projectGid+"/sections");
    let response = UrlFetchApp.fetch(url,options);
    let data = JSON.parse(response.getContentText());
    let sections = Array.from(data['data']);
    //console.log(sections);
    let sectionName = taskSection;
    let sectionGid, i = 0;
    let sl = sections.length;
    for(i; i < sl; ++i) {
      if( sections[i].name == sectionName)
      {
        sectionGid = sections[i].gid;
      }
    }
    return sectionGid;
  }

  static getTaskGid(taskName) {
    let options = {
      "method" : "GET",
      "headers": {
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let projectGid = Asana.getProjectGid();
    let url = Asana.getProjectUrl(projectGid+"/tasks");
    let response = UrlFetchApp.fetch(url,options);
    let data = JSON.parse(response.getContentText());
    let tasks = Array.from(data['data']);
    //console.log(tasks);
    let searchVal = taskName;
    let taskGid, i = 0;
    let tl = tasks.length;
    for(i; i < tl; ++i){
      if( tasks[i].name == searchVal) {
        taskGid = tasks[i].gid;
      }
    }
    return taskGid;
  }

  static setTaskDependents() {  // Asana premium only feature
    let body = {
      "data" : {
        "dependents": ["1201136448252155"]
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
    let url = Asana.getTaskUrl("1201136539911585/addDependents");
    let response = UrlFetchApp.fetch(url,options);
    trace(`${response}`);
  }
}

PROJECT_NAME = Asana.getProjectName();                           // Returns active project name
Asana.projectUrl = "https://app.asana.com/api/1.0/projects";     // For basic project operations
Asana.taskUrl    = "https://app.asana.com/api/1.0/tasks";        // For creating the task in Asana

//=========================================================================================================
// Wrapper for project - https://developers.asana.com/docs/projects
// Class Project

class Project {
  
  static create() {
    
    let dueDate = Asana.getProjectDueDate();
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
            "title": "Project created using the active spreadsheet name"                    // Status title
          },
        //"start_on" : "",                                                                  // Project Start Date (Premium Only)
        "due_on": dueDate,                                                                  // Due Date for Project
        "html_notes": "<body>These are things we need to purchase.</body>",
        "name": PROJECT_NAME,                                                               // name for the project
        "notes": "These are things we need to purchase.",
        "owner": "me",                                                                      // owner of the project
        //"team": "monica@hour.events",                                                     // members to the project
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
    Dialog.notify("Project is Created!","Please check Asana workspace for more details of project " +  PROJECT_NAME);   
  }

  static update() {
    let project_gid = Asana.getProjectGid();
    let dueDate = Asana.getProjectDueDate()
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
          "title": "Status Update - Aug 05"
        },        
       //"start_on" : "",                                                                  // Project Start Date (Premium Only)
       "due_on": dueDate,
       "html_notes": "<body>These are things we need to purchase.</body>",
       "name": PROJECT_NAME,
       "notes": "These are things we need to purchase.",       
       "owner": "me",                                                     // owner of the project
       //"team": "",                                                                      
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

  static create() {
    let projectGid = Asana.getProjectGid();
    let body = {
      "data" : {
        "approval_status": "pending",
        "assignee": "me",
        "assignee_section": "1201147072509061",
        "assignee_status": "upcoming",
        "completed": false,
        "due_on": "2021-10-10",
        "html_notes": "<body>Testing the subtask addition</body>",
        "name": "Testing the subtask addition",
        "notes": "Mittens really likes the stuff from Humboldt (notes).",
        //"parent": "1201146451517371",
        "projects": [
          projectGid
        ],
        "resource_subtype": "default_task"
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
    let url = Asana.getTaskUrl("1201153177857031/subtasks");
    UrlFetchApp.fetch(url,options);
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
// Class TaskCreator, TaskUpdater, TaskDestroyer
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
    let projectGid = Asana.getProjectGid();
    let taskSection = row.section;
    let taskSectionGid = Asana.getProjectSectionGid(taskSection);  
    let newTask = {
      "data": {
        "approval_status": "pending",       // approved, rejected, changes_requested, pending
        "assignee": "me",                   // add monica and team to workspace
        "assignee_section": taskSectionGid,
        "assignee_status": "upcoming",      // today, later, new, inbox, upcoming
        "completed": false,
        "due_on": row.dueDate,              // due date
        //"start_on": "",                   // start date (Premium)
        "html_notes": "<body>" + row.description + "</body>",      // description
        "name": row.name,                   // task name
        "notes": row.notes,                 // notes
        "projects": [
          projectGid              
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
  }

  get trace() {
    return `{TaskCreator forced=${this.forced}}`;
  }
}

/*class SubTaskCreator {

  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }
  
  onEnd() {
    trace("SubTaskCreator.onEnd - no-op");
  }

  onRow(row) {
    let projectGid = Asana.getProjectGid();
    let taskName = row.name;
    let taskGid  = Asana.getTaskGid(taskName); 
    let taskSection = row.section;
    let taskSectionGid = Asana.getProjectSectionGid(taskSection);   
    let parentTask = {
      "data" : {
        "approval_status": "pending",
        "assignee": "me",
        "assignee_section": taskSectionGid,
        "assignee_status": "upcoming",
        "completed": false,
        "due_on": row.dueDate,
        "html_notes": "<body>"+ row.description +"</body>",
        "name": row.subtaskOf,
        "notes": row.notes,
        //"parent": taskGid,
        "projects": [
          projectGid
        ]
      }
    };
    let options = {
    "method" : "POST",
    "payload": JSON.stringify(parentTask),
    "headers": {
      "Content-Type": "application/json",
      "Accept": "application/json",
      "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let url = Asana.getTaskUrl(taskGid+"/subtasks");
    UrlFetchApp.fetch(url,options); 
  }

  get trace() {
    return `{SubTaskCreator forced=${this.forced}}`;
  }
} */

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
    let taskName = row.name;
    let taskGid  = Asana.getTaskGid(taskName); 
    let taskSection = row.section;
    let taskSectionGid = Asana.getProjectSectionGid(taskSection);
   
    let newTask = {
      "data": {
        "approval_status": "pending",       // approved, rejected, changes_requested, pending
        "assignee": "me",                   // add monica and team to workspace
        "assignee_section": taskSectionGid,
        "assignee_status": "upcoming",      // today, later, new, inbox, upcoming
        "due_on": row.dueDate,              // due date
        //"start_on": "",                   // start date (Premium)
        "html_notes": "<body>" + row.description + "</body>",      // description
        "name": row.name,                   // task name
        "notes": row.notes,                 // notes
        "resource_subtype": "default_task",       // Premium feature - milestone, approval, section, default_task
      }
    };
    let options = {
      "method" : "PUT",
      "payload": JSON.stringify(newTask),
      "headers": {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let url = Asana.getTaskUrl(taskGid);
    UrlFetchApp.fetch(url,options);
  }

  get trace() {
    return `{TaskUpdater forced=${this.forced}}`;
  }
}

//---------------------------------------------------------------------------------------------------------

class TaskDestroyer {

  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }
  
  onEnd() {
    trace("TaskDestroyer.onEnd - no-op");
  }

  onRow(row) {
    let taskName = row.name;
    let taskGid  = Asana.getTaskGid(taskName);
    let options = {
      "method" : "DELETE",
      "headers": {
        "Accept": "application/json",
        "muteHttpExceptions": true,
        "Authorization": "Bearer " + ACCESS_TOKEN
      }
    };
    let url = Asana.getTaskUrl(taskGid);
    UrlFetchApp.fetch(url,options);
  }

  get trace() {
    return `{TaskDestroyer forced=${this.forced}}`;
  }

}

//=========================================================================================================
// Wrapper for Section - https://developers.asana.com/docs/sections
// Class SectionCreator

class SectionCreator {
  
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }
  
  onEnd() {
    trace("SectionCreator.onEnd - no-op");
  }

  onRow(row) {
    let projectGid = Asana.getProjectGid();
    let sectionName= row.weddingPhase;
    let body = {
      "data": {
        "name": sectionName,
        "resource_type": "section"
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
    let url = Asana.getProjectUrl(projectGid+"/sections");
    UrlFetchApp.fetch(url,options);
  }

  get trace() {
    return `{SectionCreator forced=${this.forced}}`;
  }
}