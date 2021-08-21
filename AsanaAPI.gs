ACCESS_TOKEN = "1/1200711887518296:7fa17fbab2b58d6deb3d3d4c0da0e39a";  // Personal access token (furkan)
WORKSPACE_ID = "1200711902496585";                                     // Hour Productions Testing
//ASSIGNEE     = User.active.email;                                          

class Asana { 

  static get asanaTaskListRange() { 
    return Range.getByName("AsanaTaskList", "To Asana API").loadColumnNames(); 
  }  

  static getProjectUrl(method) {
    return `${Asana.projectUrl}/${method}&opt_pretty`;      
  }

  static getTaskUrl(method) {
    return `${Asana.taskUrl}/${method}&opt_pretty`;      
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

  static getProjectGid() {                      // Gets all projects details under the workspace & returns project_gid
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
    trace(projects);
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
// Wrapper for Subtask - https://developers.asana.com/docs/sections
// Class Section

class Section {
  
}

function onCreateTask() {
  trace("onCreateTask");
  let taskList = new TaskList();
  let taskCreator = new TaskCreator();
  taskList.apply(taskCreator);
}


//=========================================================================================================
// Wrapper for Task - https://developers.asana.com/docs/tasks
// Class Task

class Task {

create(row) {
  let projectGid = Asana.getProjectGid();   // returns current project_gid for creating task under it
  let newTask = {
    "data": {
      "approval_status": "pending",         // approved, rejected, changes_requested, pending
      "assignee": "me",
      //"assignee_section": {
      //  "name" : "Onboarding"
      //},
      "assignee_status": "upcoming",        // today, later, new, inbox, upcoming
      "completed": false,
      "due_on": "2021-08-25",               // due date
       //"start_on": "",                    // start date (Premium)
      "html_notes": "<body>Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are <em>Projects and Tasks</em></body>",                         // description
      "name": "Client receipt  1st Deposit. Add reminder for 2nd deposit",    // task title
      "notes": "Work towards parameterisation of the API wrapper - The two main entities we will deal with in Asana are Projects and Tasks",
      "projects": [
        projectGid               
      ],
      "resource_subtype": "default_task",   // milestone, approval, section, default_task
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
    //let data = JSON.parse(response.getContentText());
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

//=============================================================================================
// Class TaskCreator
//

class TaskCreator {
  constructor(forced) {
    this.forced = forced;
    trace("NEW " + this.trace);
  }
  
  onEnd() {
    trace("TaskCreator.onEnd");
  }

  
  createTask(row) {
    trace("TaskCreator.onTitle");
    this.itemNo = 0;
    ++this.sectionNo;
    this.sectionTitleRange = row.range;
    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateSectionNo();
    }
  }

  onRow(row) {
    ++this.itemNo;
    trace("EventDetailsUpdater.onRow " + row.itemNo);
    let a1_currency = row.getA1Notation("Currency");
    let a1_quantity = row.getA1Notation("Quantity");
    let a1_nativeUnitCost = row.getA1Notation("NativeUnitCost");
    let a1_vat = row.getA1Notation("VAT");
    let a1_nativeUnitCostWithVAT = row.getA1Notation("NativeUnitCostWithVAT");
    let a1_unitCost = row.getA1Notation("UnitCost");
    let a1_commissionPercentage = row.getA1Notation("CommissionPercentage");
    let a1_markup = row.getA1Notation("Markup");
    let a1_unitPrice = row.getA1Notation("UnitPrice");
    let a1_startTime = row.getA1Notation("Time");
    let a1_endTime = row.getA1Notation("EndTime");

    if (row.itemNo === "" || this.forced) { // Only set item number if empty (or forced)
      row.itemNo = this.generateItemNo();
    }
    this.setNativeUnitCost(row);
    row.nativeUnitCostWithVAT = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCost}="", ${a1_nativeUnitCost}=0), "", ${a1_nativeUnitCost}*(1+${a1_vat}))`;
    row.getCell("NativeUnitCostWithVAT").setNumberFormat(row.currencyFormat);
    row.unitCost = `=IF(OR(${a1_currency}="", ${a1_nativeUnitCostWithVAT}="", ${a1_nativeUnitCostWithVAT}=0), "", IF(${a1_currency}="GBP", ${a1_nativeUnitCostWithVAT}, ${a1_nativeUnitCostWithVAT} / EURGBP))`;
    row.totalGrossCost = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitCost}="", ${a1_unitCost}=0), "", ${a1_quantity} * ${a1_unitCost})`;
    row.totalNettCost = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitCost}="", ${a1_unitCost}=0), "", ${a1_quantity} * ${a1_unitCost} * (1-${a1_commissionPercentage}))`;
    row.unitPrice = `=IF(OR(${a1_unitCost}="", ${a1_unitCost}=0), "", ${a1_unitCost} * ( 1 + ${a1_markup}))`;
    row.totalPrice = `=IF(OR(${a1_quantity}="", ${a1_quantity}=0, ${a1_unitPrice}="", ${a1_unitPrice}=0), "", ${a1_quantity} * ${a1_unitPrice})`;
//  row.commission = `=IF(OR(${a1_commissionPercentage}="", ${a1_commissionPercentage}=0), "", ${a1_quantity} * ${a1_unitCost} * ${a1_commissionPercentage})`;
    if (row.isStaffTicked && (row.quantity === "")) { // Calculate quantity if staff and time fields are filled 
      row.quantity = `=IF(OR(${a1_startTime}="", ${a1_endTime}=""), "", ((hour(${a1_endTime})*60+minute(${a1_endTime}))-(hour(${a1_startTime})*60+minute(${a1_startTime})))/60)`;
    }
    if (row.isInStock && this.forced) { // Set mark-up and commission for in-stock items
      trace("- Set In Stock commision & mark-up on " + this.itemNo);
      row.commissionPercentage = 0.5;
      row.markup = 0;
    }
  }

  generateSectionNo() {
    return Utilities.formatString("#%02d", this.sectionNo);
  }

  generateItemNo() {
    if (this.itemNo == 0) 
      return this.generateSectionNo();
    else
      return Utilities.formatString("%s-%02d", this.generateSectionNo(), this.itemNo);
  }

  setNativeUnitCost(row) {
    let budgetUnitCostCell = row.getCell("BudgetUnitCost");
    let a1_budgetUnitCost = budgetUnitCostCell.getA1Notation();
    let nativeUnitCost = String(row.nativeUnitCost);
    let nativeUnitCostCell = row.getCell("NativeUnitCost");
    let currencyFormat = row.currencyFormat;
    //trace(`Cell: ${nativeUnitCostCell.getA1Notation()} ${currencyFormat}`);
    budgetUnitCostCell.setNumberFormat(currencyFormat);
    nativeUnitCostCell.setNumberFormat(currencyFormat);
    if (nativeUnitCost === "") { // Set native unit cost equal to budget unit cost if not set 
      //trace(`Set nativeUnitCost to =${budgetUnitCostA1}`);
      row.nativeUnitCost = `=${a1_budgetUnitCost}`;
      nativeUnitCostCell.setFontColor("#ccccff"); // Make text grey
    } else {
      //trace(`setNativeUnitCost: nativeUnitCost="${nativeUnitCost} - make it black`);
      nativeUnitCostCell.setFontColor("#000000"); // Black
    }
  }

  get trace() {
    return `{EventDetailsUpdater forced=${this.forced}}`;
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
















