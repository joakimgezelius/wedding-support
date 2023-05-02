// Sheets
const ParamsSheetName = "Params";
const CoordinationSheetName = "Coordinator";

// Named Ranges
const UpcomingProjects = "UpcomingProjects";
const ProjectFoldersRoot = "ProjectFoldersRoot";
const PaymentsFoldersRoot = "PaymentsFoldersRoot";
const SelectedTemplateProjectFolder = "SelectedTemplateProjectFolder";
const SelectedTemplateProjectSheet = "SelectedTemplateProjectSheet";
const SelectedProjectName = "SelectedProjectName";
const SelectedProjectType = "SelectedProjectType";
const SelectedTemplateName = "SelectedTemplateName";

// Go through raw list of projects, 
//  - find those that are work-in-progress
//
function onUpdateProjects() {
  trace("onUpdateProjects");
  let projects = new Projects;
  projects.update(projectsNoReply);
}

// Open project sheet for the selected project, 
// - option to create a new sheet if no such sheet exists
//
function onOpenProjectSheet() {
  trace("onOpenProjectSheet");
  let projects = new Projects;
  projects.selected.openProjectSheet();
}

function onCreateNewProjectSheet() {
  trace("onCreateNewProjectSheet");
  let projects = new Projects;
  projects.selected.createNewProjectSheet();
}

function onDraftSelectedEmail() {
  trace("onDraftSelectedEmail");
  let projects = new Projects;
  projects.selected.draftSelectedEmail();
}

function onPrepareProjectStructure() {
  trace("onPrepareProjectStructure");
  let projects = new Projects;
  let templateFolderLink = Spreadsheet.getCellValueLinkUrl(SelectedTemplateProjectFolder);
  let templateProjectSheetLink = Spreadsheet.getCellValueLinkUrl(SelectedTemplateProjectSheet);
  projects.selected.prepareProjectStructure(templateFolderLink, templateProjectSheetLink);
}

function onDeleteProjectDocumentStructure() {
  trace("onDeleteProjectDocumentStructure");
  let projects = new Projects;
  projects.selected.deleteProjectDocumentStructure();
}


// This event handler triggers when a new project row is selected in the master sheet, the primary event trigger 
// onSelectionChange is located in the master sheet itself, and flters out unrelated selections before calling 
// this handler. The handler identifies the selected project, and fills in the name and type on the master sheet
// for visual feedback.
//
function onNewProjectRowSelection() {
  trace("> onNewProjectRowSelection");
  let projects = new Projects;
  let selectedProjectName = "-";
  let selectedProjectType = "-";
  try {
    var selectedProject = projects.selected;
    trace(`selectedProject: ${selectedProject.trace}`);
    selectedProjectName = selectedProject.name;
    selectedProjectType = selectedProject.type;
  } catch {
    trace("onNewProjectRowSelection - not a valid project");
  }
  Range.getByName(SelectedProjectName).value = selectedProjectName;
  Range.getByName(SelectedProjectType).value = selectedProjectType;
}

//========================================================================================================

class Projects {

  constructor(rangeName = null) {
    rangeName = (rangeName === null) ? UpcomingProjects : rangeName;
    this.range = Range.getByName(rangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  get selected() {
    let selection = this.range.findSelectedRow();
    if (selection === null) {
      Error.fatal("Please select a valid project.");
    }
    // Range is now positioned to selection
    let selectedProject = new Project(this.range);
    if (!selectedProject.isValid) {
      Error.fatal("Please select a valid project.");
    }
    return selectedProject;
  }
  
  update(target) {
    trace(`${this.trace}.update`);
    for (var rowOffset = 0; rowOffset < this.range.height; rowOffset++) {
      let project = new Project(); // Fix this!!
      if (!project.isValid) {
        break;
      }
      target.append(project);
    }
  }
  
  // Append an project to a list 
  //
  append(project) {
    trace(`${this.trace}.append ${project.trace}`);    
    this.range.findFirstTrailingEmptyRow();
    //  trace(`Create target project based on ${this.range.trace}, ${this.range.currentRowOffset}`);    
    let targetProject = new Project(); // Fix this!!
    project.copyTo(targetProject);
  }

  get trace() { return `{Projects ${this.range.trace}}`; }
  
} // Projects


//================================================================================================


class Project extends RangeRow {
  
  constructor(range) {
    super(range);
    this._name = this.name;
    this._rowOffset = range.currentRowOffset;
    this._isValid = this.validate();
    trace("NEW " + this.trace);
  }
  
  validate() {
    let isValid = (
      (this.name !== "") &&
      (this.date != "")
    );
    return isValid;
  }

  createNewProjectSheet() {
    trace(`createNewProjectSheet for ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid project.");
    };
    let projectSheetName = `${this.name}`;
    //  let targetFolder = Folder.getById(projectsFolderId);
    let weddingProjectTemplateFile = File.getById(weddingProjectTemplateSpreadsheetId);
    let projectSpreadsheetFile = weddingProjectTemplateFile.copyTo(targetFolder, projectSheetName);
    this.projectSheet = Spreadsheet.openById(projectSpreadsheetFile.id);
    this.sheetLink = this.projectSheet.url; // Note, don't set labelled hyperlink, as it creates problems downstream
    Browser.newTab(this.projectSheet.url);
  }

  prepareProjectStructure(templateFolderLink, templateProjectSheetLink) {
    trace(`> prepareProjectStructure for ${this.trace}, templateFolder=${templateFolderLink}, templateProjectSheet=${templateProjectSheetLink}`);
    let templateFolder = Folder.getByUrl(templateFolderLink);
    let destinationFolderURL = Spreadsheet.getCellValueLinkUrl(ProjectFoldersRoot);
    let destinationFolder = Folder.getByUrl(destinationFolderURL);
    trace(`destinationFolder: ${destinationFolder.trace}`);
    let paymentsFoldersRootURL = Spreadsheet.getCellValueLinkUrl(PaymentsFoldersRoot);
    let paymentsFoldersRoot = Folder.getByUrl(paymentsFoldersRootURL);
    let templateName = Spreadsheet.getCellValue(SelectedTemplateName);
    if (destinationFolder.folderExists(this.fileName)) {
      Dialog.notify("Project folder already exists!", "Please check the Weddings & Events folder for more details.");      
    } 
    else {
      if (Dialog.confirm("Preparing Project Document Structure", `Preparing a new project document structure for ${this.name}, using template ${templateName}. Are you sure?`)) {
        Dialog.toast("Preparing document structure, this may take a minute or two...");
        templateFolder.copyTo(destinationFolder, this.fileName);                    // Copies source folder contents to target folder
        let paymentsFolderName =  this.fileName + " - Payments";
        paymentsFoldersRoot.createFolder(paymentsFolderName);

        let templateSheetFile = File.getByUrl(templateProjectSheetLink);

        let projectFolder = destinationFolder.getSubfolder(this.fileName);         // Gets newly created project folder by name
        let projectFolderLink = projectFolder.url;                                 // Gets the URL of newly created project folder 
        this.folderLink = projectFolderLink;                                       // Sets the folder link in the cell in FolderLink Column

        let paymentsFolder = paymentsFoldersRoot.getSubfolder(paymentsFolderName); // Gets newly created payment folder by name
        let paymentsFolderLink = paymentsFolder.url;                               // Returns URL to Payment Folder
        this.paymentsLink = paymentsFolderLink;                                    // Sets the URL in the Master sheet

        //let targetFolderName = "Office Use";                                     // Folder name to look for copying the template file in it
        //let subFolder = projectFolder.getSubfolder(targetFolderName);
        //let subFolderId = subFolder.id;                                          // Gets the id of found subfolder "Office Use"

        let targetFolder = projectFolder;                                          // Gets the folder by id to copy the template file in it
        if (targetFolder) {
            projectFolder.createShortcut(paymentsFolder, "Payments");              // Creates shortcut to Payment Folder in Project Folder
            templateSheetFile.copyTo(targetFolder,this.fileName);       
            let newProjectSheet = targetFolder.getFile(this.fileName);             // Gets the newly copied file with given name
            let newProjectSheetId = newProjectSheet.id;                            // Returns the id of found file
            let projectTemplate = Spreadsheet.openById(newProjectSheetId);
            projectTemplate.setActive();                                           // Sets the active spreadsheet Confirmed W & E to new project sheet
            let newProjectSheetLink = newProjectSheet.url;                         // Gets the URL of newly copied template file
            this.sheetLink= newProjectSheetLink;                                   // Sets the project sheet link to the cell in SheetLink Column
            Dialog.notify("Project Document Structure Created!", 'Please check columns "Project Sheet", "Project Folder" & "Project Payments Folder" for more details.');
            Browser.newTab(newProjectSheetLink);
        } 
        else {
            Dialog.notify("Folder not found!","Template folder does not exist! Couldn't make a copy of template sheet, please check source folder.");
        }
      }
    }
  }

  preparePaymentsFolder() {
    trace(`preparePaymentsFolder for ${this.trace}`);
  }

  copyTo(destination) {
    trace(`${this.trace}.copyTo ${destination.trace}`);
    const fields = ["Name", "EmailAddress", "Who"];
    this.copyFieldsTo(destination, fields);
  }

  openProjectSheet() {
    trace(`openProjectSheet of ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid project.");
    }
    return;
    let sheetId = this.sheetId;
    if (sheetId === "") {
      if (Dialog.confirm("No project sheet found", `There is no project sheet recorded for prospect ${this.name}, create one now?`) == false) {
        return;
      }
      this.createNewProjectSheet();
    } else {
      this.projectSheet = null;
      Browser.newTab(this.projectSheet.url);
    }
  }

  deleteProjectDocumentStructure() {
    if (Dialog.confirm("Delete Project Document Structure", `Remove the document structure created for project ${this.name}. Files will be moved to the bin, and permanently deleted after 30 days. Are you sure?`)) {
      if (Dialog.confirm("Delete Project Document Structure", `Files for project ${this.name} will be moved to the bin, and permanently deleted after 30 days. Are you REALLY sure?`)) {
        Dialog.toast("Deleting document structure, this may take a few seconds...");

        //Dialog.toast(`Document structure for project ${this.name} has been moved to the bin.`);
      }
    };
  }

  draftSelectedEmail() {
    trace(`draftSelectedEmail to ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid project.");
    }
    if (Dialog.confirm("Draft Email", `Draft email to ${this.name}, are you sure?`) == false) {
      return;
    }    
  }  
  
  get name()              { return this.get("Name", "string"); }
  get type()              { return this.get("Type", "string"); }
  get date()              { return this.get("EventDate"); }
  // Time-Zone Changes in the Summer Time begins and ends at 1:00 a.m ( Universal Time (GMT))
  get fileName()          { return `${Utilities.formatDate(this.date, "GMT+2", "yyyy-MM-dd")} ${this.name}`; }  
  get sheetLink()         { return this.get("SheetLink", "string"); }  
  get folderLink()        { return this.get("FolderLink", "string"); }
  get paymentsLink()      { return this.get("PaymentsLink", "string"); }
  get isValid()           { return this._isValid; }
  get rowOffset()         { return this._rowOffset; }
  set sheetLink(value)    { this.set("SheetLink", value); }
  set folderLink(value)   { this.set("FolderLink", value); }
  set paymentsLink(value) { this.set("PaymentsLink", value); }

  get trace() { return `{Project #${this.rowOffset} ${this._name} ${this.isValid ? "(valid)" : "(invalid)"}}`; }
  
} // Project

//================================================================================================
               
class Prospects {
  
  constructor() {
    this.range = Range.getByName(ProjectsWaitingResponseRangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  get trace() { return `{Prospects ${this.range.trace}}`; }

}


//================================================================================================

class Prospect {
  
  constructor() {
    this.range = Range.getByName(ProjectsWaitingResponseRangeName).loadColumnNames();
    trace("NEW " + this.trace);
  }

  get trace() { return `{Prospects ${this.range.trace}}`; }

}
