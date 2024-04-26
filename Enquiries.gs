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
const SelectedProjectIndex = "SelectedProjectIndex";
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

function onCreateProjectSheet() {
  trace("onCreateProjectSheet");
  //let projects = new Projects;
  //projects.selected.createProjectSheet();
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
//  let selection = this.range.findSelectedRow(); NOTE: we explicitly pick the project now, we don't rely on active cell selection as it is too unstable.
    const selection = Range.getByName(SelectedProjectIndex).value;
    this.range.currentRowOffset = selection;
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

/*
  createProjectSheet() {
    trace(`createProjectSheet for ${this.trace}`);
    if (!this.isValid) {
      Error.fatal("Please select a valid project.");
    };
    let projectSheetName = `${this.name}`;
    //  let projectFolder = Folder.getById(projectsFolderId);
    let weddingProjectTemplateFile = File.getById(weddingProjectTemplateSpreadsheetId);
    let projectSpreadsheetFile = weddingProjectTemplateFile.copyTo(projectFolder, projectSheetName);
    this.projectSheet = Spreadsheet.openById(projectSpreadsheetFile.id);
    this.sheetLink = this.projectSheet.url; // Note, don't set labelled hyperlink, as it creates problems downstream
    Browser.newTab(this.projectSheet.url);
  }
*/

  prepareProjectStructure(templateFolderLink, templateProjectSheetLink) {
    trace(`> prepareProjectStructure for ${this.trace}, templateFolder=${templateFolderLink}, templateProjectSheet=${templateProjectSheetLink}`);
    let templateFolder = Folder.getByUrl(templateFolderLink);
    let templateSheetFile = File.getByUrl(templateProjectSheetLink);
    let projectFolderRootURL = Spreadsheet.getCellValueLinkUrl(ProjectFoldersRoot);
    let projectFolderRoot = Folder.getByUrl(projectFolderRootURL);
    trace(`projectFolderRoot: ${projectFolderRoot.trace}`);
    let paymentsFoldersRootURL = Spreadsheet.getCellValueLinkUrl(PaymentsFoldersRoot);
    let paymentsFoldersRoot = Folder.getByUrl(paymentsFoldersRootURL);
    let templateName = Spreadsheet.getCellValue(SelectedTemplateName);
    if (!Dialog.confirm("Preparing Project Document Structure", `Preparing a new project document structure for ${this.name}, using template ${templateName}. Are you sure?`)) {
      return;
    }
    let projectFolder = projectFolderRoot.getSubfolder(this.fileName);
    if (projectFolder !== null) {
      Dialog.notify("Project folder already exists!", `Existing folder structure "${this.fileName}" will be used.`);
    }
    else {
      Dialog.toast("Preparing document structure, this may take a minute or two...");
      // First, create the basic project folder structure as a copy of the template structure
      projectFolder = templateFolder.copyTo(projectFolderRoot, this.fileName);  // Copies source folder contents (project template) to target folder (project structure), returns new target
      trace(`projectFolder: ${projectFolder}`);
    }
    let projectFolderLink = projectFolder.url;                                  // Gets the URL of newly created project folder 
    this.folderLink = projectFolderLink;                                        // Sets the folder link in the cell in FolderLink Column

    // Next, create the payments structure and link it to the main structure, be resilient to it already existing
    let paymentsFolderName = this.fileName + " - Payments";
    let paymentsFolder = paymentsFoldersRoot.getSubfolder(paymentsFolderName);
    if (paymentsFolder !== null) {
      Dialog.notify("Payments folder already exists!", "Linking it to the new structure.");
    } else {
      paymentsFolder = paymentsFoldersRoot.createFolder(paymentsFolderName);
    }
    let paymentsFolderLink = paymentsFolder.url;                              // Get the URL to payment folder
    this.paymentsLink = paymentsFolderLink;                                   // Set the URL in the master sheet

    if (!projectFolder.fileExists("Payments")) {
      projectFolder.createShortcut(paymentsFolder, "Payments");               // Creates shortcut to payment folder in project folder if it doesn't exist
    }
    else {
      Dialog.notify("Payments folder shortcut already exists!", "Please check that it is valid.");
    }

    let projectSheet = projectFolder.getFile(this.fileName);
    if (projectSheet !== null) {
      Dialog.notify("Project sheet already exists!", "Please check that it is valid.");
    }
    else {
      projectSheet = templateSheetFile.copyTo(projectFolder, this.fileName);
    }
    let projectSheetLink = projectSheet.url;                              // Get the URL of new client project sheet
    this.sheetLink= projectSheetLink;                                     // Set the project sheet link to the cell in SheetLink column

    Dialog.notify("Project Document Structure Created!", 'Please check columns "Project Sheet", "Project Folder" & "Project Payments Folder" for more details.');

//  let projectSheetId = projectSheet.id;                                 // Get the id of the newly created client sheet
//  let projectTemplate = Spreadsheet.openById(projectSheetId);
//  projectTemplate.setActive();                                          // Set the active spreadsheet Confirmed W & E to new project sheet
//  Browser.newTab(projectSheetLink);
    trace(`< prepareProjectStructure for ${this.trace} completed`);
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
