//=============================================================================================
// Class TaskList
//

class TaskList {

  constructor() {
    trace("constructing Asana TaskList object...");
    this.range = Asana.asanaTaskListRange;
    this.rowCount = this.range.height;
    trace("NEW " + this.trace);
  }

  // Method apply
  // Iterate over all rows (using Range Row Iterator), call handler methods 

  apply(handler) {
    trace(`${this.trace}.apply`);
    this.range.forEachRow((range) => {
      const row = new TaskRow(range);
      handler.onRow(row);
    });
    handler.onEnd();
  }

  get trace() {
    return `{Asana TaskList range=${this.range.trace} rowCount=${this.rowCount}`;
  }

}

//=============================================================================================
// Class TaskRow
//

class TaskRow extends RangeRow {

  constructor(range, values = null) {
    super(range, values);
  }

  get emailToSend()   { return this.get("EmailToSend", "string"); }
  get type()          { return this.get("Type", "string"); }
  get category()      { return this.get("Category", "string"); }
  get name()          { return this.get("Name", "string"); }
  get subtaskOf()     { return this.get("Subtaskof", "string"); }
  get dependents()    { return this.get("Dependents", "string"); }
  get description()   { return this.get("Description", "string"); }
  get notes()         { return this.get("Notes", "string"); }
  get section()       { return this.get("Section", "string"); }
  get assignee()      { return this.get("Assignee", "string"); }
  get startDate()     { return this.get("StartDate"); }
  get dueDate()       { return this.get("DueDate"); }

}


//=============================================================================================
// Class WeddingPhaseList
//

class WeddingPhaseList {

  constructor() {
    trace("constructing Asana TaskList object...");
    this.range = Asana.asanaWeddingPhaseRange;
    this.rowCount = this.range.height;
    trace("NEW " + this.trace);
  }

  // Method apply
  // Iterate over all rows (using Range Row Iterator), call handler methods 

  apply(handler) {
    trace(`${this.trace}.apply`);
    this.range.forEachRow((range) => {
      const row = new PhaseRow(range);
      handler.onRow(row);
    });
    handler.onEnd();
  }

  get trace() {
    return `{Asana TaskList range=${this.range.trace} rowCount=${this.rowCount}`;
  }

}

//=============================================================================================
// Class PhaseRow
//

class PhaseRow extends RangeRow {

  constructor(range, values = null) {
    super(range, values);
  }

  get weddingPhase()   { return this.get("WeddingPhases", "string"); }

}



