// Set up the document variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetsCount = ss.getNumSheets();
var sheets = ss.getSheets();
var mysheet = ss.getSheets()[0];
var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function onOpen() { 
  // Create element in the menu
  try{
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Lucify')
    .addItem('Create Table', 'tableMenu') 
    .addToUi(); 
  } 
  
  // Catch errors
  catch (e){Logger.log(e)}
  
  // If error then use the old method
  finally{
    var items = [
      {name: 'Create Table', functionName: 'tableMenu'},
      
    ];
      ss.addMenu('Lucify', items);
  }
}

function tableMenu() {
  // Create the sidebar from the sidebar.html file
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
  .setTitle('Choose a Table');

  // Display that in the spreadsheet
  SpreadsheetApp.getUi().showSidebar(ui);
}

function basicTableChoice() {
  // Create the basic table page, from the basicTable.html file
  var ui = HtmlService.createHtmlOutputFromFile('basicTable')
  .setTitle('Customize your table');

  SpreadsheetApp.getUi().showSidebar(ui);
}

function databaseChoice() {
  // Create the database making page, from the database.html file
  var ui = HtmlService.createHtmlOutputFromFile('database')
  .setTitle('Customize your table');

  SpreadsheetApp.getUi().showSidebar(ui);
}

function createBasicTable(indep, dep, trials, check, rows) {
  // Find which cell the user has clicked, because the table will start there 
  var currentCell = SpreadsheetApp.getCurrentCell();
  var currentCellID = currentCell.getA1Notation();

  // Letters of the alphabet, to use for the a1 annotation, eg. Z3, F5
  letters = alphabet.split("");
  listID = currentCellID.split("");
  
  // Set the boundaries for the spreadsheet
  var myRange = mysheet.getRange(currentCellID + ":" + "Z1000");

  // The independent variable goes in the top left cell of the table
  currentCell.setValue(indep);
  // Merge that cell with the one below
  ss.getRange(currentCellID+":"+myRange.getCell(2,1).getA1Notation()).merge();
  currentCell.setWrap(true);

  // Loop horizontally across the spreadsheet, based on the number of trials
  for(var i = 0; i < trials; i++) {
    // Merge the cells for the dependent variable
    ss.getRange(myRange.getCell(1,2).getA1Notation()+":"+myRange.getCell(1,2+i).getA1Notation()).merge();

    // Put the name of the cell, based on which trial it is
    var trialCell = myRange.getCell(2, 2+i);
    trialCell.setValue("Trial " + (i+1).toString());

    // If this is the last iteration, check whether the user wants to have an average column
    if (i == (trials - 1)) {
      if (check == true) {
        // Put the name of the last trial cell
        ss.getRange(myRange.getCell(1,2).getA1Notation()+":"+myRange.getCell(1,3+i).getA1Notation()).merge();
        var trialCell = myRange.getCell(2, 3+i);
        // Put the word average, in the average cell
        var final = trialCell.getA1Notation();
        trialCell.setValue("Average");
        var bottom = myRange.getCell(2+rows, 3+i);
      } else if (check == false) {
        // If the user doesn't want an average column, just label the last trial cell
        var bottom = myRange.getCell(2+rows, 2+i);
      }
    }
  }
  // Set the name of the dependent variable cell
  var depCell = myRange.getCell(1, 2);
  depCell.setValue(dep); 
  depCell.setWrap(true);

  // Get the last cell, bottom right
  var bottomID = bottom.getA1Notation()
  // Set borders around all of the cells
  mysheet.getRange(currentCellID + ":" + bottomID).setBorder(true, true, true, true, true, true);
}
