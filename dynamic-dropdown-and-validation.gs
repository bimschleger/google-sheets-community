/*

GOOGLE SHEETS AND APPS SCRIPT COMMUNITY ANSWER
In response to this question: https://support.google.com/docs/thread/56629498?hl=en&dark=1

Mostly asking about dynamic dropdowns and applying validation dynamically.

*/



/*

Update the sheet on every edit.

*/

function onEdit(){
  
  // Set function variables
  var tabLists = "Game Object Structure";
  var tabValidation = "Main One";
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datass = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabLists);
  
  // Get active cell
  var activeCell = ss.getActiveCell();
  
  if(activeCell.getColumn() == 2 && activeCell.getRow() > 2 && ss.getSheetName() == tabValidation){
    
    // Get header values for Game Object Structure headers
    var makes = datass.getRange(1, 1, 1, datass.getLastColumn()).getValues();
    
    // Get position of Active cell value in the list of Game Object Structure header values
    var makeIndex = makes[0].indexOf(activeCell.getValue()) + 1;
    
    // If the Active cell value is in the list of Game Object Structure header values
    if(makeIndex != 0){
    
      // Set validation values and rule
      var lastValidationRow = datass.getLastRow();
      var validationRange = datass.getRange(3, makeIndex, lastValidationRow);
      var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
      
      // Define variables that we will use in the while loop.
      var activeRow = activeCell.getRow();  // The row of teh active cell
      var activeColumn = 9;  // Define sarting column to loop (equivalent to column I)
      var lastColumn = ss.getLastColumn();  // Get last column, equivalent to column AE
      var updatedCell;  // The cell value to which we apply the validation rules
      
      // Increment through each of the columns from I to AE, applying the validation.
      while (activeColumn < lastColumn) {
        updatedCell = ss.getRange(activeRow, activeColumn);
        updatedCell.clearContent().clearDataValidations();
        updatedCell.setDataValidation(validationRule);
        activeColumn++;
      }
    }
  }
}