/*

GOOGLE SHEETS COMMUNITY ANSWER
I posted a response to a community question on 2020.07.01.

https://support.google.com/docs/thread/56639613?hl=en&dark=1

Hi, is there a script for google sheets to delete certain values inside a cell?
---
Hi, I need to delete certain values inside a cell but the problem is there are multiple values inside that cell. 
Is there anyway to delete cell values by using font color? 
I need to delete the ones in RED and keep the ones in BLACK font color.

*/


/*

Sets a menu UI item in the top navigation to remove red values.

*/

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Delete red values', functionName: 'deleteRedValues'}
  ];
  spreadsheet.addMenu('Delete...', menuItems);
}


/*

Gets the text style runs of a cell. If the color is red, delete the run. Then repack the content and add to the cell.

*/

function deleteRedValues() {
  
  // Gets the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeCell = sheet.getActiveCell();
  
  var newContentArray = [];
  var colorRed = "#ff0000";
  
  // Gets the rich text information from the cell
  var richText = activeCell.getRichTextValue();  // Reference: https://developers.google.com/apps-script/reference/spreadsheet/rich-text-value
  
  // Each run is the longest possible substring having a consistent text style. 
  var runs = richText.getRuns(); // Reference: https://developers.google.com/apps-script/reference/spreadsheet/rich-text-value#getRuns()
  
  // For each run, check if the text color is red. If no, add the text to a new array.
  runs.forEach(function(run) {
    var runText = run.getText(); // Reference: https://developers.google.com/apps-script/reference/spreadsheet/rich-text-value#getText()
    var runTextColor = run.getTextStyle().getForegroundColor(); // Reference: https://developers.google.com/apps-script/reference/spreadsheet/text-style#getForegroundColor()
    
    Logger.log("Got run value: '" + runText + "' and text color '" + runTextColor + "'.");  // useful to verify what hex your text is
    
    if (runTextColor != colorRed) {
      newContentArray.push(runText); // Add non-red values to a new array
      Logger.log("Added '" + runText + "' to newContentArray.");
    }
  });
  
  // Join all non-red values together in a new string.
  let newContent = newContentArray.join("");

  // Set active cell to the new string on non-red values.
  activeCell.setValue(newContent);
}