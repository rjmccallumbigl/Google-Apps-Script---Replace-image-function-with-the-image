/*********************************************************************************************************
*
* Instructions
* 1. Click/focus/select a cell with your =image(...) function in it.
* 2. Open Google Apps Script for your sheet.
* 3. Delete all text in the scripting window and paste all this code.
* 4. Run onOpen(). 
* 5. Then run 'Replace URL with Image' from the spreadsheet or 'insertImageOnSpreadsheet' from the Code.
* 6. Accept the permissions and after running, the spreadsheet should update.
*
*********************************************************************************************************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Functions')
    .addItem('Replace URL with Image', 'insertImageOnSpreadsheet')
    .addToUi();
}

/*********************************************************************************************************
*
* Replace highlighted  =image(...) function with the direct image it is referencing. This should make the
*   function less volatile.
* 
* References
* https://stackoverflow.com/questions/63859529/appsscript-replicate-put-image-in-selected-cell-on-google-sheets
* https://www.reddit.com/r/googlesheets/comments/mrtvlv/trying_to_display_picture_uploaded_via_google/
*
*********************************************************************************************************/
function insertImageOnSpreadsheet() {

  // Declare variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getCurrentCell();

  range.copyTo(range, { contentsOnly: true });
  SpreadsheetApp.flush();
}

