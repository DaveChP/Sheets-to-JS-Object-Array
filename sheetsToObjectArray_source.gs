  /******************************************************* 
  *              SheetToObjectArray.gs                   *
  *******************************************************/
  const version = "1.0";
  /*******************************************************
   * Output will be saved to user's Google Drive 
   * if exportToDrive is set true. Set exportToDrive
   * false below to cancel export.
   * Edit docName below with preferred name.
   *******************************************************/
   
   const exportToDrive = true;
   const docName = "JSexport.doc";

  /*******************************************************
   * This script extracts data from a Google spreadsheet
   * and outputs it as a formatted javascript object array
   * and as JSON. 
   * To use, open a new Apps Script from the Extensions
   * menu of the spreadsheet to be extracted, paste this 
   * entire code, save, and run.
   * The output can be copied from the execution log
   * message on the Apps Script page, or retrieved from the
   * new Google Doc created in the Drive of the logged-in 
   * user.
   * *****************************************************/

   /*******************************************************
    * User Comments to be included in output can be added
    * to the notes variable below:
   * *****************************************************/

   const notes = ``;


function sheetToObjectArray() {

  const appUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const appId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const appSheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  const appName = SpreadsheetApp.getActiveSpreadsheet().getName();
  
  let outputTitle = "Javascript Object Array and JSON data Extracted from Google Spreadsheet\n";
  let sheetDetails= `Spreadsheet url - ${appUrl}\n\nSpreadsheet name - ${appName} (sheet name - ${appSheetName})\n`;
  let info = `Data extracted using SheetToObjectArray.gs (version ${version}), a Google Apps Script utility created by David Pritlove.\nGithub Repository: https://github.com/DaveChP/Sheets-to-JS-Object-Array \n\n`;

  const currentSheet = SpreadsheetApp.getActiveSheet();
  const rows = currentSheet.getRange(2,1,(currentSheet.getLastRow()-1),currentSheet.getLastColumn()).getValues();  
    if (rows[0][rows[0].length-1] == "") {
      for (let row=0; row<rows.length; row++) {rows[row].pop()}
    }
  const headData = currentSheet.getRange(1,1,1,rows[0].length).getValues(); 
  const head = headData[0].map(x => camelize(x));

  let jsObjectArray = "\n// javascript array of objects:\nconst data = [";
  let json = "// JSON:\n["

    for (let row=0; row<rows.length; row++) {
      jsObjectArray += "\n  {";
      json += "\n  {";

      for (let col=0; col<rows[0].length; col++) {
          if (rows[row][col] == parseFloat(rows[row][col])) {
            jsObjectArray += `${head[col]}: ${rows[row][col]}`;
            json += `"${head[col]}": ${rows[row][col]}`;
          }
          else {
            jsObjectArray += `${head[col]}: "${clean(rows[row][col].toString())}"`;
            json += `"${head[col]}": "${clean(rows[row][col].toString())}"`;
          }
        

        if(col == rows[0].length-1) {jsObjectArray += "}"; json += "}"} // end if last col;
        else {jsObjectArray += ", "; json += ", ";} // end else last col;

      } // next col;

    if (row == rows.length-1) {jsObjectArray += "\n];"; json += "\n];"} // end if last row;
    else {jsObjectArray += ", "; json += ", ";} // end else last row;

    } // next row;

  
  let date = new Date().toString();

  // output to Apps Script log;
  Logger.log(date + "\n\n" + outputTitle + "\n\n" + sheetDetails + "\n\n" + notes + "\n\n" + jsObjectArray + "\n\n" + json + "\n\n" + info);

  if (exportToDrive) {

  // output to Google Doc in Drive;
  const fileName = docName || "JSexport.doc";
  const newDoc = DocumentApp.create(fileName);
  const doc = DriveApp.getFileById(newDoc.getId());
  const body = DocumentApp.openById(newDoc.getId()).getBody();

  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Courier New';
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  style[DocumentApp.Attribute.UNDERLINE] = true;

  let titleLine = body.appendParagraph(outputTitle);
  titleLine.setAttributes(style);


  let dateLine = body.appendParagraph(date + "\n");
  let detailsLine = body.appendParagraph(sheetDetails);
  style[DocumentApp.Attribute.UNDERLINE] = false;
  dateLine.setAttributes(style);
  detailsLine.setAttributes(style);

  if (notes.length > 0) {
    let notesLine = body.appendParagraph(notes + "\n")
    notesLine.setAttributes(style);
  } // end if notes;

  let arrayLine = body.appendParagraph(jsObjectArray);
  style[DocumentApp.Attribute.FONT_SIZE] = 10;
  arrayLine.setAttributes(style);

  let jsonLine = body.appendParagraph("\n\n" + json);
  jsonLine.setAttributes(style);

  let infoLine = body.appendParagraph("\n\n" + info)
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  style[DocumentApp.Attribute.ITALIC] = true;
  infoLine.setAttributes(style);

  } // end if exportToDrive;

} // end sheetToObjectArray function;

function clean(val) {
  // backslash-escapes double quote marks in passed string;
  return val.replace(/"/g, '\\"');
}

function camelize(str) {
  // function returns the camelCase version of the passed string;
  // see https://stackoverflow.com/questions/2970525/converting-any-string-into-camel-case
  // Christian C. Salvadó's answer to SO question;

  return str.replace(/(?:^\w|[A-Z]|\b\w|\s+)/g, function(match, index) {
    if (+match === 0) return ""; // or if (/\s+/.test(match)) for white spaces
    return index === 0 ? match.toLowerCase() : match.toUpperCase();
  });
} // end camelize function;

