/**
 * Add menu option to sheet UI
 * Users have the options to perform 2 functions on click:
 * Generate folders and create contract
 * Or generate a downloadable PDF link
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Documents')
      .addItem('Generate Folders & Contracts', 'start')
      .addItem('Generate Downloadable PDFs', 'downloadFileAsPDF')
      .addToUi();
}

/** 
 * Create new folder in chosen parent folder using parameters passed in from start() function
 */

function createNewFolders(destinationID, folderName) {
  var parentLocation = DriveApp.getFolderById(destinationID);
  var createFolder = parentLocation.createFolder(folderName);
  createdfolderID = createFolder.getId();
  createdfolderLocation = DriveApp.getFolderById(createdfolderID);
  Logger.log(createFolder.getId());
  
}


/** 
 * Global variables to reference ID of created folder and its location
 */

var createdfolderID;
var createdfolderLocation;


/**
 * This function is called when users click to generate folders and contracts
 * Reads spreadsheet instances and passes parameters to relevant functions to alter contract type and destination
 */

function start(){

  var location = '1zCtK8J2Dqx06-xb9GVQthoUyUcz7lWXU';      // Default destination folder of folders to be generated

  const sheet = SpreadsheetApp.openById('1Ed2Xymlpe5NcSQnswnuqDThG6TL_uQRcjjUwlS0zsis').getSheetByName('Contract Details');
  const rows = sheet.getDataRange().getValues();
  var employmentBasis;
  Logger.log(rows)

  const empFullNameIndex = rows[0].indexOf('{{EMPLOYEE FULL NAME}}');
  Logger.log(empFullNameIndex);
  const linkToFolderIndex = rows[0].indexOf('Folder ID');
  Logger.log(linkToFolderIndex);
  const employmentType = rows[0].indexOf('Employment Type');
  const empRoleTypeIndex = rows[0].indexOf('Employment Role');

  rows.forEach(
    function(row, index) {

      if (index === 0) return; // Don't execute function for header row

      if (row[linkToFolderIndex]) return;       // If there is an existing value in the 'Link to Folder Cell', don't execute the function

      if (row[empRoleTypeIndex] == 'Hub Assistant'){
        location = '1zCtK8J2Dqx06-xb9GVQthoUyUcz7lWXU'
      }
      else if (row[empRoleTypeIndex] == 'Rider'){
        location = '1zCtK8J2Dqx06-xb9GVQthoUyUcz7lWXU'
      }
      else return

      var fName = row[empFullNameIndex];
      createNewFolders(location, fName); // Create folder using employee name, in the destination folder

      sheet.getRange(index + 1, linkToFolderIndex + 1).setValue(createdfolderID);

      switch (row[employmentType]) {
        case 'Casual':
          employmentBasis = 'C';
          createNewContract(index, employmentBasis);
          break;
        case 'PT':
          employmentBasis = 'PT';
          createNewContract(index, employmentBasis);
          break;
        case 'FT':
          employmentBasis = 'FT';
          createNewContract(index, employmentBasis);
          break;
        default:
          Logger.log(row[employmentType]);
          break;
      }
    })
}

/**
 * Generates individual contracts based on given paramaters
 */

function createNewContract(rowIndex, contractType) {
  
  var googleDocTemplate;
  var empType;
  const sheet = SpreadsheetApp.openById('1Ed2Xymlpe5NcSQnswnuqDThG6TL_uQRcjjUwlS0zsis').getSheetByName('Contract Details');
  const rows = sheet.getDataRange().getValues();

  // Set row reference index for columns
  const empFullNameIndex = rows[0].indexOf('{{EMPLOYEE FULL NAME}}');
  const docFileIndex = rows[0].indexOf('Link to Completed Contract');
  const empRoleTypeIndex = rows[0].indexOf('Employment Role');

  const headers = [];

  // Set Templates and document title depending on employment type and role
  if (rows[rowIndex][empRoleTypeIndex] == 'Rider'){

    if (contractType == 'C'){
      googleDocTemplate = DriveApp.getFileById('1b3Ex0v7vf2ojEArk6kYOiTQ6EGW3VUF7BedJjl4SnKM');
      empType = 'MILKRUN Offer of Rider Casual Employment - ';
      Logger.log("this passes through cas");
    }
    else if (contractType == 'PT'){
      googleDocTemplate = DriveApp.getFileById('1SGjufiBcOuXscoCJLSEUXKNhU8vqatPMAUNuyH3dzGE');
      empType = 'MILKRUN Part-Time Rider Employment Agreement - ';
      Logger.log("this passes through pt");
    }
    else if (contractType == 'FT'){
      googleDocTemplate = DriveApp.getFileById('1vsvsrOpIvI1aL4TSDnMJbpV4dApbmoqq2ZI4Rb9M2EY');
      empType = 'MILKRUN Full-Time Rider Employment Agreement - ';
      Logger.log("this passes through ft");
    }
  }
  else if (rows[rowIndex][empRoleTypeIndex] == 'Hub Assistant'){

    if (contractType == 'C'){
      googleDocTemplate = DriveApp.getFileById('1b3Ex0v7vf2ojEArk6kYOiTQ6EGW3VUF7BedJjl4SnKM');
      empType = 'MILKRUN Offer of Hub Assistant Casual Employment - ';
      Logger.log("this passes through cas");
    }
    else if (contractType == 'PT'){
      googleDocTemplate = DriveApp.getFileById('1SGjufiBcOuXscoCJLSEUXKNhU8vqatPMAUNuyH3dzGE');
      empType = 'MILKRUN Part-Time Hub Assistant Employment Agreement - ';
      Logger.log("this passes through pt");
    }
    else if (contractType == 'FT'){
      googleDocTemplate = DriveApp.getFileById('1vsvsrOpIvI1aL4TSDnMJbpV4dApbmoqq2ZI4Rb9M2EY');
      empType = 'MILKRUN Full-Time Hub Assistant Employment Agreement - ';
      Logger.log("this passes through ft");
    }
  }
  
  rows.forEach(
    function(row, index) {

      if (index === 0) {
        headers.push(row);
      }

      if (index === rowIndex){

        const copy = googleDocTemplate.makeCopy(empType + row[empFullNameIndex], createdfolderLocation);         // Create copy of template and set custom name
        const doc = DocumentApp.openById(copy.getId());
        Logger.log("Copy log " + copy.getDownloadUrl());
        const body = doc.getBody();

        // currently not dynamic
        for (let i = 0; i < headers[0].length - 5; i++) {
          body.replaceText(headers[0][i].toString(), row[i])       // Find all matching instances of column placeholders in the document body 
          Logger.log("Let's go");
        }
        
        doc.saveAndClose();

        const url = doc.getUrl();
        sheet.getRange(index + 1, docFileIndex + 1).setValue(url);

      }
      else {return}

    }
  )
}

/**
 * Enables users to generate a downloadable link which exports the respective Google Documents into a PDF
 */

function downloadFileAsPDF () {
  const sheet = SpreadsheetApp.openById('1Ed2Xymlpe5NcSQnswnuqDThG6TL_uQRcjjUwlS0zsis').getSheetByName('Contract Details');
  const rows = sheet.getDataRange().getValues();
  const pdfFileIndex = rows[0].indexOf('Link to Download PDF');
  const docFileIndex= rows[0].indexOf('Link to Completed Contract')

  rows.forEach(function (row, index) {
    if (index === 0) {return}

    if (row[pdfFileIndex]) {return}

    if (row[docFileIndex] == "") {return}

    let docFileID = row[docFileIndex].replace('https://docs.google.com/open?id=', "");
    Logger.log(docFileID);

    let pdfFile = 'https://docs.google.com/document/d/' + docFileID + '/export?format=pdf';
    Logger.log(pdfFile);

    sheet.getRange(index + 1, pdfFileIndex + 1).setValue(pdfFile);

  })
}
