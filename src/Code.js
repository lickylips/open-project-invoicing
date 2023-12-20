

/**
 * getSettings - Function to get settings from the tab called "Settings" in the spreadsheet
 * @param {string} docId 
 * @return {object} settings object containing the settings  
 */
function getSettings(docId){
  const ss = SpreadsheetApp.openById(docId);
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsArray = settingsSheet.getDataRange().getValues();
  const settings = {};
  for(i in settingsArray){
    settings[settingsArray[i][0]] = settingsArray[i][1];
  }
  settings.docId = ss.getId();
  return settings;
}

/**
 * getDetails pulls billable hours from the spreadsheet to include in the next invoice
 * @param {string} docId 
 * @returns {array} details An array of classes containing details to be invoiced
 */
function getDetails(docId){
  const ss = SpreadsheetApp.openById(docId);
  const settings = getSettings(docId);
  class InvoiceDetail {
    constructor(date, project, hours, invoiced, rate, notes, settings, index){
      this.date = date;
      this.project = project;
      this.hours = hours;
      this.invoiced = invoiced;
      this.rate = rate;
      this.notes = notes;
      this.settings = settings;
      this.index = index;
    }
    invoiceNumber(){
      const date = new Date();
      const dateFormatted = Utilities.formatDate(date, "GMT", "yyyyMMdd");
      const invoiceNumber = this.settings.companyId+dateFormatted;
      return invoiceNumber;
    }
    amount(){
      const amount = Number(this.hours)*Number(this.rate);
      return amount;
    }
    invoiceDueDate(){
      const date = new Date();
      date.setMonth(date.getMonth()+1);
      const invoiceDueDate = Utilities.formatDate(date, "GMT", "dd MMMMM, yyyy");
      return invoiceDueDate;
    }
  }
  const detailsSheet = ss.getSheetByName("Billable Hours");
  const detailsArray = detailsSheet.getDataRange().getValues();
  //Define Header Columns
  let dateCol, projectCol, hoursCol, invoicedCol, rateCol, notesCol;
  for(i in detailsArray[0]){ 
    if(detailsArray[0][i]=="Date"){dateCol = Number(i);}
    if(detailsArray[0][i]=="Project"){projectCol = Number(i);}
    if(detailsArray[0][i]=="Hours"){hoursCol = Number(i);}
    if(detailsArray[0][i]=="Invoiced"){invoicedCol = Number(i);}
    if(detailsArray[0][i]=="Rate"){rateCol = Number(i);}
    if(detailsArray[0][i]=="Notes"){notesCol = Number(i);}
  }
  detailsArray.shift();
  const details = [];
  for(i in detailsArray){
    let date=detailsArray[i][dateCol];
    let project=detailsArray[i][projectCol];
    let hours=detailsArray[i][hoursCol];
    let invoiced=detailsArray[i][invoicedCol];
    let rate=detailsArray[i][rateCol];
    let notes = detailsArray[i][notesCol];
    if(!invoiced && date != ""){
      let detail = new InvoiceDetail(date, project, hours, invoiced, rate, notes, settings, Number(i)+2);
      details.push(detail);
    }
  }
  return details;
}

/**
 * buildInvoice A function to build an invoice document
 * @param {object} settings 
 * @param {array} details 
 */
function buildInvoice(settings, details){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName("Billable Hours");
  const today = new Date();
  const todayFormatted = Utilities.formatDate(today, "GMT", "dd MMMMM, yyyy");
  //copy template
  const invoiceFolder = DriveApp.getFolderById("1o21n-QVGuMZHyX5mN59SoRoADdkE0rMq");
  const invoiceTemplateId = "1ZdPSflib6fRPTn1VPsixzGoCGuGXMR1SHTJ2QTpj4Ak";
  const invoiceTemplate = DriveApp.getFileById(invoiceTemplateId);
  const invoiceDoc = invoiceTemplate.makeCopy();
  invoiceDoc.setName("Invoice "+details[0].invoiceNumber());
  invoiceDoc.moveTo(invoiceFolder);
  const invoiceDocUrl = invoiceDoc.getUrl();
  const invoiceDocId = invoiceDoc.getId();

  //open doc for editing
  const invoice = DocumentApp.openById(invoiceDocId);
  const body = invoice.getBody();
  body.replaceText("{{INVOICE NUMBER}}", details[0].invoiceNumber());
  body.replaceText("{{COMPANY NAME}}", settings.companyName);
  body.replaceText("{{COMPANY ADDRESS}}", settings.companyAddress);
  body.replaceText("{{ISSUE DATE}}", todayFormatted);
  body.replaceText("{{DUE DATE}}", details[0].invoiceDueDate());

  //find details table
  let total = 0;
  let detailsTable;
  const tables = body.getTables();
  for(i in tables){
    if(tables[i].getText().includes("Details")){
      detailsTable = tables[i];
      for(j in details){
        //add a new row
        let newRow = detailsTable.appendTableRow();
        //put in the details
        Logger.log(details[j].date)
        let formatDate = Utilities.formatDate(details[j].date, "GMT", "yyyy-MM-dd");
        newRow.appendTableCell().setText(formatDate);
        newRow.appendTableCell().setText("["+details[j].project+"] "+details[j].notes);
        newRow.appendTableCell().setText(parseFloat(details[j].hours).toFixed(2));
        newRow.appendTableCell().setText("€"+details[j].rate);
        newRow.appendTableCell().setText("€"+details[j].amount().toFixed(2));
        total += details[j].amount();
        //mark as invoiced in sheet
        detailSheet.getRange(details[j].index, 5).setValue(true);
        //add invoice number to sheet
        const formula = '=HYPERLINK("' + invoiceDocUrl + '", "' + details[j].invoiceNumber() + '")'
        detailSheet.getRange(details[j].index, 8).setFormula(formula)
      }
      
    }
  }
  //remove extra row
  detailsTable.removeRow(2);
  body.replaceText("{{TOTAL PRICE}}", "€"+parseFloat(total).toFixed(2));
}

/**
 * runMain Function that will be called from the spreadsheet to create an invoice
 */
function runMain(){
  const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const details = getDetails(docId);
  const settings = getSettings(docId);
  if(details.length == 0){
    Logger.log("No proejcts to invoice")
  } else{
    buildInvoice(settings, details);
  }
}

/**
 * Calculate the distance between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DISTANCE("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The distance in miles
 * @customFunction
 */
const GOOGLEMAPS_DISTANCE = (origin, destination, mode) => {
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();

  if (!data) {
    throw new Error('No route found!');
  }

  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
  return distance;
};




/**
 * fetchOpenProjectData function to pull the required info from open project
 * @param {object} settings 
 * @returns 
 */
function fetchOpenProjectData(settings){
  const baseUrl = settings.apiBaseUrl;
  const projectsUrl = generateProjectAPICall(settings.companyProjectString);
  let data = importJSON(projectsUrl, settings.docId);
  const output = [];
  for(element of data._embedded.elements){
    let workPackageUrl = generateWorkPackageAPICall(element.id);
    let workPackages = importJSON(workPackageUrl);
    for(wpElement of workPackages._embedded.elements){
      let timeEntriesUrl = baseUrl+wpElement._links.timeEntries.href;
      let timeEntries = importJSON(timeEntriesUrl);
      for(timeEntry of timeEntries._embedded.elements){
        if(timeEntry){
          let hours = convertDurationToHours(timeEntry.hours);
          let row = {
            date: timeEntry.spentOn,
            project: wpElement._links.project.title,
            comment: timeEntry.comment.raw,
            type: wpElement._links.type.title,
            hours: hours,
            title: wpElement.subject,
            timeEntryId: timeEntry.id,
            user: timeEntry._links.user.title
          };
          output.push(row);
        }
      }
    }
  }
  return output
}

/**
 * importTimeEntries - Function to import time entries to spreadsheet
 * 
 */
function importTimeEntries() {
  // Get the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Get the settings from the settings worksheet
  const settings = getSettings(ss.getId());
  // Get the billable hours sheet
  const billableSheet = ss.getSheetByName("Billable Hours");
  // Get the header row and find the right column numbers
  const headerCols = billableHeaders(billableSheet.getRange(1, 1, 1, billableSheet.getLastColumn()).getValues()[0]);
  // Get the data from the billable hours sheet
  const billableHours = billableSheet.getRange(2, headerCols.dateCol + 1, billableSheet.getLastRow() - 1, billableSheet.getLastColumn()).getValues();
  // Create the time entries array
  const timeEntries = [];
  let bsRow = 2;
  let lastTimeEntryId = 0;
  // Iterate through each row in the billable hours sheet
  for (row of billableHours) {
    let timeEntry = createTimeEntry(row, headerCols);
    timeEntry.row = bsRow;
    // Check if the time entry is invoiced if not update sheet
    if(timeEntry.invoiced === false){
      Logger.log("Not Invoiced, needs updating row "+bsRow+" with time entry id "+timeEntry.timeEntryId);
      let updatedTimeEntry = createTimeEntryObject(timeEntry.timeEntryId, settings);
      updatedTimeEntry.row = bsRow;
      Logger.log(updatedTimeEntry)
      updateRow(updatedTimeEntry, headerCols);
    }
    bsRow++;
    timeEntries.push(timeEntry);
    // Check if the time entry ID is greater than the last time entry ID
    if(lastTimeEntryId < timeEntry.timeEntryId){
      lastTimeEntryId = timeEntry.timeEntryId;
    }
  }
  // Find the URL for the api call that will pull un imported time entries
  Logger.log("Last Time Entry ID "+lastTimeEntryId);
  const now = new Date();
  const dateOfLastTimeEntry = Utilities.parseDate(createTimeEntryObject(lastTimeEntryId, settings).createdAt, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const start = new Date(dateOfLastTimeEntry.getTime()+1000);
  const filters = [
    {
      "created_at": {
        "operator": "<>d",
        "values": [start, now.toISOString()]
      }
    }
  ];
  let url = "https://projects.lickylips.duckdns.org/api/v3/time_entries/?"+"filters="+JSON.stringify(filters);
  let stringifyUrl = encodeURI(url);
  Logger.log(stringifyUrl);

  // Get the un imported time entries
  const newTimeEntryResult = importJSON(stringifyUrl, settings);
  const newTimeEntries = newTimeEntryResult._embedded.elements;
  Logger.log(newTimeEntries.length)
  for(newTimeEntry of newTimeEntries){
    Logger.log("New Time Entry "+newTimeEntry.id);
    let newTimeEntryObject = createTimeEntryObject(newTimeEntry.id, settings);
    newTimeEntryObject.row = bsRow;
    bsRow++;
    timeEntries.push(newTimeEntryObject);
    newTimeEntryObject.invoiced = false;
    newTimeEntryObject.invoiceNumber = "";
    Logger.log("Adding new Time Entry " + newTimeEntryObject.timeEntryId+" on row "+bsRow);
    updateRow(newTimeEntryObject, headerCols);
  }
}

/**
 * Function to figure ot the columns of the data
 * @param {array} headerRow row containing headers  
 * @returns 
 */
function billableHeaders(headerRow){
  const headers = {};
  for(i in headerRow){
    if(headerRow[i] == "Date"){headers.dateCol = Number(i);}
    if(headerRow[i] == "Project"){headers.projectCol = Number(i);}
    if(headerRow[i] == "Hours"){headers.hoursCol = Number(i);}
    if(headerRow[i] == "Company Sponsor"){headers.companyCol = Number(i);}
    if(headerRow[i] == "Invoiced"){headers.invoicedCol = Number(i);}
    if(headerRow[i] == "Rate"){headers.rateCol = Number(i);}
    if(headerRow[i] == "Amount"){headers.amountCol = Number(i);}
    if(headerRow[i] == "Invoice Number"){headers.invoiceNumberCol = Number(i);}
    if(headerRow[i] == "Notes"){headers.notesCol = Number(i);}
    if(headerRow[i] == "Time Entry ID"){ headers.timeEntryIdCol = Number(i);}
  }
  return headers;
}

/**
 * Function to create a time entry object
 * @param {array} row row containing data
 * @param {object} headers object containing the columns
 * @returns {object} timeEntry object
 */
 function createTimeEntry(row, headers){
  const timeEntry = {};
  timeEntry.date = row[headers.dateCol];
  timeEntry.project = row[headers.projectCol];
  timeEntry.hours = row[headers.hoursCol];
  timeEntry.company = row[headers.companyCol];
  timeEntry.invoiced = row[headers.invoicedCol];
  timeEntry.rate = row[headers.rateCol];
  timeEntry.amount = row[headers.amountCol];
  timeEntry.invoiceNumber = row[headers.invoiceNumberCol];
  timeEntry.notes = row[headers.notesCol];
  timeEntry.timeEntryId = row[headers.timeEntryIdCol];
  return timeEntry;
 }

 /**
 * importJSON function to take url and return data
 * @param {string} url 
 * @returns 
 */
function importJSON(url, settings) {
  const username = "apikey";
  const password = settings.apiKey
  const apiToken = Utilities.base64Encode(username+":"+password);
  const headers = {
      "Authorization" : "Basic "+apiToken, // Insert a Basic Auth Token of an OpenProject account to get access to the API
  };
  const params = {
      "method": "GET",
      "headers": headers
  };
  const response = UrlFetchApp.fetch(url, params);
  const json = response.getContentText();
  const data = JSON.parse(json);
  return data;
} 

/**
 * convertDurationToHours function to take a string of duration and convert it to an amount of hours
 * @param {string} durationString 
 * @returns 
 */
function convertDurationToHours(durationString) {
  // Extract hours, minutes, and seconds using regular expressions
  var regex = /PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/;
  var match = durationString.match(regex);

  if (match) {
    var hours = match[1] ? parseInt(match[1]) : 0;
    var minutes = match[2] ? parseInt(match[2]) : 0;
    var seconds = match[3] ? parseInt(match[3]) : 0;
    
    // Calculate total duration in hours
    var totalHours = hours + minutes / 60 + seconds / 3600;
    return totalHours;
  } else {
    throw new Error("Invalid duration format");
  }
}

/**
 * updateRow function to update a row in a spreadsheet
 * @param {object} timeEntry object
 */
function updateRow(timeEntry, headerCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const billableSheet = ss.getSheetByName("Billable Hours");
  const row = [];
  row[headerCols.dateCol] = timeEntry.date;
  row[headerCols.projectCol] = timeEntry.project;
  row[headerCols.hoursCol] = timeEntry.hours;
  row[headerCols.companyCol] = timeEntry.company;
  row[headerCols.invoicedCol] = timeEntry.invoiced;
  row[headerCols.rateCol] = timeEntry.rate;
  row[headerCols.amountCol] = timeEntry.hours * timeEntry.rate;
  row[headerCols.invoiceNumberCol] = timeEntry.invoiceNumber;
  row[headerCols.notesCol] = timeEntry.notes;
  row[headerCols.timeEntryIdCol] = timeEntry.timeEntryId;
  billableSheet.getRange(timeEntry.row, 1, 1, billableSheet.getLastColumn()).setValues([row]);
  billableSheet.getRange(timeEntry.row, headerCols.invoicedCol+1).insertCheckboxes();
}

/**Function to create a time entry object from the result of an api call
 * @param {object} timeEntryResult result of time entry api call
 * @param {object} settings object containing settings
 * @returns {object} timeEntry object
 */
function createTimeEntryObject(timeEntryId, settings) {
  const url = "https://projects.lickylips.duckdns.org/api/v3/time_entries/"+timeEntryId;
  const timeEntryResult = importJSON(url, settings);
  const timeEntry = {};
  timeEntry.date = timeEntryResult.spentOn;
  timeEntry.project = timeEntryResult._embedded.project.name;
  timeEntry.hours = convertDurationToHours(timeEntryResult.hours);
  timeEntry.company = settings.companyName;
  timeEntry.rate = settings.rate;
  timeEntry.amount = timeEntry.hours * settings.rate;
  timeEntry.notes = timeEntryResult._embedded.workPackage.subject + " - " + timeEntryResult.comment.raw;
  timeEntry.timeEntryId = timeEntryResult.id;
  timeEntry.createdAt = timeEntryResult.createdAt;
  return timeEntry;
}

/**
 * On Open of billable hours spreadsheet
 * 
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a main menu item
  ui.createMenu('Invoicing')
      .addItem('Generate Invoices', 'runMain')
      .addItem("Import Open Project Data", "importTimeEntries")
      .addToUi();
}