/**
 * On Open of billable hours spreadsheet
 * 
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a main menu item
  ui.createMenu('Invoicing')
      .addItem('Generate Invoices', 'runMain')
      .addItem("Import Open Project Data", "writeTimeEntries")
      .addToUi();
}

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
 * importJSON function to take url and return data
 * @param {string} url 
 * @returns 
 */
function importJSON(url) {
    let username = "apikey";
    let password = "eae40f509c71e16019bbd96c0923d7e0ed52c03b2af7452c0bb287ff6889bd8e"
    let apiToken = Utilities.base64Encode(username+":"+password);
    var headers = {
        "Authorization" : "Basic "+apiToken, // Insert a Basic Auth Token of an OpenProject account to get access to the API
    };
    var params = {
        "method": "GET",
        "headers": headers
    };
    var response = UrlFetchApp.fetch(url, params);
    var json = response.getContentText();
    var data = JSON.parse(json);
    return data;
} 

/**
 * fetchOpenProjectData function to pull the required info from open project
 * @param {object} settings 
 * @returns 
 */
function fetchOpenProjectData(settings){
  const baseUrl = settings.apiBaseUrl;
  const projectsUrl = generateProjectAPICall(settings.companyProjectString);
  let data = importJSON(projectsUrl);
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
 * writeTimeEntries function called from the spreadsheet to write new time entries
 */
function writeTimeEntries(){
  Logger.log("Updateing Data From OpenProject")
  const ss = SpreadsheetApp.getActiveSpreadsheet(); //open spreadsheet
  const settings = getSettings(ss.getId()); //pull settings
  const billableSheet = ss.getSheetByName("Billable Hours"); //get the sheet we will write to
  const existingHours = billableSheet.getDataRange().getValues(); //get the existing data
  const newHours = fetchOpenProjectData(settings);
  const headerCols = billableHeaders(existingHours[0]); //find the col numbers of the existing data
  existingHours.shift(); //Drop the header row from the existing data
  //create array of existing timeIds
  const existingTimeIdsRaw = billableSheet.getRange(2, headerCols.timeEntryIdCol+1,billableSheet.getLastRow()-1).getValues();
  const existingTimeIds = [];
  for(row of existingTimeIdsRaw){
    existingTimeIds.push(row);
  }
  //create array of new time IDs
  const newTimeIds = []
  for(row of newHours){
    newTimeIds.push(row.timeEntryId);
  }
  let toDelete = []; //Initialise variable for deleting obsolete rows
  //Iterate through each row in the billable sheet
  for(i=2; i<billableSheet.getLastRow()+1; i++){
    let range = billableSheet.getRange(i, 1, 1, billableSheet.getLastColumn());
    let rowData = range.getValues();
    let rowTimeId = rowData[0][headerCols.timeEntryIdCol];//get the time ID of this row
    let invoice = rowData[0][headerCols.invoiceNumberCol];//get if this row has a created invoice listed
    Logger.log("Processing Row "+range.getA1Notation()+" With time id "+ rowTimeId);
    //Identify rows to delete if time ID is no longer present and row has not been invoiced yet
    if(newTimeIds.indexOf(rowTimeId)==-1 && !invoice){
      Logger.log("Time ID "+rowTimeId+" No Longer Exists, marking row for deletion "+range.getA1Notation());
      toDelete.push(range.getRow());
    }

    //If new time ID is present in row update relivant fields and drop new time id 
    if(newTimeIds.indexOf(rowTimeId)!=-1){
      Logger.log("Time ID "+rowTimeId+" Exists on row "+range.getA1Notation()+". Updating row & dropping Item")
      let newHour
      for(k in newHours){
        if(newHours[k].timeEntryId == rowTimeId){
            newHour=newHours[k];
            newHours.splice(Number(k), 1);
          }
      }
      Logger.log(newHour)
      billableSheet.getRange(range.getRow(), headerCols.dateCol + 1).setValue(newHour.date);
      billableSheet.getRange(range.getRow(), headerCols.hoursCol + 1).setValue(newHour.hours);
      billableSheet.getRange(range.getRow(), headerCols.notesCol + 1).setValue(newHour.title + " - " + newHour.comment);
    }
    Logger.log("=====================================");
  }

  //Write new hours to sheet
  for(i in newHours){
    Logger.log("Adding new Time Entry " + newHours[i].timeEntryId);
    let row = newHours[i];
    let newRow = [
        row.date,
        row.project,
        row.hours,
        settings.companyName,
        false,
        settings.rate,
        row.hours * settings.rate,
        "",
        row.title + " - " + row.comment,
        row.timeEntryId
    ];
    let newSheetRow = billableSheet.appendRow(newRow);
    newSheetRow.getRange(billableSheet.getLastRow(), headerCols.invoicedCol + 1)
        .insertCheckboxes();
    let rateCell = newSheetRow.getRange(billableSheet.getLastRow(), headerCols.rateCol + 1).getA1Notation();
    let hoursCell = newSheetRow.getRange(billableSheet.getLastRow(), headerCols.hoursCol + 1).getA1Notation();
    newSheetRow.getRange(billableSheet.getLastRow(), headerCols.amountCol + 1)
        .setFormula("=" + rateCell + "*" + hoursCell);
  }
  //Delete obsolete rows
  for(row of toDelete){
    Logger.log("Deleting Obsolete Row "+row);
    billableSheet.deleteRow(row);
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
 * Function to generate the URL required to call projects
 * @param {*} string 
 * @returns 
 */
function generateProjectAPICall(string) {
  const baseUrl = "https://projects.lickylips.duckdns.org";
  let url = baseUrl+"/api/v3/projects/?";
  const filters = [
    {
      "name_and_identifier": {
        "operator": "~",
        "values": [string]
      }
    }
  ];
  const pageSize = 100;

  url += "filters=" + JSON.stringify(filters);
  url += "&pageSize=" + pageSize;
  return encodeURI(url);
}

/**
 * Function to generate the URL required to call work packages
 * @param {*} string 
 * @returns 
 */
function generateWorkPackageAPICall(string) {
  const baseUrl = "https://projects.lickylips.duckdns.org";
  let url = baseUrl+"/api/v3/projects/";
  url += string+"/work_packages/?";
  const pageSize = 100;
  const filters = [
    {
      "status":{
        "operator": "*",
        "values": []
      }
    }
  ];
  url += "filters="+JSON.stringify(filters);
  url += "&pageSize=" + pageSize;
  return encodeURI(url);
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

