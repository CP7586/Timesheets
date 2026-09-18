const SPREADSHEET_ID = '1qsA8Gq3z83Ihxx3MBnHhxn-xK8GsgZxN9MpVpdTQ4A0';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Service Call Worksheet Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fetch the URL of the Web App to handle proper page reloading inside the iframe
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Auto-initialize required sheets and header rows
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Users Sheet setup
  let userSheet = ss.getSheetByName("Users");
  if (!userSheet) {
    userSheet = ss.insertSheet("Users");
    userSheet.appendRow(["Username", "Password", "Role", "FullName"]);
    userSheet.appendRow(["admin", "admin123", "Admin", "System Administrator"]); // Default master admin
    userSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#c9daf8");
  }

  // 2. Time sheets Sheet setup
  let sheet = ss.getSheetByName("Time sheets");
  if (!sheet) {
    sheet = ss.insertSheet("Time sheets");
    sheet.appendRow([
      "Timestamp", "SubmittedBy", "Date", "Service Call #", "Customer Name", "Customer Phone",
      "CL Initials", "CW Initials", "Job Name", "Site Address", "Lot/Unit #", 
      "Time Left Shop", "Time Arrived Site", "Time Left Site", "Lunch Time",
      
      // Service 1
      "S1 Point/Room", "S1 Job Request", "S1 Labor", "S1 Billable", "S1 Billable Why",
      "S1 Job Done", "S1 Next Step", "S1 Billable Hours", "S1 Material Status", "S1 Material/PO#",
      
      // Service 2
      "S2 Point/Room", "S2 Job Request", "S2 Labor", "S2 Billable", "S2 Billable Why",
      "S2 Job Done", "S2 Next Step", "S2 Billable Hours", "S2 Material Status", "S2 Material/PO#",
      
      // Service Manager Use Only
      "Mgr Sign Off", "Mgr Info Accurate", "Mgr Request Schedule Not Completed",
      
      // Service Coordinator Use Only
      "Coord Sign Off", "Coord CC on File", "Coord Unbilled Handling", 
      "Coord Discounts/Negotiations", "Coord Approx Price Given",

      // Audit Info
      "Edited By"
    ]);
    sheet.getRange("A1:AR1").setFontWeight("bold").setBackground("#d9ead3");
  }
}

// Authenticate user credentials
function loginUser(username, password) {
  setupSheets();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === username.toString().toLowerCase().trim() && data[i][1].toString() === password.toString()) {
      return {
        success: true,
        user: {
          username: data[i][0],
          role: data[i][2],
          fullName: data[i][3]
        }
      };
    }
  }
  return { success: false, message: "Invalid username or password." };
}

// Register a new worker account (Admin functionality)
function registerWorker(newWorkerData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase().trim() === newWorkerData.username.toString().toLowerCase().trim()) {
      return { success: false, message: "Username already exists." };
    }
  }

  sheet.appendRow([
    newWorkerData.username,
    newWorkerData.password,
    newWorkerData.role,
    newWorkerData.fullName
  ]);

  return { success: true, message: "Worker account created successfully!" };
}

// Helper to format cell values safely into clean primitive strings
function formatCellVal(val, header) {
  if (val === undefined || val === null) return "";
  if (val instanceof Date) {
    if (header === "Date") {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (header.toLowerCase().includes("time") && !header.toLowerCase().includes("stamp")) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
    } else {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    }
  }
  return val.toString();
}

// Fetch dashboard timesheets sorted by date
function getTimesheets(username, role) {
  setupSheets();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Time sheets");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);
  let results = [];

  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let entry = { rowIndex: i + 2 };
    let rowSubmitter = "";

    headers.forEach((header, colIdx) => {
      let formattedVal = formatCellVal(row[colIdx], header);
      entry[header] = formattedVal;
      if (header === "SubmittedBy") {
        rowSubmitter = formattedVal.toString().toLowerCase().trim();
      }
    });

    // Admin sees all, Workers see only their own
    if (role === 'Admin' || rowSubmitter === username.toString().toLowerCase().trim()) {
      results.push(entry);
    }
  }

  // Sort by Date / Timestamp descending (newest first)
  results.sort((a, b) => {
    let dateA = new Date(a.Date || a.Timestamp || 0).getTime();
    let dateB = new Date(b.Date || b.Timestamp || 0).getTime();
    if (isNaN(dateA)) dateA = 0;
    if (isNaN(dateB)) dateB = 0;
    return dateB - dateA;
  });

  return results;
}

// Save or Update worksheet submission in Spreadsheet
function processForm(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Time sheets");

  // Helper to safely extract service entry data
  function getServiceData(i) {
    let prefix = "s" + i;
    let isBillable = formObject[prefix + "IsBillable"] || "No";
    let amountType = formObject[prefix + "BillableAmountType"];
    let specHours = formObject[prefix + "BillableHours"];
    
    let amount = "N/A";
    if (isBillable === 'Yes') {
      amount = (amountType === 'Specific Hours') ? (specHours || "") : 'Entire Time';
    }

    return [
      formObject[prefix + "PointRoom"] || "",
      formObject[prefix + "JobRequest"] || "",
      formObject[prefix + "Labor"] || "",
      isBillable,
      formObject[prefix + "BillableWhy"] || "",
      formObject[prefix + "JobDone"] || "No",
      formObject[prefix + "NextStep"] || "",
      amount,
      formObject[prefix + "MaterialStatus"] || "",
      formObject[prefix + "MaterialPO"] || ""
    ];
  }

  let s1Data = getServiceData(1);
  let hasService2 = formObject.s2PointRoom || formObject.s2JobRequest || formObject.s2Labor;
  let s2Data = hasService2 ? getServiceData(2) : ["", "", "", "No", "", "No", "", "N/A", "", ""];

  const rowData = [
    new Date(),
    formObject.submittedBy || "",
    formObject.date || "",
    formObject.serviceCallNumber || "",
    formObject.customerName || "",
    formObject.customerPhone || "",
    formObject.clInitials || "",
    formObject.cwInitials || "",
    formObject.jobName || "",
    formObject.siteAddress || "",
    formObject.lotUnitNumber || "",
    formObject.timeLeftShop || "",
    formObject.timeArrivedSite || "",
    formObject.timeLeftSite || "",
    formObject.lunchTime || "",

    // Service 1
    ...s1Data,

    // Service 2
    ...s2Data,

    // Manager
    formObject.mgrSignOff || "",
    formObject.mgrAccurate || "",
    formObject.mgrRequestSchedule || "",

    // Coordinator
    formObject.coordSignOff || "",
    formObject.coordCCOnFile || "",
    formObject.coordHandling || "",
    formObject.coordDiscounts || "",
    formObject.coordApproxPrice || "",

    // Audit
    ""
  ];

  // If updating an existing entry
  if (formObject.rowIndex && parseInt(formObject.rowIndex) > 1) {
    const rowNum = parseInt(formObject.rowIndex);
    const existingRange = sheet.getRange(rowNum, 1, 1, rowData.length);
    const existingVals = existingRange.getValues()[0];
    
    rowData[0] = existingVals[0]; // Retain original Timestamp
    rowData[1] = existingVals[1]; // ALWAYS retain original Submitter (Worker 1)

    // Mark who edited this row
    if (formObject.editorUser) {
      rowData[43] = formObject.editorUser;
    }

    sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
    return "Worksheet Updated Successfully!";
  } else {
    sheet.appendRow(rowData);
    return "Worksheet Saved Successfully!";
  }
}

// Delete a timesheet entry from the Spreadsheet
function deleteTimesheet(rowIndex, username, role) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Time sheets");
  if (!sheet) return { success: false, message: "Sheet not found." };

  const rowNum = parseInt(rowIndex);
  if (isNaN(rowNum) || rowNum <= 1 || rowNum > sheet.getLastRow()) {
    return { success: false, message: "Invalid row index." };
  }

  // Permission check
  const submitter = sheet.getRange(rowNum, 2).getValue().toString().toLowerCase().trim();
  if (role !== 'Admin' && submitter !== username.toString().toLowerCase().trim()) {
    return { success: false, message: "Unauthorized to delete this timesheet." };
  }

  sheet.deleteRow(rowNum);
  return { success: true, message: "Timesheet deleted successfully." };
}