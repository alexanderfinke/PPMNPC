// Google Apps Script backend for Cost Tracking Application

const SHEET_NAME = "Costs";

// Initialize the spreadsheet with headers if not exists
function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ["ID", "Item Name", "Category", "Planned Cost", "Actual Cost", "Variance", "Status", "Date Created", "Notes"];
    sheet.appendRow(headers);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#4285F4");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
  }
  
  return sheet;
}

// Get all cost entries
function getCostEntries() {
  const sheet = initializeSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return [];
  }
  
  const headers = data[0];
  const entries = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    entries.push({
      id: row[0],
      itemName: row[1],
      category: row[2],
      plannedCost: row[3],
      actualCost: row[4],
      variance: row[5],
      status: row[6],
      dateCreated: row[7],
      notes: row[8]
    });
  }
  
  return entries;
}

// Add a new cost entry
function addCostEntry(itemName, category, plannedCost, notes) {
  const sheet = initializeSheet();
  const id = Utilities.getUuid();
  const dateCreated = new Date().toLocaleDateString();
  const variance = "-"; // Empty until actual cost is added
  const status = "Pending";
  
  sheet.appendRow([
    id,
    itemName,
    category,
    plannedCost,
    "", // Actual cost (empty initially)
    variance,
    status,
    dateCreated,
    notes
  ]);
  
  return {
    success: true,
    message: "Entry added successfully",
    id: id
  };
}

// Update actual cost for an entry
function updateActualCost(id, actualCost) {
  const sheet = initializeSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const plannedCost = data[i][3];
      const variance = actualCost - plannedCost;
      const status = actualCost <= plannedCost ? "Within Budget" : "Over Budget";
      
      sheet.getRange(i + 1, 5).setValue(actualCost);
      sheet.getRange(i + 1, 6).setValue(variance);
      sheet.getRange(i + 1, 7).setValue(status);
      
      return {
        success: true,
        message: "Actual cost updated successfully",
        variance: variance,
        status: status
      };
    }
  }
  
  return {
    success: false,
    message: "Entry not found"
  };
}

// Delete a cost entry
function deleteCostEntry(id) {
  const sheet = initializeSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return {
        success: true,
        message: "Entry deleted successfully"
      };
    }
  }
  
  return {
    success: false,
    message: "Entry not found"
  };
}

// Get summary statistics
function getSummary() {
  const entries = getCostEntries();
  
  let totalPlanned = 0;
  let totalActual = 0;
  let withActualCost = 0;
  
  entries.forEach(entry => {
    if (entry.plannedCost) {
      totalPlanned += parseFloat(entry.plannedCost) || 0;
    }
    if (entry.actualCost) {
      totalActual += parseFloat(entry.actualCost) || 0;
      withActualCost++;
    }
  });
  
  const totalVariance = totalActual - totalPlanned;
  const percentageVariance = totalPlanned > 0 ? ((totalVariance / totalPlanned) * 100).toFixed(2) : 0;
  
  return {
    totalPlanned: totalPlanned.toFixed(2),
    totalActual: totalActual.toFixed(2),
    totalVariance: totalVariance.toFixed(2),
    percentageVariance: percentageVariance,
    entriesWithActual: withActualCost,
    totalEntries: entries.length
  };
}

// Web app entry point
function doGet() {
  return HtmlService.createHtmlTemplateFromFile("index").evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
