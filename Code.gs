/**
 * @OnlyCurrentDoc
 * Spool and Reel Management System
 * Updated: Import now supports TOR.csv and auto-detects Tab vs Comma delimiters.
 * Updated: Auto-add unregistered items to Master Inventory during checkout.
 */

// --- CONFIGURATION & CONSTANTS ---

const SHEET_NAMES = {
  BORROW: "Item History", // Transaction Log
  MASTER: "Inventory"     // Master list
};

// Column Indices for 'Item History' Sheet (0-based)
const COLS_HISTORY = {
  BORROWER: 0,
  ITEM_NO: 1,
  DESC: 2,
  LOCATION: 3,     
  DATE_BORROW: 4,  
  DATE_RETURN: 5   
};

// Column Indices for the 'Inventory' (Master) Sheet (0-based)
const COLS_MASTER = {
  ITEM_NO: 0,      // Col A
  DESC: 1,         // Col B
  LOCATION: 2      // Col C
};

const STATUS = {
  CHECKED_OUT: 'Checked Out',
  AVAILABLE: 'Available'
};

// --- MENU & UI HANDLERS ---

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Spool/Reel Management')
      .addItem('Open Main Menu', 'showMainMenuDialog')
      .addSeparator()
      .addItem('Update Inventory', 'showImportDialog')
      .addToUi();

  SpreadsheetApp.getActiveSpreadsheet().toast("Click 'Spool/Reel Management' > 'Open Main Menu' to start.", "System Ready", 8);
}

function showMainMenuDialog() { createModal('MainMenu', 'Spool & Reel Menu', 700, 500); }
function showBorrowDialog()   { createModal('BorrowDialog', 'Check Out Item', 700, 500); }
function showReturnDialog(itemNo) { 
  const template = HtmlService.createTemplateFromFile('ReturnDialog');
  template.assetId = itemNo || ''; 
  const html = template.evaluate().setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Return Item');
}
function showFindDialog()     { createModal('FindDialog', 'Find Item', 700, 500); }
function showImportDialog()   { createModal('ImportDialog', 'Update Inventory', 700, 400); }

function createModal(filename, title, width, height) {
  const html = HtmlService.createHtmlOutputFromFile(filename)
      .setWidth(width)
      .setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

// --- CORE BUSINESS LOGIC ---

function processBorrowForm(formObject) {
  try {
    const { historySheet, masterSheet } = getSheetsOrThrow();
    const pcName = formObject.pcName;
    const itemNo = formObject.assetId.toUpperCase(); 

    // 1. Validation: Check Strict Availability based on History
    const latestLog = getLatestHistoryEntry(historySheet, itemNo);
    
    if (latestLog && latestLog.returnDate === "") {
      const borrowDateStr = new Date(latestLog.borrowDate).toLocaleDateString();
      return `Error: Item No. '${itemNo}' is currently checked out by '${latestLog.borrower}' (Date: ${borrowDateStr}). Please return it first.`;
    }

    // 2. Find Item in Master to get Description and LOCATION
    const masterData = masterSheet.getDataRange().getValues();
    let masterRowIndex = -1;
    let itemDescription = "";
    let itemLocation = "";

    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][COLS_MASTER.ITEM_NO].toString().toUpperCase() === itemNo) {
        masterRowIndex = i + 1;
        itemDescription = masterData[i][COLS_MASTER.DESC];
        itemLocation = masterData[i][COLS_MASTER.LOCATION]; 
        break;
      }
    }

    if (masterRowIndex === -1) {
      // Auto-add the new item to the Master Inventory sheet
      masterRowIndex = masterSheet.getLastRow() + 1;
      const newMasterRow = [];
      newMasterRow[COLS_MASTER.ITEM_NO] = itemNo;
      newMasterRow[COLS_MASTER.DESC] = "Auto-added during checkout";
      newMasterRow[COLS_MASTER.LOCATION] = "";
      
      masterSheet.getRange(masterRowIndex, 1, 1, newMasterRow.length).setValues([newMasterRow]);
      
      // Update local variables for history logging
      itemDescription = newMasterRow[COLS_MASTER.DESC];
      itemLocation = newMasterRow[COLS_MASTER.LOCATION];
    }

    // Insert into Item History Log (Top of list)
    historySheet.insertRowAfter(1);
    const newRow = [];
    newRow[COLS_HISTORY.BORROWER] = pcName;
    newRow[COLS_HISTORY.ITEM_NO] = itemNo;
    newRow[COLS_HISTORY.DESC] = itemDescription;
    newRow[COLS_HISTORY.LOCATION] = itemLocation; 
    newRow[COLS_HISTORY.DATE_BORROW] = new Date();
    newRow[COLS_HISTORY.DATE_RETURN] = ""; // Explicitly empty means "Checked Out"

    historySheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);

    return `Success: Item '${itemNo}' checked out by '${pcName}'.`;

  } catch (e) {
    return `Error: ${e.toString()}`;
  }
}

function processReturnForm(formObject) {
  try {
    const { historySheet, masterSheet } = getSheetsOrThrow();
    const itemNo = formObject.assetId.toUpperCase();

    // 1. Find the latest open entry in logs
    const historyData = historySheet.getDataRange().getValues();
    let logRowIndex = -1;

    for (let i = 1; i < historyData.length; i++) {
      const row = historyData[i];
      if (row[COLS_HISTORY.ITEM_NO].toString().toUpperCase() === itemNo && row[COLS_HISTORY.DATE_RETURN].toString().trim() === "") {
        logRowIndex = i + 1; 
        break; 
      }
    }

    if (logRowIndex === -1) {
      return `Error: Item No. '${itemNo}' is not currently checked out.`;
    }

    // 2. Find Item in Master
    const masterData = masterSheet.getDataRange().getValues();
    let masterRowIndex = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][COLS_MASTER.ITEM_NO].toString().toUpperCase() === itemNo) {
        masterRowIndex = i + 1;
        break;
      }
    }

    // 3. Execute Updates
    historySheet.getRange(logRowIndex, COLS_HISTORY.DATE_RETURN + 1).setValue(new Date());
    
    return `Success: Item '${itemNo}' has been returned.`;

  } catch (e) {
    return `Error: ${e.toString()}`;
  }
}

function findAsset(formObject) {
  try {
    const { historySheet, masterSheet } = getSheetsOrThrow();
    const itemNo = formObject.assetId.toUpperCase();

    const masterData = masterSheet.getDataRange().getValues();
    let targetItem = null;
    let masterRowIndex = -1;

    // 1. Fetch Static Data from Master
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][COLS_MASTER.ITEM_NO].toString().toUpperCase() === itemNo) {
        targetItem = masterData[i];
        masterRowIndex = i + 1;
        break;
      }
    }

    if (!targetItem) return `Error: Item No. '${itemNo}' not found in Inventory.`;

    // 2. Determine Real Status from History
    const latestLog = getLatestHistoryEntry(historySheet, itemNo);
    
    let realStatus = STATUS.AVAILABLE; 
    let assignedTo = "";
    let displayData = { borrowDate: "", lastBorrower: "", lastReturnDate: "" };

    if (latestLog) {
      displayData.borrowDate = new Date(latestLog.borrowDate).toLocaleDateString();
      if (latestLog.returnDate === "") {
        realStatus = STATUS.CHECKED_OUT;
        assignedTo = latestLog.borrower; 
      } else {
        realStatus = STATUS.AVAILABLE;
        assignedTo = ""; 
        displayData.lastBorrower = latestLog.borrower;
        displayData.lastReturnDate = new Date(latestLog.returnDate).toLocaleDateString();
      }
    }

    // 3. Build Output Message
    let message = `Item No: ${itemNo}\nDescription: ${targetItem[COLS_MASTER.DESC]}\nLocation: ${targetItem[COLS_MASTER.LOCATION]}\nStatus: ${realStatus}`;

    if (realStatus === STATUS.CHECKED_OUT) {
       if (assignedTo) message += `\nAssigned To: ${assignedTo}`;
       if (displayData.borrowDate) message += `\nChecked Out On: ${displayData.borrowDate}`;
    } else if (realStatus === STATUS.AVAILABLE && displayData.lastBorrower) {
       message += `\n\nLast Borrower: ${displayData.lastBorrower}`;
       message += `\nLast Returned: ${displayData.lastReturnDate}`;
    }

    return message;

  } catch (e) {
    return `Error: ${e.toString()}`;
  }
}

function importNewAssets(csvText) {
  try {
    const { masterSheet } = getSheetsOrThrow();

    const masterRange = masterSheet.getDataRange();
    const masterData = masterRange.getValues();
    const existingItemNos = new Map();
    
    // Map existing items to their row index (to allow updating)
    if (masterData.length > 1) {
       for (let i = 1; i < masterData.length; i++) {
         const id = masterData[i][COLS_MASTER.ITEM_NO].toString().toUpperCase().trim();
         if (id) existingItemNos.set(id, i);
       }
    }

    const delimiter = csvText.indexOf('\t') !== -1 ? '\t' : ',';
    const csvData = Utilities.parseCsv(csvText, delimiter); 
    
    if (csvData.length < 2) return "Error: The provided file is empty or lacks data rows.";

    // Dynamic Header Resolution to handle varying CSV structures safely
    const headers = csvData[0].map(h => h.toString().toLowerCase().trim());
    let idxItem = headers.findIndex(h => /item|asset|no|id|spool|reel/i.test(h));
    let idxDesc = headers.findIndex(h => /desc|detail|name/i.test(h));
    let idxLoc  = headers.findIndex(h => /loc|area|place/i.test(h));

    // Smart fallback to legacy indices or standard indices if dynamic mapping fails
    if (idxItem === -1) idxItem = csvData[0].length > 2 ? 2 : 0;
    if (idxDesc === -1) idxDesc = csvData[0].length > 9 ? 9 : 1;
    if (idxLoc === -1)  idxLoc = csvData[0].length > 3 ? 3 : 2;

    const rowsToAdd = [];
    let stats = { added: 0, updated: 0, skipped: 0 };
    let needsMasterUpdate = false;

    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (!row || !row[idxItem]) {
        stats.skipped++;
        continue;
      }
      
      const csvItemNo = row[idxItem].toString().toUpperCase().trim();
      
      if (csvItemNo === "") {
        stats.skipped++;
        continue;
      }
      
      const newDesc = idxDesc !== -1 && row[idxDesc] ? row[idxDesc].trim() : "";
      const newLoc = idxLoc !== -1 && row[idxLoc] ? row[idxLoc].trim() : "";

      if (existingItemNos.has(csvItemNo)) {
        // Check for updates to existing items
        const rowIndex = existingItemNos.get(csvItemNo);
        
        // Skip updates for items that were already added in this same file iteration (mapped to -1)
        if (rowIndex === -1) {
          stats.skipped++;
          continue;
        }

        let updated = false;

        if (newDesc && masterData[rowIndex][COLS_MASTER.DESC] !== newDesc) {
          masterData[rowIndex][COLS_MASTER.DESC] = newDesc;
          updated = true;
        }
        if (newLoc && masterData[rowIndex][COLS_MASTER.LOCATION] !== newLoc) {
          masterData[rowIndex][COLS_MASTER.LOCATION] = newLoc;
          updated = true;
        }

        if (updated) {
          stats.updated++;
          needsMasterUpdate = true;
        } else {
          stats.skipped++;
        }
      } else {
        // Add completely new item
        const newEntry = [];
        newEntry[COLS_MASTER.ITEM_NO] = csvItemNo;
        newEntry[COLS_MASTER.DESC] = newDesc;
        newEntry[COLS_MASTER.LOCATION] = newLoc;

        rowsToAdd.push(newEntry);
        existingItemNos.set(csvItemNo, -1); // Mark as added to prevent duplicates in the same CSV
        stats.added++;
      }
    }

    // 1. Write existing updates back to sheet if any changed
    if (needsMasterUpdate) {
      masterRange.setValues(masterData);
    }

    // 2. Append new rows
    if (rowsToAdd.length > 0) {
      masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    return `Import complete: ${stats.added} added, ${stats.updated} updated, ${stats.skipped} skipped/unchanged.`;

  } catch (e) {
    return `Error: ${e.toString()}`;
  }
}

// --- HELPER FUNCTIONS ---

function getSheetsOrThrow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName(SHEET_NAMES.BORROW);
  let masterSheet = ss.getSheetByName(SHEET_NAMES.MASTER);

  if (!historySheet) throw new Error(`Sheet '${SHEET_NAMES.BORROW}' not found.`);
  
  if (!masterSheet) {
    masterSheet = ss.insertSheet(SHEET_NAMES.MASTER);
    const headers = [['Item No.', 'Description', 'Location']];
    masterSheet.getRange(1, 1, 1, 3).setValues(headers);
    masterSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    masterSheet.setFrozenRows(1);
  }

  return { historySheet, masterSheet };
}

function getLatestHistoryEntry(historySheet, itemNo) {
  const data = historySheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS_HISTORY.ITEM_NO].toString().toUpperCase() === itemNo) {
      return {
        row: i + 1,
        borrower: data[i][COLS_HISTORY.BORROWER],
        borrowDate: data[i][COLS_HISTORY.DATE_BORROW],
        returnDate: data[i][COLS_HISTORY.DATE_RETURN].toString().trim()
      };
    }
  }
  return null;
}

function getAssetIds() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'item_nos';
  const cached = cache.get(cacheKey);
   
  if (cached) return JSON.parse(cached);

  try {
    const { masterSheet } = getSheetsOrThrow();
    if (masterSheet.getLastRow() <= 1) return [];

    const data = masterSheet.getRange(2, COLS_MASTER.ITEM_NO + 1, masterSheet.getLastRow()-1, 1).getValues();
    const itemNos = data.flat().filter(String);
    cache.put(cacheKey, JSON.stringify(itemNos), 600);
    return itemNos;
  } catch (e) {
    return [];
  }
}
