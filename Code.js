/**
 * ==========================================
 * LAD MANAGEMENT SYSTEM - HIGH PERFORMANCE BACKEND
 * ==========================================
 */


/*Test Comment*/

const APP_TITLE = "LAD Management System";

// --- 1. CONFIGURATION & LOGIC DICTIONARY ---
const LOOKUP_MAP = {
  'contact_id':      { targetSheet: 'Contacts', nameColIndex: 2 }, 
  'original_id':     { targetSheet: 'Original', nameColIndex: 1 }, 
  'product_id':      { targetSheet: 'Product',  nameColIndex: 1 }, 
  'product_type_id': { targetSheet: 'Product_Type_Model', nameColIndex: 1 }, 
  'material_id':     { targetSheet: 'Material', nameColIndex: 6 }, 
  'shipping_id':     { targetSheet: 'Packaging', nameColIndex: 10 } 
};

// Unified mapping for all short values (Type, Status, Medium, etc.)
const SYSTEM_MAPPINGS = {
  "type": {
    "Original": "OR", "Print": "PR", "Skulptur": "SC", "Sculpture": "SC",
    "Foto": "PH", "Photo": "PH", 
    "Digital": "DG", "Digital Art": "DG"
  },
  "status": {
    "Available": "AV", "Sold": "SD", "Reserved": "RS"
  },
  "medium": {
    "Reproduction on Paper": "RP",
    "Reproduction on Wood": "RW",
    "Reproduction on Glas": "RG"
    // Add your other mediums here
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi(); // <--- Line 54 or above: Ensure this is defined!

  // MENU 1: DATA MANAGEMENT (The "Backend" stuff)
  ui.createMenu('🗄️ LAD Data')
    .addItem('👥 Contacts', 'openContacts')
    .addItem('🎨 Originals', 'openOriginal')
    .addItem('🖼️ Products', 'openProduct')
    .addSeparator()
    .addItem('1. Update Lookups [Sheet.Col]', 'updateHeaderComputedColumns')
    .addItem('2. Update Short Values (_short_value)', 'updateAllComputedValues')
    .addItem('3. Build All SKUs', 'updateAllSkuCodes')
    .addSeparator()
    .addItem('📦 Update Product Unit List', 'syncProductInventory')
    .addToUi();

  // MENU 2: ORDERS & BILLING (The "Daily" stuff)
ui.createMenu('🛒 LAD Orders')
    .addItem('➕ New Customer Order', 'openOrderModal') // <--- Must match function name below
    .addItem('📄 Generate PDF (Invoice/Cert)', 'generatePdfForSelectedRow')
    .addToUi();
}



function openContacts() { activateAndOpen('Contacts'); }
function openOriginal() { activateAndOpen('Original'); }
function openProduct()  { activateAndOpen('Product'); }

function activateAndOpen(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.getSheets().find(s => s.getName().toLowerCase().includes(sheetName.toLowerCase()));
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet not found: " + sheetName);
    return;
  }
  ss.setActiveSheet(sheet);
  showSidebar();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setWidth(1000).setHeight(700).setTitle(APP_TITLE);
  SpreadsheetApp.getUi().showModalDialog(html, APP_TITLE);
}

/**
 * Launcher for the Order Modal. 
 * This is the only part of Code.js that needs to change.
 */
function showOrderModal() {
  const html = HtmlService.createHtmlOutputFromFile('OrderModal')
      .setWidth(800)
      .setHeight(600)
      .setTitle('Create New Order');
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Order');
}

function openOrderModal() {
  const html = HtmlService.createHtmlOutputFromFile('OrderModal')
      .setWidth(950)
      .setHeight(750)
      .setTitle('🛒 Create New Order');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * ==========================================
 * SECTION 3: SIDEBAR BACKEND (CRUD)
 * ==========================================
 */

function getSheetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastCol === 0) return { error: "Sheet is empty" };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // BUILD COLUMNS
    const columns = [];
    headers.forEach((header, index) => {
      const hStr = String(header).trim();
      
      // LOGIC: Hide Computed Columns [Sheet.Col] from UI
      if (hStr.startsWith('[') && hStr.endsWith(']')) return;

      const colObj = { name: hStr, index: index, options: [] };
      const headerKey = hStr.toLowerCase();

      // LOGIC: Foreign Key Lookup
      const isSelfId = (headerKey === sheetName.toLowerCase().replace(/s$/, "") + "_id") || 
                       (headerKey === sheetName.toLowerCase() + "_id") ||
                       (sheetName === 'Contacts' && headerKey === 'contact_id') ||
                       (sheetName === 'Original' && headerKey === 'original_id');

      if (LOOKUP_MAP[headerKey] && !isSelfId) {
        colObj.options = fetchLinkedOptions(ss, LOOKUP_MAP[headerKey]);
      } 
      else if (lastRow > 1) {
        const rule = sheet.getRange(2, index + 1).getDataValidation();
        if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          colObj.options = rule.getCriteriaValues()[0];
        }
      }
      columns.push(colObj);
    });

    // FETCH DATA
    let data = [];
    if (lastRow > 1) {
      const rawData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      data = rawData.map(row => {
        const uiRow = columns.map(col => {
          let val = row[col.index];
          if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
          return val;
        });
        if(uiRow.length > 0) uiRow[0] = String(uiRow[0]);
        return uiRow;
      });
    }

    return { sheetName: sheetName, columns: columns, data: data };
  } catch(e) { return { error: e.message }; }
}

function fetchLinkedOptions(ss, config) {
  let targetSheet = ss.getSheetByName(config.targetSheet);
  if (!targetSheet) targetSheet = ss.getSheets().find(s => s.getName().includes(config.targetSheet));
  if (!targetSheet) return [];

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return [];

  const ids = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const names = targetSheet.getRange(2, config.nameColIndex + 1, lastRow - 1, 1).getValues();
  
  const useFallback = (config.targetSheet === 'Contacts');
  let firstNames, lastNames;
  if(useFallback) {
     firstNames = targetSheet.getRange(2, 4, lastRow - 1, 1).getValues(); 
     lastNames = targetSheet.getRange(2, 6, lastRow - 1, 1).getValues();
  }

  const options = [];
  for (let i = 0; i < ids.length; i++) {
    let id = String(ids[i][0]);
    if (!id) continue;
    let label = names[i][0];
    if (useFallback && (!label || label === "")) label = `${firstNames[i][0]} ${lastNames[i][0]}`.trim();
    if (!label) label = id;
    options.push({ id: id, name: label });
  }
  return options.sort((a, b) => a.name.localeCompare(b.name));
}

// --- OPTIMIZED SAVE FUNCTION (BATCH WRITE) ---
function processForm(formObject) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { error: "System busy. Please try again." };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = formObject.sheetName;
    let sheet = sheetName ? ss.getSheetByName(sheetName) : SpreadsheetApp.getActiveSheet();
    if (!sheet) return { error: "Sheet not found" };

    const lastCol = sheet.getLastColumn();
    // 1. Get Headers to Map Indices
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // 2. Identify UI Columns (Skip computed)
    const uiColMap = {}; // Maps "col-0" -> realIndex
    headers.forEach((h, i) => {
      const hStr = String(h).trim();
      if (!(hStr.startsWith('[') && hStr.endsWith(']'))) {
        // Find which UI index matches this real index. 
        // We re-calculate because we don't pass the map from UI.
        // Simple strategy: The UI sends "col-0", "col-1"... 
        // We simply iterate the Object.keys starting with 'col-' to be safe, 
        // OR we rely on the order. 
        // BETTER STRATEGY: Rerun the filter logic quickly.
      }
    });
    
    // Re-run the column logic to get exact UI indices
    const uiCols = [];
    headers.forEach((header, index) => {
      const hStr = String(header).trim();
      if (!(hStr.startsWith('[') && hStr.endsWith(']'))) {
        uiCols.push({ realIndex: index, uiIndex: uiCols.length });
      }
    });

    const formId = formObject['col-0'] ? String(formObject['col-0']).trim() : "";
    let isNew = (formId === "");
    let rowToUpdate = -1;
    let currentValues = [];

    // 3. Find Row & Fetch Current Data (1 Read)
    if (!isNew) {
      const allIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(id => String(id).trim());
      rowToUpdate = allIds.indexOf(formId);
      if (rowToUpdate === -1) return { error: "ID Not Found. It may have been deleted." };
      rowToUpdate += 2; // Adjust for header + 0-index
      
      // Fetch the whole row to preserve existing data in non-UI columns
      currentValues = sheet.getRange(rowToUpdate, 1, 1, lastCol).getValues()[0];
    } else {
      rowToUpdate = sheet.getLastRow() + 1;
      currentValues = new Array(lastCol).fill(""); // Empty row
    }

    const idToSave = isNew ? Utilities.getUuid() : formId;

    // 4. Update Values in Memory
    uiCols.forEach((col, i) => {
      let val = formObject['col-' + i]; // 'col-0', 'col-1' matches the UI order
      if (col.realIndex === 0) val = idToSave;
      currentValues[col.realIndex] = val;
    });

    // 5. Batch Write (1 Write)
    sheet.getRange(rowToUpdate, 1, 1, lastCol).setValues([currentValues]);

    // 6. Run Automation (Targeted)
    runTargetedTriggers(sheet, rowToUpdate, idToSave);

    return { message: isNew ? "Created Successfully" : "Updated Successfully" };

  } catch (e) {
    return { error: "Save Error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteItem(id) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const allIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
  const index = allIds.indexOf(String(id));
  if (index > -1) {
    sheet.deleteRow(index + 2);
    return { message: "Deleted" };
  }
  return { error: "ID Not Found" };
}

/**
 * ==========================================
 * SECTION 4: AUTOMATION TRIGGERS (OPTIMIZED)
 * ==========================================
 */

function forceRunAllTriggers() {
  updateHeaderComputedColumns();
  updateAllSkuCodes();
  updateAllComputedValues();
  SpreadsheetApp.getActive().toast("Full sync complete.");
}

// TARGETED UPDATE: Only processes the current sheet/row if possible
function runTargetedTriggers(sheet, rowIndex, rowId) {
  // Update SKUs for this sheet only
  processSkuUpdate(sheet.getParent(), sheet.getName(), rowIndex);
  
  // Update Linked Columns (Incoming) for this sheet only
  processHeaderComputedColumns(sheet.getParent(), sheet, rowIndex);
  
  // Update Short IDs (Contacts)
  if (sheet.getName() === "Contacts") updateAllComputedValues();
}

// A. LINKED COLUMNS [Sheet.Column]
function updateHeaderComputedColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => processHeaderComputedColumns(ss, sheet));
}

function processHeaderComputedColumns(ss, sheet, targetRowIndex = null) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = headers.map(h => String(h).toLowerCase());

  // 1. Identify which columns have the [Sheet.Column] syntax
  const computed = [];
  headers.forEach((h, i) => {
    const match = String(h).match(/^\[([\w\s]+)\.([\w\s_]+)\]$/);
    if (match) {
      computed.push({ 
        colIndex: i, 
        targetSheetName: match[1], 
        targetFieldName: match[2] 
      });
    }
  });

  if (computed.length === 0) return;

  // 2. Prepare Data Range (Single Row or All)
  let startRow = 2;
  let numRows = lastRow - 1;
  let data = targetRowIndex 
    ? sheet.getRange(targetRowIndex, 1, 1, sheet.getLastColumn()).getValues()
    : sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();

  const cache = {};
  let updated = false;

  computed.forEach(comp => {
    // A. Use the LOOKUP_MAP to find the required Key for this target sheet
    // Example: Find the entry where targetSheet is "Contacts" -> Key is "contact_id"
    const lookupEntry = Object.entries(LOOKUP_MAP).find(([key, val]) => 
      val.targetSheet.toLowerCase() === comp.targetSheetName.toLowerCase()
    );

    if (!lookupEntry) return; // Skip if target sheet isn't in our dictionary

    const keyColumnName = lookupEntry[0]; // e.g., "contact_id"
    const linkIdx = headerMap.findIndex(h => 
  h === keyColumnName.toLowerCase() || h.endsWith('.' + keyColumnName.toLowerCase() + ']')
);

    if (linkIdx === -1) return; // The current sheet is missing the required ID column

    // B. Lazy-load the target sheet data into a Map for high performance
    const tKey = comp.targetSheetName.toLowerCase();
    if (!cache[tKey]) {
      const tSheet = ss.getSheets().find(s => s.getName().toLowerCase() === tKey);
      if (tSheet) {
        const tData = tSheet.getDataRange().getValues();
        const tMap = new Map();
        if (tData.length > 1) {
          tData.slice(1).forEach(r => tMap.set(String(r[0]), r));
          cache[tKey] = { 
            map: tMap, 
            headers: tData[0].map(x => String(x).toLowerCase()) 
          };
        }
      }
    }

    const tCache = cache[tKey];
    if (!tCache) return;

    const targetColIdx = tCache.headers.indexOf(comp.targetFieldName.toLowerCase());
    if (targetColIdx === -1) return;

    // C. Perform the Dynamic Lookup
    for (let i = 0; i < data.length; i++) {
      const linkId = String(data[i][linkIdx]);
      if (linkId && tCache.map.has(linkId)) {
        const newVal = tCache.map.get(linkId)[targetColIdx];
        if (String(data[i][comp.colIndex]) !== String(newVal)) {
          data[i][comp.colIndex] = newVal;
          updated = true;
        }
      }
    }
  });

  if (updated) {
    const writeRow = targetRowIndex || 2;
    sheet.getRange(writeRow, 1, data.length, data[0].length).setValues(data);
  }
}

// B. SKU CODES
function updateAllSkuCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Original', 'Product', 'Product_Units'].forEach(name => processSkuUpdate(ss, name));
}



function processSkuUpdate(ss, sheetName, targetRowIndex = null) {
  let sheet = null;
  if (typeof sheetName === 'string') {
     sheet = ss.getSheetByName(sheetName);
     if (!sheet) sheet = ss.getSheets().find(s => s.getName().includes(sheetName));
  } else {
     sheet = sheetName; 
  }
  
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // --- NEW: FETCH SYNTAX FROM CENTRAL SHEET ---
  const centralSyntax = getCentralSkuSyntax(ss, sheet.getName());
  if (!centralSyntax) return; // Skip if this sheet is not configured in "SKU Syntax"

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  
  // Find the column to write the SKU into (e.g., 'sku_code')
  const skuColName = centralSyntax.targetCol.toLowerCase();
  const skuIdx = headers.indexOf(skuColName);
  
  if (skuIdx === -1) return;

  // Determine Range
  let startRow = 2;
  let numRows = lastRow - 1;
  let data = [];

  if (targetRowIndex) {
    startRow = targetRowIndex;
    numRows = 1;
    data = sheet.getRange(startRow, 1, 1, sheet.getLastColumn()).getValues();
  } else {
    data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
  }

  const cache = {}; 
  let updated = false;

  for (let i = 0; i < data.length; i++) {
    const syntax = centralSyntax.pattern;
    
    // We load cache only if we actually need to parse a syntax
    if(Object.keys(cache).length === 0) {
       ss.getSheets().forEach(s => {
         const d = s.getDataRange().getValues();
         if(d.length > 1) cache[s.getName().toLowerCase()] = { headers: d[0].map(x=>String(x).toLowerCase()), data: d };
       });
    }

    const newCode = parseSkuSyntax(syntax, data[i], headers, cache, sheet.getName().toLowerCase());
    if (newCode !== data[i][skuIdx]) {
      data[i][skuIdx] = newCode;
      updated = true;
    }
  }

  if (updated) {
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * Retrieves SKU configuration from the "SKU Syntax" sheet.
 * Col 1: Sheet Name | Col 2: Column to Write To | Col 3: Syntax Pattern
 */
function getCentralSkuSyntax(ss, sheetName) {
  const syntaxSheet = ss.getSheetByName("SKU Syntax");
  if (!syntaxSheet) return null;
  
  const data = syntaxSheet.getDataRange().getValues();
  // Skip header, look for matching sheet name in Col 1
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === sheetName.toLowerCase()) {
      return {
        targetCol: String(data[i][1]).trim(),
        pattern: String(data[i][2]).trim()
      };
    }
  }
  return null;
}

function parseSkuSyntax(syntax, row, headers, cache, currentSheet) {
  return syntax.replace(/\[([\w\s]+)\.([\w\s_]+)\]/g, (match, tSheet, tCol) => {
    tSheet = tSheet.toLowerCase();
    tCol = tCol.toLowerCase();
    
    if (tSheet === currentSheet || tSheet === 'product_units' && currentSheet === 'inventar') {
      const idx = headers.indexOf(tCol);
      return idx > -1 ? row[idx] : "";
    }
    
    if (cache[tSheet]) {
      let linkColName = tSheet + "_id";
      if(tSheet === 'contacts') linkColName = 'contact_id';
   const linkIdx = headers.findIndex(h => 
  h === linkColName || h.endsWith('.' + linkColName + ']')
);
      if(linkIdx === -1) return "";
      
      const linkId = String(row[linkIdx]);
      const targetRow = cache[tSheet].data.find(r => String(r[0]) === linkId);
      if(targetRow) {
         const tIdx = cache[tSheet].headers.indexOf(tCol);
         return tIdx > -1 ? targetRow[tIdx] : "";
      }
    }
    return "";
  }).replace(/&""-""&/g, "-").replace(/"/g, ""); 
}

/**
 * UNIFIED AUTOMATION: Handles Dictionary Mappings, Size Logic, and Contact IDs
 */


function updateAllComputedValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const range = sheet.getDataRange();
    const data = range.getValues();
    if (data.length < 2) return; 
    
    const headers = data[0].map(h => String(h).toLowerCase());
    let updated = false;

    headers.forEach((header, colIdx) => {
      if (header.endsWith('_short_value')) {
        const sourceBase = header.replace('_short_value', '').trim();
        const sourceIdx = headers.indexOf(sourceBase);
        
        if (sourceIdx > -1) {
          for (let i = 1; i < data.length; i++) {
            let sourceVal = String(data[i][sourceIdx]).trim();
            if (!sourceVal) continue;

            let newVal = "";

            // A. Check Unified Global Mapping (Point 1: Handles type, status, AND medium)
            if (SYSTEM_MAPPINGS[sourceBase] && SYSTEM_MAPPINGS[sourceBase][sourceVal]) {
              newVal = SYSTEM_MAPPINGS[sourceBase][sourceVal];
            } 
            // B. SIZE LOGIC: Improved to handle original_size (Point 2)
            else if (sourceBase === 'size') {
              newVal = sourceVal.toLowerCase().replace(/x/g, "").replace(/cm/g, "").replace(/\s/g, "").trim();
            }
            // C. Fallback: Standard 3-letter shortening
            else {
              newVal = "Error review columns names";
            }

            if (data[i][colIdx] !== newVal) {
              data[i][colIdx] = newVal;
              updated = true;
            }
          }
        }
      }
      // ... (Rest of contact_short_id logic remains the same)
    });

    if (updated) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }
  });
  ss.toast("Short Values (including Medium) and Sizes updated.");
}


// --- UTILS & PDF ---
function generatePdfForSelectedRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (row < 2) { ss.toast("⚠️ Select a valid row."); return; }
  
  const templateId = getSetting("PDF_TEMPLATE_ID");
  const folderId = getSetting("PDF_FOLDER_ID");
  if (!templateId || !folderId) { SpreadsheetApp.getUi().alert("⚠️ Config Missing in Settings sheet."); return; }

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  
  const mergeData = {};
  headers.forEach((h, i) => {
    let key = h.toString().toLowerCase().trim();
    let val = values[i];
    if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
    mergeData[key] = val;
  });

  const uuid = values[0]; 
  if (uuid) {
    try {
      const qrUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodeURIComponent(sheet.getName()+":"+uuid)}`;
      mergeData['qr_code'] = UrlFetchApp.fetch(qrUrl).getBlob();
    } catch (e) { console.log(e); }
  }

  try {
    const fileName = (mergeData['sku_code'] || "File_" + row) + ".pdf";
    const pdfUrl = createPdfFromTemplate(templateId, folderId, fileName, mergeData);
    
    let targetColIdx = headers.findIndex(h => h.toLowerCase().includes('sheet') || h.toLowerCase().includes('link'));
    if (targetColIdx > -1) sheet.getRange(row, targetColIdx + 1).setValue(pdfUrl);
    
    ss.toast("✅ PDF Created!");
  } catch (e) { SpreadsheetApp.getUi().alert("Error: " + e.message); }
}

function createPdfFromTemplate(templateId, folderId, fileName, data) {
  const templateFile = DriveApp.getFileById(templateId);
  const parentFolder = DriveApp.getFolderById(folderId);
  const tempFile = templateFile.makeCopy("TEMP_" + fileName, parentFolder);
  const tempDoc = DocumentApp.openById(tempFile.getId());
  const body = tempDoc.getBody();
  
  for (const [key, value] of Object.entries(data)) {
    if (value && value.toString() === 'Blob') {
       const el = body.findText(`{{${key}}}`);
       if (el) {
         const t = el.getElement().asText();
         t.setText(""); 
         t.getParent().asParagraph().insertInlineImage(0, value);
       }
       continue; 
    }
    body.replaceText(`{{${key}}}`, String(value));
  }
  
  tempDoc.saveAndClose();
  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = parentFolder.createFile(pdfBlob).setName(fileName);
  tempFile.setTrashed(true);
  return pdfFile.getUrl();
}

function syncProductInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prodSheet = ss.getSheetByName("Product");
  let unitSheet = ss.getSheetByName("Product_Units") || ss.getSheetByName("Inventar");
  if (!prodSheet || !unitSheet) return;

  const prodData = prodSheet.getDataRange().getValues();
  const prodHeaders = prodData[0].map(h => h.toString().toLowerCase());
  const pIdIdx = prodHeaders.indexOf('product_id');
  const pLimitIdx = prodHeaders.indexOf('limited_edition');

  const unitData = unitSheet.getDataRange().getValues();
  const unitHeaders = unitData[0].map(h => h.toString().toLowerCase());
  
  const getHeaderIdx = (name) => {
    let idx = unitHeaders.indexOf(name);
    if(idx === -1) { 
      unitSheet.getRange(1, unitHeaders.length + 1).setValue(name);
      unitHeaders.push(name);
      idx = unitHeaders.length - 1;
    }
    return idx;
  };

  const uProdIdIdx = getHeaderIdx('product_id');
  const uEdIdx = getHeaderIdx('edition');
  const uStatusIdx = getHeaderIdx('status');
  const uIdIdx = getHeaderIdx('inventory_id');

  const unitMap = {}; 
  for (let i = 1; i < unitData.length; i++) {
    const pid = unitData[i][uProdIdIdx];
    if (!unitMap[pid]) unitMap[pid] = [];
    unitMap[pid].push({ rowIndex: i + 1, edition: unitData[i][uEdIdx], status: unitData[i][uStatusIdx] });
  }

  let created = 0;
  for (let i = 1; i < prodData.length; i++) {
    const pid = prodData[i][pIdIdx];
    const limitRaw = prodData[i][pLimitIdx];
    const limitNum = parseInt(String(limitRaw).replace(/[^0-9]/g, ''), 10);

    if (pid && !isNaN(limitNum) && limitNum > 0) {
      const existing = unitMap[pid] || [];
      const diff = limitNum - existing.length;
      if (diff > 0) {
        let maxEd = 0;
        existing.forEach(u => {
           const n = parseInt(String(u.edition).replace(/[^0-9]/g, ''), 10);
           if(!isNaN(n) && n > maxEd) maxEd = n;
        });

        const newRows = [];
        for (let k = 1; k <= diff; k++) {
          const nextNum = maxEd + k;
          const suffix = "N" + String(nextNum).padStart(3, '0');
          const rowArr = new Array(unitHeaders.length).fill("");
          rowArr[uIdIdx] = Utilities.getUuid();
          rowArr[uProdIdIdx] = pid;
          rowArr[uEdIdx] = suffix;
          rowArr[uStatusIdx] = "Draft";
          newRows.push(rowArr);
        }
        unitSheet.getRange(unitSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        created += diff;
      }
    }
  }
  updateAllSkuCodes(); 
  SpreadsheetApp.getUi().alert(`Sync Complete. Created ${created} new units.`);
}

function getSetting(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return data[i][1];
  }
  return null;
}