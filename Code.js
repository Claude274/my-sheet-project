/**
 * ==========================================
 * LAD MANAGEMENT SYSTEM - HIGH PERFORMANCE BACKEND
 * ==========================================
 */


/*Test Comment*/

const APP_TITLE = "LAD Management System";

// --- 1. CONFIGURATION ---
const LOOKUP_MAP = {
  'contact_id':      { targetSheet: 'Contacts', nameColIndex: 2 }, 
  'original_id':     { targetSheet: 'Original', nameColIndex: 1 }, 
  'product_id':      { targetSheet: 'Product',  nameColIndex: 1 }, 
  'product_type_id': { targetSheet: 'Product_Type_Model', nameColIndex: 1 }, 
  'material_id':     { targetSheet: 'Material', nameColIndex: 6 }, 
  'shipping_id':     { targetSheet: 'Packaging', nameColIndex: 10 } 
};

const TYPE_MAPPING = {
  "Original": "OR", "Print": "PR", "Skulptur": "SC", "Sculpture": "SC","Reproduction on Paper" : "RP",
  "Foto": "PH", "Photo": "PH", "Digital": "DG", "Digital Art": "DG"
};

// --- 2. MENU & NAVIGATION ---
function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_TITLE)
    .addItem('👥 Update Contacts', 'openContacts')
    .addItem('🎨 Update Original', 'openOriginal')
    .addItem('🖼️ Update Product', 'openProduct')
    .addSeparator()
    .addItem('📄 Generate PDF for Selected Row', 'generatePdfForSelectedRow')
    .addItem('📦 Sync Product Inventory', 'syncProductInventory')
    .addItem('🛒 Create New Order (POS)', 'openOrderModal')
    .addSeparator()
    .addItem('🔄 Force Update Links & SKUs', 'forceRunAllTriggers')
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

function openOrderModal() {
  try {
    const html = HtmlService.createTemplateFromFile('OrderModal')
      .evaluate().setTitle('🛒 New Customer Order').setWidth(900).setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, '🛒 New Customer Order');
  } catch(e) {
    SpreadsheetApp.getUi().alert("OrderModal file missing or POS not configured.");
  }
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
  
  const computed = [];
  headers.forEach((h, i) => {
    const match = String(h).match(/^\[([\w\s]+)\.([\w\s_]+)\]$/);
    if (match) computed.push({ col: i, targetSheet: match[1], targetField: match[2] });
  });

  if (computed.length === 0) return;

  // If targeting 1 row, only read that row. Else read all.
  let startRow = 2;
  let numRows = lastRow - 1;
  let data = [];
  
  if (targetRowIndex) {
    startRow = targetRowIndex;
    numRows = 1;
    data = sheet.getRange(startRow, 1, 1, sheet.getLastColumn()).getValues();
  } else {
    data = sheet.getDataRange().getValues().slice(1);
  }

  const cache = {};
  let updated = false;

  computed.forEach(comp => {
    const tKey = comp.targetSheet.toLowerCase();
    if (!cache[tKey]) {
      const tSheet = ss.getSheets().find(s => s.getName().toLowerCase() === tKey || s.getName().toLowerCase().includes(tKey));
      if (tSheet) {
        const tData = tSheet.getDataRange().getValues();
        const tMap = new Map();
        if(tData.length > 1) {
          tData.slice(1).forEach(r => tMap.set(String(r[0]), r));
          cache[tKey] = { map: tMap, headers: tData[0].map(x => String(x).toLowerCase()) };
        }
      }
    }
    
    const tCache = cache[tKey];
    if (!tCache) return;

    let linkIdx = headers.findIndex(h => String(h).toLowerCase().includes(comp.targetSheet.toLowerCase()) && String(h).toLowerCase().includes('id'));
    if (linkIdx === -1) linkIdx = headers.findIndex(h => String(h).toLowerCase() === 'original_id');
    
    if (linkIdx === -1) return;

    const targetColIdx = tCache.headers.indexOf(comp.targetField.toLowerCase());
    if (targetColIdx === -1) return;

    for (let i = 0; i < data.length; i++) {
      const linkId = String(data[i][linkIdx]);
      if (linkId && tCache.map.has(linkId)) {
        const newVal = tCache.map.get(linkId)[targetColIdx];
        if (String(data[i][comp.col]) !== String(newVal)) {
          data[i][comp.col] = newVal;
          updated = true;
        }
      }
    }
  });

  if (updated) {
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
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
     sheet = sheetName; // Handle object passed
  }
  
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const skuIdx = headers.indexOf('sku_code');
  const syntaxIdx = headers.indexOf('sku_syntax');
  
  if (skuIdx === -1 || syntaxIdx === -1) return;

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

  const cache = {}; // Lazy load inside parsing if needed

  let updated = false;
  for (let i = 0; i < data.length; i++) {
    const syntax = data[i][syntaxIdx];
    // Recalculate if syntax exists
    if (syntax) {
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
  }

  if (updated) {
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  }
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
      const linkIdx = headers.indexOf(linkColName);
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

// C. SHORT IDS
function updateAllComputedValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Contacts');
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h=>String(h).toLowerCase());
  const shortIdx = headers.indexOf('contact_short_id');
  
  if(shortIdx > -1) {
    let updated = false;
    for(let i=1; i<data.length; i++) {
       if(!data[i][shortIdx]) {
          const fn = data[i][headers.indexOf('first_name')] || "";
          const ln = data[i][headers.indexOf('last_name')] || "";
          if(fn && ln) {
             data[i][shortIdx] = (fn.substring(0,2) + ln.substring(0,2)).toUpperCase();
             updated = true;
          }
       }
    }
    if(updated) sheet.getRange(1,1,data.length, data[0].length).setValues(data);
  }
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