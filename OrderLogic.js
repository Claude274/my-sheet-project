/**
 * SERVER-SIDE: Order & Accounting Bridge
 */

const ACCOUNTING_SHEET_ID = '1XvCu3_1-89zdbL5NKEpnQsPJCoVyNsLqszgPA6Xz8EM'; 

const ACC_TEMPLATES = {
  'Cash Sale': {
    debit: { pcn: '512', name: 'Bank', ecdfBS: 'D.IV', ecdfPL: '' },
    credit: { pcn: '701', name: 'Sales - Services', ecdfBS: '', ecdfPL: 'B.1' }
  },
  'Sales Invoice - Goods': {
    debit: { pcn: '401', name: 'Trade Receivables', ecdfBS: 'D.II.1', ecdfPL: '' },
    credit: { pcn: '702', name: 'Sales - Goods', ecdfBS: '', ecdfPL: 'B.1' }
  }
};


function postToAccounting(orderData, orderId) {
  try {
    const accSS = SpreadsheetApp.openById(ACCOUNTING_SHEET_ID);
    const accSheet = accSS.getSheetByName('Transactions'); // DO NOT DELETE THIS SHEET
    
    // Treat as Cash Sale to impact Bank and PL
    const type = 'Cash Sale'; 
    const template = ACC_TEMPLATES[type];
    const transId = "TRX-" + Math.random().toString(36).substring(2, 9).toUpperCase();

    accSheet.appendRow([
      transId, new Date(), type, "Order " + orderId, orderData.contact_name, orderData.total,
      template.debit.pcn, template.debit.name, template.debit.ecdfBS, template.debit.ecdfPL, // Bank
      template.credit.pcn, template.credit.name, template.credit.ecdfBS, template.credit.ecdfPL, // PL (Sales)
      0, 'One Time', '', 'Sales Event', 'Sync from Inventory'
    ]);
    return transId;
  } catch (e) {
    return "SYNC_FAILED";
  }
}

/**
 * Fetches contacts for the Order Modal dropdown.
 */
function getContactsForOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contacts");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase());
  const idIdx = headers.indexOf("contact_id");
  // Matches logic in Code.js for finding names
  const fNameIdx = 3; // Column D
  const lNameIdx = 5; // Column F

  return data.slice(1).map(row => {
    const name = `${row[fNameIdx]} ${row[lNameIdx]}`.trim();
    return {
      id: String(row[idIdx]),
      name: name || String(row[idIdx])
    };
  }).filter(c => c.id).sort((a,b) => a.name.localeCompare(b.name));
}

/**
 * Fetches available inventory (Product_Units) for the Order Modal grid.
 */
function getInventoryForOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Product_Units");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase());
  
  const idIdx = headers.indexOf("inventory_id");
  const skuIdx = headers.indexOf("sku_code");
  const statusIdx = headers.indexOf("status");
  // Assumes a price column exists or a computed [Product.price] column
  const priceIdx = headers.findIndex(h => h.includes("price"));

  return data.slice(1)
    .filter(row => row[statusIdx] !== "Sold") // Only show available items
    .map(row => ({
      id: String(row[idIdx]),
      sku: row[skuIdx] || "No SKU",
      price: row[priceIdx] || 0
    }))
    .filter(item => item.id);
}

function processSaveOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date();
  // Generate ID following the rule: orders_id
  const orderId = "ORD-" + Math.random().toString(36).substring(2, 9).toUpperCase();

  try {
    const accRefId = postToAccounting(orderData, orderId);

    const orderSheet = ss.getSheetByName("Orders");
    if (orderSheet) {
      // Column 1 is now orders_id
      orderSheet.appendRow([
        orderId, 
        timestamp, 
        orderData.contact_id, 
        orderData.contact_name,
        orderData.total, 
        orderData.itemIds.length,
        orderData.skuList, 
        orderData.comment,
        orderData.payment_method,
        "Auto-synced",
        accRefId 
      ]);
    }

    // Update status to Sold in Product_Units
    const invSheet = ss.getSheetByName("Product_Units");
    const invData = invSheet.getDataRange().getValues();
    const headers = invData[0].map(h => String(h).toLowerCase());
    const idIdx = headers.indexOf("product_units_id") === -1 ? headers.indexOf("inventory_id") : headers.indexOf("product_units_id");
    const statusIdx = headers.indexOf("status");

    orderData.itemIds.forEach(itemId => {
      const rowIdx = invData.findIndex(r => String(r[idIdx]) === String(itemId));
      if (rowIdx > -1) invSheet.getRange(rowIdx + 1, statusIdx + 1).setValue("Sold");
    });

    return { success: true, id: orderId };
  } catch (err) {
    throw new Error("Order failed: " + err.message);
  }
}