/**
 * SERVER-SIDE: Modern Order Logic
 */

// 1. Fetch Produced Inventory
function getInventoryForOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Product_Units");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  
  const statusIdx = headers.indexOf("status");
  const skuIdx = headers.indexOf("sku_code");
  const priceIdx = headers.indexOf("[product.sell_price]"); 
  const idIdx = headers.indexOf("inventory_id");

  return data.slice(1)
    .filter(row => String(row[statusIdx]).trim() === "Produced")
    .map(row => ({
      id: row[idIdx],
      sku: row[skuIdx],
      price: row[priceIdx] || 0
    }));
}

// 2. Fetch Customers
function getContactsForOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contacts");
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(row => row[1] === "Customer") // Column 2 is 'type'
    .map(row => ({ id: row[0], name: row[2] })); // ID and Artist/Company Name
}

// 3. Save Order & Log Transactions
function processSaveOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date();
  
  // Requirement: Use [sheet_name]_id pattern
  const orderId = "orders_" + Math.random().toString(36).substring(2, 11);
  const transId = "transactions_" + Math.random().toString(36).substring(2, 11);

  // A. Log in Orders Sheet (Includes SKU list and Comment)
  const orderSheet = ss.getSheetByName("Orders");
  if (orderSheet) {
    orderSheet.appendRow([
      orderId, timestamp, orderData.contact_id, orderData.contact_name,
      orderData.total, orderData.itemIds.length, orderData.skuList, orderData.comment
    ]);
  }

  // B. Log Financial Transaction
  const transSheet = ss.getSheetByName("Transactions");
  if (transSheet) {
    transSheet.appendRow([
      transId, timestamp, "Income", "Sale: " + orderId, orderData.contact_name, orderData.total
    ]);
  }

  // C. Update Inventory to 'Sold'
  const invSheet = ss.getSheetByName("Product_Units");
  const invData = invSheet.getDataRange().getValues();
  const idIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("inventory_id");
  const statusIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("status");

  orderData.itemIds.forEach(itemId => {
    const rowIdx = invData.findIndex(r => r[idIdx] === itemId);
    if (rowIdx > -1) invSheet.getRange(rowIdx + 1, statusIdx + 1).setValue("Sold");
  });

  return { success: true };
}


const ACCOUNTING_SHEET_ID = '1XvCu3_1-89zdbL5NKEpnQsPJCoVyNsLqszgPA6Xz8EM';

// We only need the Sales-related templates here to "push" them
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

function postToAccounting(orderData) {
  try {
    const accSS = SpreadsheetApp.openById(ACCOUNTING_SHEET_ID);
    const accSheet = accSS.getSheetByName('Transactions');
    
    // Choose template: Cash vs Bank/Invoice
    const type = (orderData.payment_method === 'Cash') ? 'Cash Sale' : 'Sales Invoice - Goods';
    const template = ACC_TEMPLATES[type];
    
    // Generate Financial ID [sheet_name]_id
    const transId = "transactions_" + Math.random().toString(36).substring(2, 11);
    const timestamp = new Date();

    // The order of these columns must match your Accounting Transactions sheet exactly
    const row = [
      transId,               // ID
      timestamp,             // Date
      type,                  // Type
      "Order: " + orderData.skuList, // Description
      orderData.contact_name, // Supplier/Customer
      orderData.total,       // Amount (€)
      template.debit.pcn,    // DR PCN
      template.debit.name,   // DR Account
      template.debit.ecdfBS, // DR eCDF BS
      template.debit.ecdfPL, // DR eCDF P&L
      template.credit.pcn,   // CR PCN
      template.credit.name,  // CR Account
      template.credit.ecdfBS,// CR eCDF BS
      template.credit.ecdfPL,// CR eCDF P&L
      0,                     // VAT (Set to 0 or calculate if needed)
      'One Time',            // Frequency
      'RS',                  // Paid By
      'Sales Event',         // Event
      'Sync from Inventory'  // Notes
    ];

    accSheet.appendRow(row);
    return transId; // Return this to be saved in the Orders sheet
    
  } catch (e) {
    Logger.log("Accounting Sync Failed: " + e.message);
    return "SYNC_FAILED";
  }
}


function processSaveOrder(orderData) {
  // ... [Existing Logic to save order locally] ...

  // TRIGGER THE BOOKKEEPER
  const accRefId = postToAccounting(orderData);
  
  // Important: Now update the 'Orders' row you just created 
  // with this accRefId so you have a visual link.
  
  return { success: true, accId: accRefId };
}