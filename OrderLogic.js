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

function processSaveOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date();
  const orderId = "orders_" + Math.random().toString(36).substring(2, 11);

  try {
    // 1. PUSH TO ACCOUNTING FILE FIRST
    const accRefId = postToAccounting(orderData, orderId);

    // 2. LOG IN LOCAL ORDERS SHEET
    // order_id, date, contact_id, contact_name, total_amount, items_count, sku_list, comment, payment_method, notes, accounting_id
    const orderSheet = ss.getSheetByName("Orders");
    if (orderSheet) {
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
        accRefId // The link to your finance file
      ]);
    }

    // 3. UPDATE INVENTORY STATUS (Produced -> Sold)
    const invSheet = ss.getSheetByName("Product_Units");
    const invData = invSheet.getDataRange().getValues();
    const idIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("inventory_id");
    const statusIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("status");

    orderData.itemIds.forEach(itemId => {
      const rowIdx = invData.findIndex(r => r[idIdx] === itemId);
      if (rowIdx > -1) invSheet.getRange(rowIdx + 1, statusIdx + 1).setValue("Sold");
    });

    if (typeof updateAllComputedValues === 'function') updateAllComputedValues();
    
    return { success: true, id: orderId };

  } catch (err) {
    Logger.log("Critical Error in processSaveOrder: " + err.message);
    throw new Error("Order failed: " + err.message);
  }
}

function postToAccounting(orderData, orderId) {
  try {
    const accSS = SpreadsheetApp.openById(ACCOUNTING_SHEET_ID);
    const accSheet = accSS.getSheetByName('Transactions');
    const type = (orderData.payment_method === 'Cash') ? 'Cash Sale' : 'Sales Invoice - Goods';
    const template = ACC_TEMPLATES[type];
    const transId = "transactions_" + Math.random().toString(36).substring(2, 11);

    // Append to accounting file (19 columns)
    accSheet.appendRow([
      transId, new Date(), type, "Order " + orderId, orderData.contact_name, orderData.total,
      template.debit.pcn, template.debit.name, template.debit.ecdfBS, template.debit.ecdfPL,
      template.credit.pcn, template.credit.name, template.credit.ecdfBS, template.credit.ecdfPL,
      0, 'One Time', 'RS', 'Sales Event', 'Sync from Inventory'
    ]);
    return transId;
  } catch (e) {
    Logger.log("Accounting Bridge Failed: " + e.message);
    return "SYNC_FAILED";
  }
}