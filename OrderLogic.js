/**
 * SERVER-SIDE: Order & Transaction Logic
 */

function getInventoryForOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Product_Units");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase());
  
  const statusIdx = headers.indexOf("status");
  const skuIdx = headers.indexOf("sku_code");
  const priceIdx = headers.indexOf("[product.sell_price]"); 
  const idIdx = headers.indexOf("inventory_id");

  return data.slice(1)
    .filter(row => row[statusIdx] === "Produced")
    .map(row => ({
      id: row[idIdx],
      sku: row[skuIdx],
      price: row[priceIdx] || 0,
      label: row[skuIdx] + " - " + (row[priceIdx] || 0) + "€"
    }));
}

function processSaveOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date();
  const orderId = "orders_" + Math.random().toString(36).substring(2, 11);

  // 1. Log in Orders Sheet
  const orderSheet = ss.getSheetByName("Orders");
  if (orderSheet) {
    orderSheet.appendRow([
      orderId, timestamp, orderData.contact_id, orderData.contact_name,
      orderData.total_amount, orderData.items.length, orderData.payment_method, orderData.notes
    ]);
  }

  // 2. Update Inventory
  const invSheet = ss.getSheetByName("Product_Units");
  if (invSheet) {
    const invData = invSheet.getDataRange().getValues();
    const idIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("inventory_id");
    const statusIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("status");

    orderData.items.forEach(itemId => {
      const rowIdx = invData.findIndex(r => r[idIdx] === itemId);
      if (rowIdx > -1) invSheet.getRange(rowIdx + 1, statusIdx + 1).setValue("Sold");
    });
  }

  // 3. Log Transaction
  const transSheet = ss.getSheetByName("Transactions");
  if (transSheet) {
    const transId = "transactions_" + Math.random().toString(36).substring(2, 11);
    transSheet.appendRow([
      transId, timestamp, "Income", "Sale: Order " + orderId,
      orderData.contact_name, orderData.total_amount
    ]);
  }

  if (typeof forceRunAllTriggers === 'function') forceRunAllTriggers();
  return { success: true, orderId: orderId };
}