/**
 * In OrderLogic.js: Separated logic for Order management
 */

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

  // Only return items with Status "Produced"
  return data.slice(1)
    .filter(row => String(row[statusIdx]).trim() === "Produced")
    .map(row => ({
      id: row[idIdx],
      sku: row[skuIdx],
      price: row[priceIdx] || 0
    }));
}

function processSaveOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderId = "orders_" + Math.random().toString(36).substring(2, 11);
  const timestamp = new Date();

  // 1. Save Order
  ss.getSheetByName("Orders").appendRow([
    orderId, timestamp, orderData.contact_id, orderData.contact_name, 
    orderData.total_amount, orderData.items.length, orderData.payment_method, orderData.notes
  ]);

  // 2. Log Transaction
  const transId = "transactions_" + Math.random().toString(36).substring(2, 11);
  ss.getSheetByName("Transactions").appendRow([
    transId, timestamp, "Income", "Sale: " + orderId, orderData.contact_name, orderData.total_amount
  ]);

  // 3. Set Inventory to 'Sold'
  const invSheet = ss.getSheetByName("Product_Units");
  const invData = invSheet.getDataRange().getValues();
  const idIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("inventory_id");
  const statusIdx = invData[0].map(h => String(h).toLowerCase()).indexOf("status");

  orderData.items.forEach(itemId => {
    const rIdx = invData.findIndex(r => r[idIdx] === itemId);
    if (rIdx > -1) invSheet.getRange(rIdx + 1, statusIdx + 1).setValue("Sold");
  });

  return { success: true };
}