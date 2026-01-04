/**
 * ==========================================
 * LAD ORDER MANAGEMENT & LOGGING SYSTEM
 * ==========================================
 */

// 1. OPEN THE POS FORM
function openOrderModal() {
  try {
    const html = HtmlService.createTemplateFromFile('OrderForm')
      .evaluate()
      .setTitle('🛒 New Customer Order')
      .setWidth(900)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, '🛒 New Customer Order');
  } catch(e) {
    SpreadsheetApp.getUi().alert("Error: 'OrderForm' file not found. Ensure your HTML file is named exactly 'OrderForm'.");
  }
}

/**
 * 2. DATA FETCHING FOR POS
 * Fetches contacts and available inventory for the OrderForm.
 */
function getOrderFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Fetch Products & Units
  const unitSheet = ss.getSheetByName('Product_Units') || ss.getSheetByName('Inventar');
  const unitData = unitSheet.getDataRange().getValues();
  const inventory = [];
  const uHeaders = unitData[0].map(h => String(h).toLowerCase());
  
  // Using our [sheet_name]_id rule
  const uIdIdx = uHeaders.indexOf('inventory_id');
  const uStatusIdx = uHeaders.indexOf('status');
  const uSukIdx = uHeaders.indexOf('suk_code');

  for (let i = 1; i < unitData.length; i++) {
    if (String(unitData[i][uStatusIdx]).toLowerCase() === 'produced') {
      inventory.push({
        id: unitData[i][uIdIdx],
        label: unitData[i][uSukIdx] || "Unit " + i
      });
    }
  }

  // Fetch Contacts
  const contactSheet = ss.getSheetByName('Contacts');
  const contactData = contactSheet.getDataRange().getValues();
  const contacts = contactData.slice(1).map(row => ({
    id: row[0],
    name: row[2] || (row[3] + " " + row[5])
  }));

  return { inventory: inventory, contacts: contacts };
}

/**
 * 3. THE COMPLETED LOGGING LOGIC
 * This handles the actual 'Create Order' event from the OrderForm.
 */
function submitOrderData(orderPayload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  
  // Wait up to 30 seconds for other processes to finish
  if (!lock.tryLock(30000)) throw new Error("System busy. Try again in a moment.");

  try {
    const date = new Date();
    // Rule: unique ID as [sheet_name]_id
    const orderId = "ORD-" + Utilities.formatDate(date, "GMT", "yyyyMMdd-HHmmss");
    
    // A. LOG TO ORDERS SHEET
    const orderSheet = ss.getSheetByName("Orders");
    // [orders_id, date, contact_id, customer_name, total_amount, item_count, status]
    orderSheet.appendRow([
      orderId, 
      date, 
      orderPayload.contactId, 
      orderPayload.contactName, 
      orderPayload.total, 
      orderPayload.items.length,
      "Completed"
    ]);

    // B. LOG TO ORDERDETAILS SHEET
    const detailSheet = ss.getSheetByName("OrderDetails");
    
    orderPayload.items.forEach((item, index) => {
      const detailId = "DET-" + orderId + "-" + (index + 1); // [sheet_name]_id
      
      detailSheet.appendRow([
        detailId,      // orderdetails_id
        orderId,       // orders_id (Foreign Key)
        item.unit_id,  // inventory_id
        item.price,
        item.qty || 1
      ]);

      // C. MARK UNIT AS SOLD
      updateUnitStatus(item.unit_id, "Sold");
    });

    return { success: true, orderId: orderId };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper to update unit status in Product_Units
 */
function updateUnitStatus(unitId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product_Units') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventar');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase());
  const idIdx = headers.indexOf('inventory_id');
  const statusIdx = headers.indexOf('status');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(unitId)) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      break;
    }
  }
}