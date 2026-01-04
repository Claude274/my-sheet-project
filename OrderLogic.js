/**
 * SERVER-SIDE: Order & Accounting Bridge
 */

const ACCOUNTING_SHEET_ID = '1XvCu3_1-89zdbL5NKEpnQsPJCoVyNsLqszgPA6Xz8EM'; 

/**
 * Main function to generate a PDF for an existing Order ID.
 */
function generateOrderInvoiceFromId(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Order Data
  const orderSheet = ss.getSheetByName("Orders");
  const orders = orderSheet.getDataRange().getValues();
  const oHeaders = orders[0].map(h => String(h).toLowerCase());
  const oRow = orders.find(r => String(r[oHeaders.indexOf("order_id")]) === orderId);
  if (!oRow) throw new Error("Order ID not found.");

  // 2. Prepare Template Data Object
  const invoiceData = {
    orderId: orderId,
    date: oRow[oHeaders.indexOf("date")],
    total: oRow[oHeaders.indexOf("total_amount")],
    company: {
      name: getSetting("COMPANY_NAME"),
      address: getSetting("COMPANY_ADDRESS"),
      phone: getSetting("COMPANY_PHONE"),
      email: getSetting("COMPANY_EMAIL"),
      website: getSetting("COMPANY_WEBSITE"),
      iban: getSetting("COMPANY_IBAN"),
      vat: getSetting("COMPANY_VAT")
    }
  };

  // 3. Get Contact Info
  const contactId = oRow[oHeaders.indexOf("contact_id")];
  const contactSheet = ss.getSheetByName("Contacts");
  const cData = contactSheet.getDataRange().getValues();
  const cHeaders = cData[0].map(h => String(h).toLowerCase());
  const cRow = cData.find(r => String(r[cHeaders.indexOf("contact_id")]) === String(contactId));
  
  invoiceData.customer = {
    id: contactId,
    name: oRow[oHeaders.indexOf("contact_name")],
    address: cRow ? cRow[cHeaders.indexOf("address")] : "N/A",
    email: cRow ? cRow[cHeaders.indexOf("email")] : "N/A",
    phone: cRow ? cRow[cHeaders.indexOf("phone")] : "N/A"
  };

  // 4. Get Product Items (Linking Units to Parent Product)
  const unitIds = String(oRow[oHeaders.indexOf("product_units_ids")] || "").split(",").map(id => id.trim());
  const unitSheet = ss.getSheetByName("Product_Units");
  const uData = unitSheet.getDataRange().getValues();
  const uHeaders = uData[0].map(h => String(h).toLowerCase());
  
  const productSheet = ss.getSheetByName("Product");
  const pData = productSheet.getDataRange().getValues();
  const pHeaders = pData[0].map(h => String(h).toLowerCase());

  invoiceData.items = unitIds.map(uId => {
    const unitRow = uData.find(r => String(r[uHeaders.indexOf("inventory_id")]) === uId);
    if (!unitRow) return { sku: uId, desc: "Not Found", price: 0 };
    
    const productId = unitRow[uHeaders.indexOf("product_id")];
    const prodRow = pData.find(r => String(r[pHeaders.indexOf("product_id")]) === String(productId));
    
    return {
      sku: unitRow[uHeaders.indexOf("sku_code")],
      desc: prodRow ? prodRow[pHeaders.indexOf("product_name")] : "Unique Art Piece",
      price: unitRow[uHeaders.indexOf("[product.sell_price]")]
    };
  });

  return createPdfInvoice(invoiceData);
}

/**
 * Creates the PDF using the Settings template and folder IDs.
 */
function createPdfInvoice(data) {
  const templateId = getSetting("PDF_TEMPLATE_ID");
  const folderId = getSetting("PDF_FOLDER_ID");
  const folder = DriveApp.getFolderById(folderId);
  
  const fileName = `Invoice_${data.orderId}.pdf`;
  const copy = DriveApp.getFileById(templateId).makeCopy(`TEMP_${data.orderId}`, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // Replacements
  body.replaceText("{{order_id}}", data.orderId);
  body.replaceText("{{date}}", Utilities.formatDate(new Date(data.date), "GMT+1", "dd/MM/yyyy"));
  body.replaceText("{{company_name}}", data.company.name);
  body.replaceText("{{contact_name}}", data.customer.name);
  body.replaceText("{{contact_address}}", data.customer.address);
  body.replaceText("{{total_amount}}", Number(data.total).toFixed(2) + " €");
  // ... add more as needed ...

  // Dynamic Table Logic
  const table = body.getTables()[0];
  const rowTemplate = table.getRow(1);
  data.items.forEach(item => {
    const newRow = table.appendTableRow(rowTemplate.copy());
    newRow.replaceText("{{item_desc}}", item.desc);
    newRow.replaceText("{{item_sku}}", item.sku);
    newRow.replaceText("{{item_price}}", Number(item.price).toFixed(2) + " €");
  });
  table.removeRow(1); // Remove template row

  doc.saveAndClose();
  const pdf = folder.createFile(copy.getAs(MimeType.PDF)).setName(fileName);
  copy.setTrashed(true);
  return pdf.getUrl();
}
/**
 * Gathers rich data by linking Orders -> Contacts -> Units -> Products
 */
function prepareInvoiceData(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  const orders = orderSheet.getDataRange().getValues();
  const oHeaders = orders[0].map(h => String(h).toLowerCase());
  
  // 1. Find the Order Row (using order_id from CSV)
  const oRow = orders.find(r => String(r[oHeaders.indexOf("order_id")]) === orderId);
  if (!oRow) throw new Error("Order ID not found in sheet.");

  // 2. Fetch Company Info from Settings
  const company = {
    name: getSetting("COMPANY_NAME") || "LAD Art",
    address: getSetting("COMPANY_ADDRESS") || "",
    vat: getSetting("COMPANY_VAT") || "",
    iban: getSetting("COMPANY_IBAN") || ""
  };

  // 3. Fetch Customer Info (using contact_id)
  const contactId = oRow[oHeaders.indexOf("contact_id")];
  const contactSheet = ss.getSheetByName("Contacts");
  const cData = contactSheet.getDataRange().getValues();
  const cHeaders = cData[0].map(h => String(h).toLowerCase());
  const cRow = cData.find(r => String(r[cHeaders.indexOf("contact_id")]) === String(contactId));
  
  const customer = {
    name: oRow[oHeaders.indexOf("contact_name")],
    email: cRow ? cRow[cHeaders.indexOf("email")] : "N/A",
    address: cRow ? cRow[cHeaders.indexOf("address")] : "N/A",
    phone: cRow ? cRow[cHeaders.indexOf("phone")] : ""
  };

  // 4. Fetch Product Details (Linking Unit to Parent Product)
  const itemIdsString = String(oRow[oHeaders.indexOf("product_units_ids")] || "");
  const itemIds = itemIdsString.split(",").map(id => id.trim()).filter(id => id);

  const unitSheet = ss.getSheetByName("Product_Units");
  const uData = unitSheet.getDataRange().getValues();
  const uHeaders = uData[0].map(h => String(h).toLowerCase());
  const uIdIdx = uHeaders.indexOf("product_units_id") !== -1 ? uHeaders.indexOf("product_units_id") : uHeaders.indexOf("inventory_id");
  
  const productSheet = ss.getSheetByName("Product");
  const pData = productSheet.getDataRange().getValues();
  const pHeaders = pData[0].map(h => String(h).toLowerCase());

  const items = itemIds.map(uId => {
    const unitRow = uData.find(r => String(r[uIdIdx]) === uId);
    const productId = unitRow ? unitRow[uHeaders.indexOf("product_id")] : null;
    const prodRow = pData.find(r => String(r[pHeaders.indexOf("product_id")]) === String(productId));
    
    return {
      sku: unitRow ? unitRow[uHeaders.indexOf("sku_code")] : "N/A",
      desc: prodRow ? prodRow[pHeaders.indexOf("product_name")] : "Unique Art Piece",
      price: unitRow ? unitRow[uHeaders.indexOf("[product.sell_price]")] : 0
    };
  });

  return { orderId, date: oRow[oHeaders.indexOf("date")], total: oRow[oHeaders.indexOf("total_amount")], company, customer, items };
}

/**
 * Injects data into Google Doc and handles the Item Table
 */
function generateProfessionalInvoice(data) {
  const templateId = getSetting("PDF_TEMPLATE_ID");
  const folderId = getSetting("PDF_FOLDER_ID");
  const fileName = `Invoice_${data.orderId}.pdf`;

  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, DriveApp.getFolderById(folderId));
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // Field Replacements
  body.replaceText("{{company_name}}", data.company.name);
  body.replaceText("{{company_address}}", data.company.address);
  body.replaceText("{{company_vat}}", data.company.vat);
  body.replaceText("{{company_iban}}", data.company.iban);
  body.replaceText("{{contact_name}}", data.customer.name);
  body.replaceText("{{contact_email}}", data.customer.email);
  body.replaceText("{{contact_address}}", data.customer.address);
  body.replaceText("{{orders_id}}", data.orderId);
  body.replaceText("{{date}}", Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), "dd/MM/yyyy"));
  body.replaceText("{{total_amount}}", Number(data.total).toFixed(2) + " €");

  // Dynamic Table Injection
  const table = body.getTables()[0]; 
  const templateRow = table.getRow(1); // Row containing placeholders {{item_sku}} etc.
  
  data.items.forEach(item => {
    const newRow = table.appendTableRow(templateRow.copy());
    newRow.replaceText("{{item_sku}}", item.sku);
    newRow.replaceText("{{item_desc}}", item.desc);
    newRow.replaceText("{{item_price}}", Number(item.price).toFixed(2) + " €");
  });
  
  table.removeRow(1); // Delete the placeholder row
  doc.saveAndClose();
  
  const pdfFile = DriveApp.getFolderById(folderId).createFile(copy.getAs(MimeType.PDF));
  copy.setTrashed(true);
  return pdfFile.getUrl();
}

/**
 * Enhanced Order Process: Logs IDs and triggers PDF
 */
function processSaveOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderId = "ORD-" + Math.random().toString(36).substring(2, 9).toUpperCase();
  const timestamp = new Date();

  try {
    const accRefId = postToAccounting(orderData, orderId);
    
    // Log to Orders Sheet (must include product_units_ids for regeneration)
    const orderSheet = ss.getSheetByName("Orders");
    if (orderSheet) {
      orderSheet.appendRow([
        orderId, timestamp, orderData.contact_id, orderData.contact_name,
        orderData.total, orderData.itemIds.length, orderData.skuList, 
        orderData.itemIds.join(", "), // Stores IDs in product_units_ids column
        orderData.comment, "Cash Sale", "Auto-synced", accRefId
      ]);
    }

    const pdfUrl = generateOrderInvoiceFromId(orderId);
    updateInventoryStatus(orderData.itemIds);

    return { success: true, id: orderId, url: pdfUrl };
  } catch (err) {
    throw new Error("Order failed: " + err.message);
  }
}

// ... inside prepareInvoiceData(orderId) ...

  const items = itemIds.map(uId => {
    // 1. Find the Unit in Product_Units
    const unitRow = uData.find(r => String(r[uHeaders.indexOf("inventory_id")]) === uId);
    
    // 2. Get the Parent Product ID from that Unit
    const productId = unitRow ? unitRow[uHeaders.indexOf("product_id")] : null;
    
    // 3. Find the Description in the Product sheet
    const prodRow = pData.find(r => String(r[pHeaders.indexOf("product_id")]) === String(productId));
    
    return {
      sku: unitRow ? unitRow[uHeaders.indexOf("sku_code")] : "N/A",
      desc: prodRow ? prodRow[pHeaders.indexOf("product_name")] : "Unique Art Piece",
      price: unitRow ? unitRow[uHeaders.indexOf("[product.sell_price]")] : 0
    };
  });