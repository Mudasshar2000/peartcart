// ============================================================
// PearlCart — Google Apps Script
// Paste this entire file into your Google Sheet's Apps Script
// Extensions > Apps Script > paste > Save > Deploy
// ============================================================

const PRODUCTS_SHEET = "Products";
const ORDERS_SHEET   = "Orders";

// ── MAIN HANDLER ──
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'saveOrder') {
      return handleSaveOrder(e);
    }
    return handleGetProducts();
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    return handleSaveOrder(e);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ── GET PRODUCTS ──
function handleGetProducts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_SHEET);
  if (!sheet) return jsonResponse({ error: "Sheet '" + PRODUCTS_SHEET + "' not found." });

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => h.toString().trim().toLowerCase().replace(/\s+/g, "_"));
  const products = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every(cell => cell === "" || cell === null)) continue;

    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });

    const product = {
      id:          Number(obj["id"])   || i,
      name:        String(obj["name"]  || "").trim(),
      price:       Number(obj["price"]) || 0,
      type:        String(obj["type"]  || "physical").toLowerCase().trim(),
      stock:       String(obj["type"]  || "").toLowerCase().trim() === "digital"
                     ? 999 : Number(obj["stock"]) || 0,
      sku:         String(obj["sku"]   || "SKU-" + i).trim(),
      category:    String(obj["category"] || "Other").trim(),
      description: String(obj["description"] || "").trim(),
      image_url:   String(obj["image_url"]   || "").trim(),
      emoji:       String(obj["emoji"]       || "📦").trim(),
      active:      String(obj["active"]      || "yes").toLowerCase().trim(),
    };

    if (["yes","true","1"].includes(product.active)) {
      products.push(product);
    }
  }

  return jsonResponse({ success: true, count: products.length, products });
}

// ── SAVE ORDER ──
function handleSaveOrder(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ORDERS_SHEET);

  // Auto-create Orders sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(ORDERS_SHEET);
    const headers = [
      "Order ID","Date","Time",
      "First Name","Last Name","Email","Phone","NIC",
      "Address Line 1","Address Line 2","City","Province","District","Postal Code","Country",
      "Items","Subtotal (Rs.)","Shipping (Rs.)","Tax (Rs.)","Total (Rs.)",
      "Payment Method","Status","Notes"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1A1A1A");
    headerRange.setFontColor("#C9A84C");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 180);
    sheet.setColumnWidth(16, 300); // Items column wider
  }

  // Parse order data
  let orderData = {};
  try {
    const body = e.postData ? e.postData.contents : (e.parameter ? e.parameter.data : "{}");
    orderData = JSON.parse(body || "{}");
  } catch (parseErr) {
    // Try reading from GET params as fallback
    if (e && e.parameter) orderData = e.parameter;
  }

  // Build row
  const row = [
    orderData.orderId    || "",
    orderData.date       || new Date().toLocaleDateString("en-GB"),
    orderData.time       || new Date().toLocaleTimeString("en-GB"),
    orderData.firstName  || "",
    orderData.lastName   || "",
    orderData.email      || "",
    orderData.phone      || "",
    orderData.nic        || "",
    orderData.address1   || "",
    orderData.address2   || "",
    orderData.city       || "",
    orderData.province   || "",
    orderData.district   || "",
    orderData.postalCode || "",
    orderData.country    || "Sri Lanka",
    orderData.items      || "",
    orderData.subtotal   || "0.00",
    orderData.shipping   || "0.00",
    orderData.tax        || "0.00",
    orderData.total      || "0.00",
    orderData.payment    || "Cash on Delivery",
    orderData.status     || "Pending",
    orderData.notes      || "",
  ];

  sheet.appendRow(row);

  // Auto-resize columns after adding data
  sheet.autoResizeColumns(1, 23);

  return jsonResponse({ success: true, orderId: orderData.orderId, message: "Order saved successfully" });
}

// ── JSON HELPER ──
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── TEST FUNCTIONS ──
function testGetProducts() {
  const result = handleGetProducts();
  Logger.log(result.getContent());
}

function testSaveOrder() {
  const mockOrder = {
    orderId: "PC-99999",
    date: "01/01/2026",
    time: "10:30",
    firstName: "Amal",
    lastName: "Perera",
    email: "amal@test.com",
    phone: "0771234567",
    nic: "199912345678",
    address1: "No. 12, Galle Road",
    address2: "",
    city: "Colombo",
    province: "Western",
    district: "Colombo",
    postalCode: "00300",
    country: "Sri Lanka",
    items: "Wireless Headphones x1 (Rs.8999.00)",
    subtotal: "8999.00",
    shipping: "500.00",
    tax: "899.90",
    total: "10398.90",
    payment: "Cash on Delivery",
    status: "Pending",
    notes: "Please call before delivery",
  };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ORDERS_SHEET) || ss.insertSheet(ORDERS_SHEET);
  Logger.log("Test order data: " + JSON.stringify(mockOrder));
}
