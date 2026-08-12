/*******************************************************
 * PLANET OF GIFTS AND TOYS — real backend (Sheets API)
 * -----------------------------------------------------
 * Ports Code.gs 1:1, but reads/writes your real Google
 * Sheet via a service account instead of running inside
 * Apps Script. Serves the untouched Index.html + a tiny
 * client shim (real-backend.js) that replaces
 * google.script.run with fetch() calls to /api/call/:fn
 *******************************************************/
const express = require("express");
const cors = require("cors");
const path = require("path");
const { google } = require("googleapis");

const app = express();
app.use(cors());
app.use(express.json({ limit: "15mb" }));
app.use(express.static(path.join(__dirname)));

// ---------- CONFIG ----------
const SELLER_PASSWORD = process.env.SELLER_PASSWORD || "Toys@5831";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "Admin@5831";
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || "planetofgiftsandtoys@gmail.com";
const UPI_ID = process.env.UPI_ID || "9711978206@ptyes";
const PAYEE_NAME = process.env.PAYEE_NAME || "Santosh Kumar";
const GST_RATE = 0.18;
const FREE_SHIPPING_MIN = 499;
const SHIPPING_CHARGE = 70;
const SHEET_ID = process.env.SHEET_ID;

const PRODUCTS_SHEET = "Products";
const ORDERS_SHEET = "Orders";
const SELLERS_SHEET = "Sellers";

const PRODUCT_HEADERS = ["Product ID","Seller ID","Seller Name","Title","Category","Price (Rs)","Stock","Description","Image URLs","Video URL"];
const SELLER_HEADERS = ["Seller ID","Timestamp","Seller Name","Business Name","Email","Contact","Address","Account Holder Name","Account Number","IFSC Code","Bank Name","UPI ID","Status","Rejection Reason","Login Password","Approved At"];
const ORDER_HEADERS = ["Timestamp","Order ID","Customer Name","Phone","Address","Items","Items JSON","Seller IDs","Item Subtotal (Rs)","Base Price (Rs)","GST 18% (Rs)","Shipping (Rs)","Grand Total (Rs)","Payment Mode","Amount To Pay Now (Rs)","Remaining On Delivery (Rs)","Payment Status","UTR / Transaction ID","Order Status","Cancellation Reason"];

// ---------- GOOGLE AUTH ----------
function getAuth() {
  const email = process.env.GOOGLE_CLIENT_EMAIL;
  const key = (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n");
  if (!email || !key) throw new Error("Missing GOOGLE_CLIENT_EMAIL / GOOGLE_PRIVATE_KEY env vars.");
  return new google.auth.JWT(email, null, key, [
    "https://www.googleapis.com/auth/spreadsheets",
  ]);
}
const sheetsApi = () => google.sheets({ version: "v4", auth: getAuth() });

// ---------- LOW-LEVEL SHEET HELPERS ----------
const colLetter = (n) => {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

async function getSpreadsheetMeta() {
  const sheets = sheetsApi();
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return res.data.sheets.map(s => s.properties);
}

async function ensureSheet(name, headers) {
  const sheets = sheetsApi();
  const props = await getSpreadsheetMeta();
  const existing = props.find(p => p.title === name);
  if (!existing) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: name } } }] }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${name}!A1:${colLetter(headers.length)}1`,
      valueInputOption: "RAW",
      requestBody: { values: [headers] }
    });
  }
}
async function ensureAllSheets() {
  await ensureSheet(PRODUCTS_SHEET, PRODUCT_HEADERS);
  await ensureSheet(SELLERS_SHEET, SELLER_HEADERS);
  await ensureSheet(ORDERS_SHEET, ORDER_HEADERS);
}

async function getRows(sheetName) {
  const sheets = sheetsApi();
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${sheetName}!A1:Z10000` });
  const values = res.data.values || [];
  const headers = values[0] || [];
  const rows = values.slice(1);
  return { headers, rows };
}
async function appendRow(sheetName, rowArray) {
  const sheets = sheetsApi();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [rowArray] }
  });
}
async function updateCell(sheetName, rowNumber1Indexed, colNumber1Indexed, value) {
  const sheets = sheetsApi();
  const range = `${sheetName}!${colLetter(colNumber1Indexed)}${rowNumber1Indexed}:${colLetter(colNumber1Indexed)}${rowNumber1Indexed}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID, range, valueInputOption: "RAW", requestBody: { values: [[value]] }
  });
}
async function deleteRow(sheetName, rowNumber1Indexed) {
  const sheets = sheetsApi();
  const props = await getSpreadsheetMeta();
  const sheetProps = props.find(p => p.title === sheetName);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId: sheetProps.sheetId, dimension: "ROWS", startIndex: rowNumber1Indexed - 1, endIndex: rowNumber1Indexed }
        }
      }]
    }
  });
}
function colIndex(headers, name) {
  const i = headers.indexOf(name);
  if (i === -1) throw new Error("Header not found: " + name);
  return i; // 0-based
}

// ---------- HELPERS ----------
const round2 = n => Math.round(n * 100) / 100;
const uid = p => p + Date.now();

function getAdvanceAmount(grandTotal) {
  if (grandTotal < 299) return 30;
  if (grandTotal < 499) return 50;
  if (grandTotal < 999) return 90;
  if (grandTotal < 1499) return 150;
  if (grandTotal < 2000) return 190;
  return 250;
}

// ---------- PRODUCTS ----------
async function getProducts() {
  await ensureAllSheets();
  const { rows } = await getRows(PRODUCTS_SHEET);
  return rows.filter(r => r[0]).map(r => ({
    id: r[0], sellerId: r[1], sellerName: r[2], title: r[3], category: r[4],
    price: Number(r[5]), stock: Number(r[6]), desc: r[7] || "",
    images: r[8] ? String(r[8]).split("|").filter(Boolean) : [],
    videoUrl: r[9] || ""
  }));
}

function verifySellerPassword(password) { return password === SELLER_PASSWORD; }

async function getApprovedSellers() {
  await ensureAllSheets();
  const { headers, rows } = await getRows(SELLERS_SHEET);
  const statusCol = colIndex(headers, "Status");
  return rows.filter(r => r[0] && r[statusCol] === "Approved").map(r => ({ id: r[0], name: r[2], businessName: r[3] }));
}

async function createProductRow(sellerId, product) {
  if (!product.title || !product.price || !product.stock) return { success: false, message: "Title, price and stock are required." };
  await ensureAllSheets();
  const sellers = await getApprovedSellers();
  const seller = sellers.find(s => s.id === sellerId);
  if (!seller) return { success: false, message: "Seller is not an approved vendor." };
  const id = uid("P");
  // Images/video: stored as data URLs directly (small images only — see README for limits)
  const imageUrls = (product.images || []).map(img => `data:${img.mime};base64,${img.base64}`).filter(u => u.length < 45000);
  const videoUrl = ""; // video upload not supported in this backend yet (data URLs too large for Sheets cells)
  await appendRow(PRODUCTS_SHEET, [id, sellerId, seller.name, product.title, product.category || "General",
    Number(product.price), Number(product.stock), product.desc || "", imageUrls.join("|"), videoUrl]);
  return { success: true, id, imageUrls, videoUrl };
}

async function addProduct(password, product) {
  if (password !== SELLER_PASSWORD) return { success: false, message: "Incorrect owner password." };
  if (!product.sellerId) return { success: false, message: "Select which approved vendor this listing belongs to." };
  return createProductRow(product.sellerId, product);
}
async function addProductAsSeller(email, password, product) {
  const auth = await authenticateSeller(email, password);
  if (!auth.success) return auth;
  return createProductRow(auth.seller.id, product);
}

// ---------- SELLERS ----------
async function registerSeller(sellerData) {
  if (!sellerData.sellerName || !sellerData.businessName || !sellerData.email || !sellerData.contact || !sellerData.address) {
    return { success: false, message: "Please fill in all required seller details." };
  }
  await ensureAllSheets();
  const { headers, rows } = await getRows(SELLERS_SHEET);
  const emailCol = colIndex(headers, "Email");
  if (rows.some(r => String(r[emailCol] || "").toLowerCase() === sellerData.email.toLowerCase())) {
    return { success: false, message: "This email is already registered. Try logging in instead, or contact us if you're waiting on approval." };
  }
  const id = uid("S");
  await appendRow(SELLERS_SHEET, [
    id, new Date().toISOString(), sellerData.sellerName, sellerData.businessName, sellerData.email, sellerData.contact,
    sellerData.address, sellerData.accountHolder || "", sellerData.accountNumber || "",
    sellerData.ifsc || "", sellerData.bankName || "", sellerData.upiId || "", "Pending", "", "", ""
  ]);
  return { success: true, id };
}

function verifyAdminPassword(password) { return password === ADMIN_PASSWORD; }

async function getAllSellers(password) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  await ensureAllSheets();
  const { rows } = await getRows(SELLERS_SHEET);
  const list = rows.filter(r => r[0]).map(r => ({
    id: r[0], timestamp: r[1], sellerName: r[2], businessName: r[3], email: r[4],
    contact: r[5], address: r[6], accountHolder: r[7], accountNumber: r[8],
    ifsc: r[9], bankName: r[10], upiId: r[11], status: r[12], rejectionReason: r[13]
  })).reverse();
  return { success: true, sellers: list };
}

async function updateSellerStatus(password, sellerId, newStatus, reason) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  await ensureAllSheets();
  const { headers, rows } = await getRows(SELLERS_SHEET);
  const idCol = colIndex(headers, "Seller ID"), statusCol = colIndex(headers, "Status"), reasonCol = colIndex(headers, "Rejection Reason"),
        pwdCol = colIndex(headers, "Login Password"), approvedAtCol = colIndex(headers, "Approved At");
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][idCol] === sellerId) {
      const rowNum = i + 2; // 1-indexed + header row
      await updateCell(SELLERS_SHEET, rowNum, statusCol + 1, newStatus);
      if (reason) await updateCell(SELLERS_SHEET, rowNum, reasonCol + 1, reason);
      if (newStatus === "Approved") {
        let loginPassword = rows[i][pwdCol];
        if (!loginPassword) {
          loginPassword = generatePassword();
          await updateCell(SELLERS_SHEET, rowNum, pwdCol + 1, loginPassword);
        }
        await updateCell(SELLERS_SHEET, rowNum, approvedAtCol + 1, new Date().toISOString());
      }
      return { success: true };
    }
  }
  return { success: false, message: "Seller not found." };
}

function generatePassword() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let pwd = "";
  for (let i = 0; i < 8; i++) pwd += chars.charAt(Math.floor(Math.random() * chars.length));
  return pwd;
}

async function authenticateSeller(email, password) {
  if (!email || !password) return { success: false, message: "Enter your email and password." };
  await ensureAllSheets();
  const { headers, rows } = await getRows(SELLERS_SHEET);
  const emailCol = colIndex(headers, "Email"), pwdCol = colIndex(headers, "Login Password"), statusCol = colIndex(headers, "Status");
  const row = rows.find(r => String(r[emailCol] || "").toLowerCase() === String(email).toLowerCase());
  if (!row) return { success: false, message: "No seller account found for this email." };
  if (row[statusCol] !== "Approved") return { success: false, message: "Your seller account is not approved yet (status: " + row[statusCol] + ")." };
  if (String(row[pwdCol]) !== String(password)) return { success: false, message: "Incorrect password." };
  return { success: true, seller: { id: row[0], name: row[2], businessName: row[3], email: row[4] } };
}
async function sellerLogin(email, password) { return authenticateSeller(email, password); }

async function getSellerListings(email, password) {
  const auth = await authenticateSeller(email, password);
  if (!auth.success) return auth;
  const products = (await getProducts()).filter(p => p.sellerId === auth.seller.id);
  return { success: true, products };
}

async function getSellerOrders(email, password) {
  const auth = await authenticateSeller(email, password);
  if (!auth.success) return auth;
  const sellerId = auth.seller.id;
  await ensureAllSheets();
  const { headers, rows } = await getRows(ORDERS_SHEET);
  const sellerIdsCol = colIndex(headers, "Seller IDs"), itemsJsonCol = colIndex(headers, "Items JSON");
  const matching = rows.filter(r => r[1] && String(r[sellerIdsCol] || "").split(",").indexOf(sellerId) !== -1);
  const orders = matching.map(r => {
    let myItems = [];
    try { myItems = (JSON.parse(r[itemsJsonCol] || "[]")).filter(li => li.sellerId === sellerId); } catch (e) {}
    const myTotal = round2(myItems.reduce((s, li) => s + (li.lineTotal || 0), 0));
    return { timestamp: r[0], orderId: r[1], customerName: r[2], phone: r[3], address: r[4], myItems, myTotal,
      paymentMode: r[13], paymentStatus: r[16], orderStatus: r[18] };
  }).reverse();
  return { success: true, orders };
}

// ---------- PRICING ----------
async function calculateOrder(cartItems, paymentMode) {
  const products = await getProducts();
  const map = {}; products.forEach(p => map[p.id] = p);
  let subtotal = 0; const lineItems = [];
  cartItems.forEach(ci => {
    const p = map[ci.id]; if (!p) return;
    const lineTotal = round2(p.price * ci.qty);
    subtotal += lineTotal;
    lineItems.push({ id: p.id, title: p.title, price: p.price, qty: ci.qty, lineTotal, sellerId: p.sellerId, sellerName: p.sellerName });
  });
  subtotal = round2(subtotal);
  const gst = round2(subtotal * GST_RATE);
  const base = round2(subtotal - gst);
  const shipping = subtotal < FREE_SHIPPING_MIN ? SHIPPING_CHARGE : 0;
  const grandTotal = round2(subtotal + shipping);
  const advance = getAdvanceAmount(grandTotal);
  let amountDueNow, remainingAmount;
  if (paymentMode === "Prepaid") { amountDueNow = grandTotal; remainingAmount = 0; }
  else { amountDueNow = advance; remainingAmount = round2(grandTotal - advance); }
  return { lineItems, subtotal, base, gst, shipping, grandTotal, advance, paymentMode, amountDueNow, remainingAmount };
}

function getUpiLink(amount, orderId) {
  const params = ["pa=" + encodeURIComponent(UPI_ID), "pn=" + encodeURIComponent(PAYEE_NAME),
    "am=" + encodeURIComponent(amount), "cu=INR", "tn=" + encodeURIComponent("Order " + orderId)].join("&");
  return "upi://pay?" + params;
}

// ---------- ORDERS ----------
async function reduceStock(lineItems) {
  await ensureAllSheets();
  const { rows } = await getRows(PRODUCTS_SHEET);
  for (const li of lineItems) {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === li.id) {
        const current = Number(rows[i][6]);
        await updateCell(PRODUCTS_SHEET, i + 2, 7, Math.max(current - li.qty, 0));
        break;
      }
    }
  }
}
async function restoreStock(lineItems) {
  await ensureAllSheets();
  const { rows } = await getRows(PRODUCTS_SHEET);
  for (const li of lineItems) {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === li.id) {
        const current = Number(rows[i][6]);
        await updateCell(PRODUCTS_SHEET, i + 2, 7, current + li.qty);
        break;
      }
    }
  }
}

async function placeOrder(orderInfo) {
  await ensureAllSheets();
  if (!orderInfo || !orderInfo.customer || !orderInfo.cartItems || !orderInfo.cartItems.length) {
    return { success: false, message: "Missing order details." };
  }
  const calc = await calculateOrder(orderInfo.cartItems, orderInfo.paymentMode);
  if (calc.lineItems.length === 0) return { success: false, message: "Cart items not found." };
  const orderId = uid("ORD");
  const itemsSummary = calc.lineItems.map(li => li.title + " x" + li.qty).join(", ");
  const itemsJson = JSON.stringify(calc.lineItems);
  const sellerIds = [...new Set(calc.lineItems.map(li => li.sellerId).filter(Boolean))].join(",");
  await appendRow(ORDERS_SHEET, [
    new Date().toISOString(), orderId, orderInfo.customer.name, orderInfo.customer.phone, orderInfo.customer.address,
    itemsSummary, itemsJson, sellerIds, calc.subtotal, calc.base, calc.gst, calc.shipping, calc.grandTotal,
    orderInfo.paymentMode, calc.amountDueNow, calc.remainingAmount, "Pending Verification", "", "Received", ""
  ]);
  await reduceStock(calc.lineItems);
  const upiLink = getUpiLink(calc.amountDueNow, orderId);
  return { success: true, orderId, calc, upiLink };
}

async function submitUtr(orderId, utr) {
  if (!orderId || !utr) return { success: false, message: "Order ID and UTR are required." };
  await ensureAllSheets();
  const { headers, rows } = await getRows(ORDERS_SHEET);
  const idCol = colIndex(headers, "Order ID"), utrCol = colIndex(headers, "UTR / Transaction ID"), payCol = colIndex(headers, "Payment Status");
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][idCol] === orderId) {
      await updateCell(ORDERS_SHEET, i + 2, utrCol + 1, utr);
      await updateCell(ORDERS_SHEET, i + 2, payCol + 1, "UTR Submitted - Verify");
      return { success: true };
    }
  }
  return { success: false, message: "Order not found." };
}

// ---------- ADMIN: ORDERS ----------
async function getOrders(password, limit) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  await ensureAllSheets();
  const { rows } = await getRows(ORDERS_SHEET);
  const orders = rows.filter(r => r[1]).map(r => ({
    timestamp: r[0], orderId: r[1], customerName: r[2], phone: r[3], address: r[4],
    items: r[5], subtotal: r[8], base: r[9], gst: r[10], shipping: r[11], grandTotal: r[12],
    paymentMode: r[13], amountDueNow: r[14], remaining: r[15], paymentStatus: r[16],
    utr: r[17], orderStatus: r[18], cancelReason: r[19]
  })).reverse();
  return { success: true, orders: limit ? orders.slice(0, limit) : orders };
}

async function updateOrderStatus(password, orderId, newStatus, reason) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  await ensureAllSheets();
  const { headers, rows } = await getRows(ORDERS_SHEET);
  const idCol = colIndex(headers, "Order ID"), statusCol = colIndex(headers, "Order Status"),
        reasonCol = colIndex(headers, "Cancellation Reason"), itemsJsonCol = colIndex(headers, "Items JSON");
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][idCol] === orderId) {
      const prevStatus = rows[i][statusCol];
      await updateCell(ORDERS_SHEET, i + 2, statusCol + 1, newStatus);
      if (newStatus === "Cancelled") await updateCell(ORDERS_SHEET, i + 2, reasonCol + 1, reason || "");
      if (newStatus === "Cancelled" && prevStatus !== "Cancelled") {
        try { await restoreStock(JSON.parse(rows[i][itemsJsonCol] || "[]")); } catch (e) {}
      }
      return { success: true };
    }
  }
  return { success: false, message: "Order not found." };
}

async function updatePaymentStatus(password, orderId, newPaymentStatus) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  await ensureAllSheets();
  const { headers, rows } = await getRows(ORDERS_SHEET);
  const idCol = colIndex(headers, "Order ID"), payCol = colIndex(headers, "Payment Status");
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][idCol] === orderId) { await updateCell(ORDERS_SHEET, i + 2, payCol + 1, newPaymentStatus); return { success: true }; }
  }
  return { success: false, message: "Order not found." };
}

// ---------- PRODUCT EDIT / DELETE ----------
async function findProductRow(productId) {
  const { headers, rows } = await getRows(PRODUCTS_SHEET);
  const idCol = colIndex(headers, "Product ID");
  for (let i = 0; i < rows.length; i++) if (rows[i][idCol] === productId) return { rowNum: i + 2, rowValues: rows[i], headers };
  return null;
}
async function applyProductUpdate(productId, updates) {
  const found = await findProductRow(productId);
  if (!found) return { success: false, message: "Product not found." };
  const { rowNum, headers } = found;
  if (updates.title) await updateCell(PRODUCTS_SHEET, rowNum, colIndex(headers, "Title") + 1, updates.title);
  if (updates.category) await updateCell(PRODUCTS_SHEET, rowNum, colIndex(headers, "Category") + 1, updates.category);
  if (updates.price !== undefined && updates.price !== "") await updateCell(PRODUCTS_SHEET, rowNum, colIndex(headers, "Price (Rs)") + 1, Number(updates.price));
  if (updates.stock !== undefined && updates.stock !== "") await updateCell(PRODUCTS_SHEET, rowNum, colIndex(headers, "Stock") + 1, Number(updates.stock));
  if (updates.desc !== undefined) await updateCell(PRODUCTS_SHEET, rowNum, colIndex(headers, "Description") + 1, updates.desc);
  if (updates.images && updates.images.length) {
    const urls = updates.images.map(img => `data:${img.mime};base64,${img.base64}`).filter(u => u.length < 45000);
    if (urls.length) await updateCell(PRODUCTS_SHEET, rowNum, colIndex(headers, "Image URLs") + 1, urls.join("|"));
  }
  return { success: true };
}
async function deleteProductRow(productId) {
  const found = await findProductRow(productId);
  if (!found) return { success: false, message: "Product not found." };
  await deleteRow(PRODUCTS_SHEET, found.rowNum);
  return { success: true };
}
async function updateProductAsSeller(email, password, productId, updates) {
  const auth = await authenticateSeller(email, password);
  if (!auth.success) return auth;
  const found = await findProductRow(productId);
  if (!found) return { success: false, message: "Product not found." };
  const sellerIdCol = colIndex(found.headers, "Seller ID");
  if (found.rowValues[sellerIdCol] !== auth.seller.id) return { success: false, message: "You can only edit your own listings." };
  return applyProductUpdate(productId, updates);
}
async function deleteProductAsSeller(email, password, productId) {
  const auth = await authenticateSeller(email, password);
  if (!auth.success) return auth;
  const found = await findProductRow(productId);
  if (!found) return { success: false, message: "Product not found." };
  const sellerIdCol = colIndex(found.headers, "Seller ID");
  if (found.rowValues[sellerIdCol] !== auth.seller.id) return { success: false, message: "You can only delete your own listings." };
  return deleteProductRow(productId);
}
async function updateProductAsAdmin(password, productId, updates) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  return applyProductUpdate(productId, updates);
}
async function deleteProductAsAdmin(password, productId) {
  if (password !== ADMIN_PASSWORD) return { success: false, message: "Incorrect admin password." };
  return deleteProductRow(productId);
}

// ---------- FUNCTION DISPATCH TABLE (mirrors Code.gs function names) ----------
const FUNCTIONS = {
  getProducts, verifySellerPassword, getApprovedSellers, addProduct, addProductAsSeller,
  registerSeller, verifyAdminPassword, getAllSellers, updateSellerStatus, sellerLogin,
  getSellerListings, getSellerOrders, calculateOrder, getUpiLink, placeOrder, submitUtr,
  getOrders, updateOrderStatus, updatePaymentStatus, updateProductAsSeller, deleteProductAsSeller,
  updateProductAsAdmin, deleteProductAsAdmin
};

app.post("/api/call/:fn", async (req, res) => {
  const fn = FUNCTIONS[req.params.fn];
  if (!fn) return res.status(404).json({ __error: "Unknown function: " + req.params.fn });
  try {
    const args = Array.isArray(req.body) ? req.body : [];
    const result = await fn(...args);
    res.json(result === undefined ? null : result);
  } catch (err) {
    console.error(req.params.fn, err);
    res.status(500).json({ __error: err.message || String(err) });
  }
});

app.get("/healthz", (req, res) => res.send("ok"));
app.get("/__debug/files", (req, res) => {
  const fs = require("fs");
  function walk(dir, depth) {
    if (depth > 3) return [];
    let out = [];
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return [dir + " (unreadable: " + e.message + ")"]; }
    for (const e of entries) {
      if (e.name === "node_modules" || e.name === ".git") continue;
      const p = dir + "/" + e.name;
      out.push(p + (e.isDirectory() ? "/" : ""));
      if (e.isDirectory()) out = out.concat(walk(p, depth + 1));
    }
    return out;
  }
  res.type("text/plain").send(walk(__dirname, 0).join("\n"));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("POGT real backend listening on " + PORT));
