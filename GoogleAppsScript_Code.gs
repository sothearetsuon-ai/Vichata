/**
 * ==============================================================================
 * 🎂 ប្រព័ន្ធគ្រប់គ្រង និងស៊ីងទិន្នន័យហាងនំខេកវិជ្ជដា (Bakery POS Google Apps Script)
 * ==============================================================================
 * 
 * របៀបដំឡើង (Setup Instructions):
 * 1. បើក Google Sheet របស់អ្នក
 * 2. ចុច Extensions (ផ្នែកបន្ថែម) > Apps Script
 * 3. លុបកូដចាស់ទាំងអស់ចេញ ហើយ Copy & Paste កូដខាងក្រោមនេះទាំងស្រុង
 * 4. ចុច Save (រូបផ្លែថាស)
 * 5. ចុច Deploy (ដាក់ឱ្យប្រើប្រាស់) > New deployment (ការដាក់ឱ្យប្រើប្រាស់ថ្មី)
 * 6. ជ្រើសរើស Select type (រូបកង់ធ្មេញ) > Web app
 * 7. កំណត់ដូចខាងក្រោម៖
 *    - Description: Bakery POS API v2
 *    - Execute as: Me (អាសយដ្ឋានអ៊ីមែលរបស់អ្នក)
 *    - Who has access: Anyone (អ្នកណាក៏បាន - សំខាន់បំផុត!)
 * 8. ចុច Deploy > Authorize access (ផ្តល់សិទ្ធិ) > Advanced > Go to Untitled (unsafe) > Allow
 * 9. Copy "Web app URL" (ដែលមានទម្រង់ https://script.google.com/macros/s/.../exec)
 * 10. យក URL នោះទៅបិទភ្ជាប់ (Paste) ក្នុងប្រអប់ Settings នៃកម្មវិធី POS របស់អ្នកជាការស្រេច!
 */

const SHEET_SALES = "Sales";
const SHEET_PRODUCTS = "Products";
const SHEET_EXPENSES = "Expenses";
const TELEGRAM_BOT_TOKEN = "8544004471:AAFsyBJbtmKR6W46zGtqGhNVWjE2AIvzWC4";

// ==============================================================================
// 1. GET REQUEST (ទាញទិន្នន័យពី Sheet ទៅកាន់ POS)
// ==============================================================================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNotExist(ss);
    
    const sales = getSalesData(ss);
    const products = getProductsData(ss);
    const expenses = getExpensesData(ss);
    
    const result = {
      status: "success",
      sales: sales,
      products: products,
      expenses: expenses,
      timestamp: new Date().toISOString()
    };
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==============================================================================
// 2. POST REQUEST (ទទួលការលក់, ចំណាយ, និងទំនិញពី POS ចូល Sheet & Telegram Webhook)
// ==============================================================================
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheetsIfNotExist(ss);
    
    let data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    // --- ករណីពិសេស: TELEGRAM WEBHOOK (ទទួលសារ និង Menu ពី Telegram) ---
    if (data.message || data.callback_query) {
      handleTelegramWebhook(ss, data);
      return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const action = data.action || "";
    
    // --- ករណី ១: កត់ត្រា ឬកែប្រែការលក់ (New or Edit Sale) ---
    if (action === "new_sale" || action === "edit_sale" || action === "update_sale") {
      const inv = data.invoice || data;
      const res = appendSaleRow(ss, inv);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- ករណី ២: បញ្ជូនការលក់ជាច្រើនក្នុងពេលតែមួយ (Batch Sales Sync) ---
    if (action === "sync_sales" || action === "batch_sales") {
      const list = data.sales || data.transactions || [];
      const res = batchAppendSales(ss, list);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- ករណី ៣: កត់ត្រាចំណាយ (Add Expense) ---
    if (action === "add_expense") {
      const exp = data.expense || data;
      const res = appendExpenseRow(ss, exp);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- ករណី ៤: ស៊ីងចំណាយទាំងអស់ (Sync Expenses) ---
    if (action === "sync_expenses") {
      const expList = data.expenses || [];
      const res = syncAllExpenses(ss, expList);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // --- ករណី ៥: ស៊ីងទំនិញ ឬកែសម្រួលស្តុក (Sync Products / Update Stock) ---
    if (action === "sync_products" || action === "update_stock" || action === "sync_all") {
      // សម្គាល់៖ មិនបង្កើតជួរលក់ក្លែងក្លាយជាដាច់ខាត! ដំណើរការតែ Products ប៉ុណ្ណោះ
      const prodList = (data.products || []).concat(data.ingredients || []);
      let res = { status: "success", message: "Products updated" };
      if (prodList.length > 0) {
        res = syncAllProducts(ss, prodList);
      }
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // បើសិនជាមាន invoice ច្បាស់លាស់ ទើបកត់ត្រាចូលការលក់
    if (data.invoice) {
      const res = appendSaleRow(ss, data.invoice);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "No action performed"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==============================================================================
// 3. HELPER FUNCTIONS: SALES (ការលក់)
// ==============================================================================
function appendSaleRow(ss, inv) {
  const sheet = ss.getSheetByName(SHEET_SALES);
  const id = String(inv.id || ("BK-" + Date.now()));
  
  const dateStr = inv.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const timeStr = inv.time || Utilities.formatDate(new Date(), "GMT+7", "HH:mm");
  const custName = inv.customerName || "អតិថិជនទូទៅ";
  const custPhone = inv.customerPhone || "គ្មាន";
  const itemsSummary = inv.itemsSummary || "";
  const itemsCount = Number(inv.itemsCount || 1);
  const totalKHR = Number(inv.totalKHR || 0);
  const depositKHR = Number((inv.depositKHR !== undefined && inv.depositKHR !== null) ? inv.depositKHR : totalKHR);
  const remainingKHR = Number(inv.remainingKHR || 0);
  const status = inv.status || (remainingKHR > 0 ? "នៅខ្វះ" : "Paid (គ្រប់)");
  const method = inv.method || "ប្រាក់សុទ្ធ (៛)";
  const note = inv.note || "";
  const photoLink = inv.photoLink || "";

  // ត្រួតពិនិត្យមើលថាតើវិក្កយបត្រនេះមានរួចហើយឬនៅ
  const existingIds = getExistingColumnValues(sheet, 1);
  const existingIndex = existingIds.indexOf(id);
  
  if (existingIndex !== -1) {
    // ប្រសិនបើមានរួចហើយ (ករណីកែប្រែវិក្កយបត្រ) ធ្វើការ Update ជួរចាស់នោះ
    const targetRow = existingIndex + 2; // បូក ២ ព្រោះជួរទី ១ ជា Header
    sheet.getRange(targetRow, 1, 1, 14).setValues([[
      id,
      dateStr,
      timeStr,
      custName,
      custPhone,
      itemsSummary,
      itemsCount,
      totalKHR,
      depositKHR,
      remainingKHR,
      status,
      method,
      note,
      photoLink
    ]]);
    return { status: "success", id: id, updated: true, message: "Invoice updated successfully" };
  }
  
  sheet.appendRow([
    id,
    dateStr,
    timeStr,
    custName,
    custPhone,
    itemsSummary,
    itemsCount,
    totalKHR,
    depositKHR,
    remainingKHR,
    status,
    method,
    note,
    photoLink
  ]);
  
  return { status: "success", id: id, photoUrl: photoLink };
}

function batchAppendSales(ss, salesList) {
  if (!Array.isArray(salesList) || salesList.length === 0) {
    return { status: "success", count: 0 };
  }
  
  const sheet = ss.getSheetByName(SHEET_SALES);
  const existingIds = getExistingColumnValues(sheet, 1);
  const existingSet = new Set(existingIds);
  
  const rowsToAdd = [];
  salesList.forEach(inv => {
    const id = String(inv.id || ("BK-" + Date.now()));
    if (!existingSet.has(id)) {
      existingSet.add(id);
      rowsToAdd.push([
        id,
        inv.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"),
        inv.time || Utilities.formatDate(new Date(), "GMT+7", "HH:mm"),
        inv.customerName || "អតិថិជនទូទៅ",
        inv.customerPhone || "គ្មាន",
        inv.itemsSummary || "",
        Number(inv.itemsCount || 1),
        Number(inv.totalKHR || 0),
        Number(inv.depositKHR || (inv.totalKHR || 0)),
        Number(inv.remainingKHR || 0),
        inv.status || "Paid (គ្រប់)",
        inv.method || "ប្រាក់សុទ្ធ (៛)",
        inv.note || "",
        inv.photoLink || ""
      ]);
    }
  });
  
  if (rowsToAdd.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 14).setValues(rowsToAdd);
  }
  
  return { status: "success", addedCount: rowsToAdd.length };
}

function getSalesData(ss) {
  const sheet = ss.getSheetByName(SHEET_SALES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  return values.map(row => ({
    id: String(row[0] || ""),
    date: row[1] ? (row[1] instanceof Date ? Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd") : String(row[1])) : "",
    time: row[2] ? (row[2] instanceof Date ? Utilities.formatDate(row[2], "GMT+7", "HH:mm") : String(row[2])) : "",
    customerName: String(row[3] || "អតិថិជនទូទៅ"),
    customerPhone: String(row[4] || "គ្មាន"),
    itemsSummary: String(row[5] || ""),
    itemsCount: Number(row[6] || 1),
    totalKHR: Number(row[7] || 0),
    depositKHR: Number(row[8] || 0),
    remainingKHR: Number(row[9] || 0),
    status: String(row[10] || "Paid (គ្រប់)"),
    method: String(row[11] || "ប្រាក់សុទ្ធ (៛)"),
    note: String(row[12] || ""),
    photoLink: String(row[13] || "")
  })).filter(s => s.id !== "");
}

// ==============================================================================
// 4. HELPER FUNCTIONS: EXPENSES (ចំណាយ)
// ==============================================================================
function appendExpenseRow(ss, exp) {
  const sheet = ss.getSheetByName(SHEET_EXPENSES);
  const id = String(exp.id || ("EXP-" + Date.now()));
  const dateStr = exp.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const title = exp.title || "ចំណាយទូទៅ";
  const cat = exp.cat || "other";
  const amount = Number(exp.amount || 0);
  const note = exp.note || "";
  
  sheet.appendRow([id, dateStr, title, cat, amount, note]);
  return { status: "success", id: id };
}

function syncAllExpenses(ss, expList) {
  const sheet = ss.getSheetByName(SHEET_EXPENSES);
  sheet.clearContents();
  sheet.appendRow(["ID", "កាលបរិច្ឆេទ", "មុខចំណាយ", "ប្រភេទ", "ចំនួនទឹកប្រាក់ (៛)", "កំណត់ចំណាំ"]);
  
  if (expList.length > 0) {
    const rows = expList.map(exp => [
      String(exp.id || ("EXP-" + Date.now())),
      exp.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"),
      exp.title || "",
      exp.cat || "other",
      Number(exp.amount || 0),
      exp.note || ""
    ]);
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  return { status: "success", count: expList.length };
}

function getExpensesData(ss) {
  const sheet = ss.getSheetByName(SHEET_EXPENSES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return values.map(row => ({
    id: String(row[0] || ""),
    date: row[1] ? (row[1] instanceof Date ? Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd") : String(row[1])) : "",
    title: String(row[2] || ""),
    cat: String(row[3] || "other"),
    amount: Number(row[4] || 0),
    note: String(row[5] || "")
  })).filter(e => e.id !== "");
}

// ==============================================================================
// 5. HELPER FUNCTIONS: PRODUCTS (ទំនិញ & គ្រឿងផ្សំ)
// ==============================================================================
function syncAllProducts(ss, prodList) {
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  sheet.clearContents();
  sheet.appendRow(["ID", "ឈ្មោះទំនិញ", "ប្រភេទ", "ស្លាក", "តម្លៃលក់ (៛)", "ស្តុកនៅសល់", "តំណភ្ជាប់រូបភាព"]);
  
  if (prodList.length > 0) {
    const rows = prodList.map(p => [
      String(p.id || ""),
      p.name || "",
      p.cat || "cakes",
      p.badge || "",
      Number(p.priceKHR || 0),
      Number(p.stock !== undefined ? p.stock : 10),
      (p.img && typeof p.img === 'string' && p.img.startsWith('http')) ? p.img : ""
    ]);
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return { status: "success", count: prodList.length };
}

function getProductsData(ss) {
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  return values.map(row => ({
    id: isNaN(row[0]) ? String(row[0]) : Number(row[0]),
    name: String(row[1] || ""),
    cat: String(row[2] || "cakes"),
    badge: String(row[3] || ""),
    priceKHR: Number(row[4] || 0),
    stock: Number(row[5] !== undefined ? row[5] : 0),
    img: String(row[6] || "")
  })).filter(p => p.name !== "");
}

// ==============================================================================
// 6. INITIALIZATION & UTILITIES
// ==============================================================================
function initSheetsIfNotExist(ss) {
  let sales = ss.getSheetByName(SHEET_SALES);
  if (!sales) {
    sales = ss.insertSheet(SHEET_SALES);
    sales.appendRow([
      "ID វិក្កយបត្រ", "កាលបរិច្ឆេទ", "ម៉ោង", "ឈ្មោះអតិថិជន", "លេខទូរស័ព្ទ",
      "មុខទំនិញដែលបានទិញ", "ចំនួនសរុប", "ទឹកប្រាក់សរុប (៛)", "ប្រាក់បានកក់/បង់ (៛)",
      "ប្រាក់នៅខ្វះ (៛)", "ស្ថានភាព", "វិធីទូទាត់", "កំណត់ចំណាំ", "តំណភ្ជាប់រូបភាព"
    ]);
    sales.getRange("A1:N1").setBackground("#E07A5F").setFontColor("#FFFFFF").setFontWeight("bold");
    sales.setFrozenRows(1);
  }
  
  let products = ss.getSheetByName(SHEET_PRODUCTS);
  if (!products) {
    products = ss.insertSheet(SHEET_PRODUCTS);
    products.appendRow(["ID", "ឈ្មោះទំនិញ", "ប្រភេទ", "ស្លាក", "តម្លៃលក់ (៛)", "ស្តុកនៅសល់", "តំណភ្ជាប់រូបភាព"]);
    products.getRange("A1:G1").setBackground("#3B82F6").setFontColor("#FFFFFF").setFontWeight("bold");
    products.setFrozenRows(1);
  }
  
  let expenses = ss.getSheetByName(SHEET_EXPENSES);
  if (!expenses) {
    expenses = ss.insertSheet(SHEET_EXPENSES);
    expenses.appendRow(["ID", "កាលបរិច្ឆេទ", "មុខចំណាយ", "ប្រភេទ", "ចំនួនទឹកប្រាក់ (៛)", "កំណត់ចំណាំ"]);
    expenses.getRange("A1:F1").setBackground("#EF4444").setFontColor("#FFFFFF").setFontWeight("bold");
    expenses.setFrozenRows(1);
  }
}

function getExistingColumnValues(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const vals = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  return vals.map(r => String(r[0]));
}

// ==============================================================================
// 7. TELEGRAM BOT WEBHOOK & INTERACTIVE MENU HANDLERS
// ==============================================================================
function handleTelegramWebhook(ss, data) {
  let chatId = null;
  let text = "";
  let isCallback = false;
  let callbackQueryId = null;

  if (data.message) {
    chatId = data.message.chat ? data.message.chat.id : null;
    text = (data.message.text || "").trim();
  } else if (data.callback_query) {
    isCallback = true;
    callbackQueryId = data.callback_query.id;
    chatId = data.callback_query.message && data.callback_query.message.chat ? data.callback_query.message.chat.id : null;
    text = (data.callback_query.data || "").trim();
  }

  if (!chatId) return;

  if (isCallback && callbackQueryId) {
    sendTelegramApi("answerCallbackQuery", { callback_query_id: callbackQueryId });
  }

  const cleanText = text.toLowerCase();

  // 1. MENU / START / HELP
  if (cleanText === "/start" || cleanText === "/menu" || cleanText === "/help" || cleanText === "menu" || text.includes("ជំនួយ") || text.includes("ម៉ឺនុយ") || text.includes("មេ")) {
    sendTelegramMainMenu(chatId);
    return;
  }

  // 2. SALES REPORT (របាយការណ៍លក់)
  if (cleanText === "/sales" || cleanText === "cmd_sales" || text.includes("របាយការណ៍លក់") || text.includes("ការលក់") || cleanText.includes("sale")) {
    sendTelegramSalesReport(ss, chatId);
    return;
  }

  // 3. STOCK REPORT (ពិនិត្យស្តុក)
  if (cleanText === "/stock" || cleanText === "cmd_stock" || text.includes("ពិនិត្យស្តុក") || text.includes("ស្តុក") || cleanText.includes("stock")) {
    sendTelegramStockReport(ss, chatId);
    return;
  }

  // 4. EXPENSE REPORT (ពិនិត្យចំណាយ)
  if (cleanText === "/expenses" || cleanText === "cmd_expenses" || text.includes("ពិនិត្យចំណាយ") || text.includes("ចំណាយ") || cleanText.includes("expense")) {
    sendTelegramExpensesReport(ss, chatId);
    return;
  }

  // 5. DAILY SUMMARY (សង្ខេបថ្ងៃនេះ / ចំណេញ-ខាត)
  if (cleanText === "/summary" || cleanText === "cmd_summary" || text.includes("សង្ខេប") || text.includes("ចំណេញ") || cleanText.includes("summary")) {
    sendTelegramSummaryReport(ss, chatId);
    return;
  }

  // DEFAULT FALLBACK
  sendTelegramMainMenu(chatId, "👋 ខ្ញុំបានទទួលសាររបស់អ្នក។ សូមជ្រើសរើស Menu ខាងក្រោមដើម្បីពិនិត្យរបាយការណ៍៖");
}

function sendTelegramMainMenu(chatId, customText) {
  const msg = customText || "👋 <b>សូមស្វាគមន៍មកកាន់ប្រព័ន្ធ Telegram Bot ហាងនំខេកវិជ្ជដា!</b> 🎂\n\nសូមចុចប៊ូតុង Menu ខាងក្រោម ដើម្បីពិនិត្យទិន្នន័យជាក់ស្តែង៖";
  
  const replyMarkup = {
    keyboard: [
      [{ text: "📊 របាយការណ៍លក់" }, { text: "📦 ពិនិត្យស្តុក" }],
      [{ text: "💸 ពិនិត្យចំណាយ" }, { text: "📈 សង្ខេបថ្ងៃនេះ" }]
    ],
    resize_keyboard: true,
    persistent: true
  };

  sendTelegramApi("sendMessage", {
    chat_id: chatId,
    text: msg,
    parse_mode: "HTML",
    reply_markup: replyMarkup
  });
}

function sendTelegramSalesReport(ss, chatId) {
  var sales = getSalesData(ss);
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var timeNow = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");

  var todaySales = sales.filter(function(s) { return s.date === today; });
  var totalKHR = todaySales.reduce(function(sum, s) { return sum + (s.totalKHR || 0); }, 0);
  var totalUSD = (totalKHR / 4000).toFixed(2);
  var cashSales = todaySales.filter(function(s) { return s.method && (s.method.indexOf("ប្រាក់សុទ្ធ") !== -1 || s.method.toLowerCase().indexOf("cash") !== -1); }).reduce(function(sum, s) { return sum + (s.totalKHR || 0); }, 0);
  var khqrSales = todaySales.filter(function(s) { return s.method && (s.method.indexOf("KHQR") !== -1 || s.method.toLowerCase().indexOf("bank") !== -1 || s.method.toLowerCase().indexOf("qr") !== -1); }).reduce(function(sum, s) { return sum + (s.totalKHR || 0); }, 0);
  var totalUnpaid = todaySales.reduce(function(sum, s) { return sum + (s.remainingKHR || 0); }, 0);

  var recentText = "";
  if (todaySales.length > 0) {
    var recent3 = todaySales.slice(-3).reverse();
    recentText = recent3.map(function(s, idx) {
      var cust = (s.customerName && s.customerName !== "អតិថិជនទូទៅ" && s.customerName !== "គ្មាន") ? " (" + s.customerName + ")" : "";
      return "  " + (idx + 1) + ". <b>#" + s.id + "</b>" + cust + " — <code>" + (s.totalKHR || 0).toLocaleString() + " ៛</code> [" + s.method + "]";
    }).join("\n");
  } else {
    recentText = "  <i>មិនទាន់មានការលក់នៅថ្ងៃនេះទេ</i>";
  }

  var text = "📊 <b>របាយការណ៍លក់ប្រចាំថ្ងៃ (" + today + ")</b>\n" +
    "🏪 <b>ហាងនំខេកវិជ្ជដា</b> • ម៉ោង៖ <code>" + timeNow + "</code>\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    "🧾 ចំនួនវិក្កយបត្រថ្ងៃនេះ៖ <b>" + todaySales.length + " វិក្កយបត្រ</b>\n" +
    "💰 ចំណូលសរុប៖ <b>" + totalKHR.toLocaleString() + " ៛</b> (≈ $" + totalUSD + ")\n" +
    "💵 សាច់ប្រាក់ (Cash)៖ <b>" + cashSales.toLocaleString() + " ៛</b>\n" +
    "📱 ស្កេន KHQR (Bank)៖ <b>" + khqrSales.toLocaleString() + " ៛</b>\n" +
    (totalUnpaid > 0 ? ("⚠️ ប្រាក់នៅខ្វះសរុប៖ <b>" + totalUnpaid.toLocaleString() + " ៛</b>\n") : "✅ បានទូទាត់គ្រប់ចំនួន ១០០%\n") +
    "━━━━━━━━━━━━━━━━━━\n" +
    "📋 <b>ការលក់ចុងក្រោយថ្ងៃនេះ៖</b>\n" + recentText;

  var inlineMarkup = {
    inline_keyboard: [
      [
        { text: "🔄 Refresh លក់", callback_data: "cmd_sales" },
        { text: "📦 ពិនិត្យស្តុក", callback_data: "cmd_stock" }
      ],
      [
        { text: "💸 ចំណាយ", callback_data: "cmd_expenses" },
        { text: "📈 សង្ខេបថ្ងៃនេះ", callback_data: "cmd_summary" }
      ]
    ]
  };

  sendTelegramApi("sendMessage", {
    chat_id: chatId,
    text: text,
    parse_mode: "HTML",
    reply_markup: inlineMarkup
  });
}

function sendTelegramStockReport(ss, chatId) {
  var products = getProductsData(ss);
  var timeNow = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");

  var lowStock = products.filter(function(p) { return Number(p.stock) <= 5; });
  var inStock = products.filter(function(p) { return Number(p.stock) > 5; });

  var lowStockText = "";
  if (lowStock.length > 0) {
    lowStockText = "⚠️ <b>ទំនិញជិតអស់ពីស្តុក (Low Stock):</b>\n" +
      lowStock.map(function(p) { return "  • <b>" + p.name + "</b>៖ នៅសល់ <b>" + p.stock + "</b> " + (p.stock <= 0 ? '(អស់)' : ''); }).join("\n");
  } else {
    lowStockText = "✅ <b>ទំនិញទាំងអស់មានស្តុកគ្រប់គ្រាន់!</b> (គ្មានទំនិញជិតអស់ទេ)";
  }

  var sampleText = "";
  var samples = products.slice(0, 5);
  if (samples.length > 0) {
    sampleText = "\n━━━━━━━━━━━━━━━━━━\n📋 <b>ស្តុកទំនិញមួយចំនួន៖</b>\n" +
      samples.map(function(p) { return "  • " + p.name + "៖ <code>" + p.stock + "</code> (លក់: " + (p.priceKHR || 0).toLocaleString() + "៛)"; }).join("\n");
  }

  var text = "📦 <b>របាយការណ៍ស្តុកទំនិញ (Inventory Status)</b>\n" +
    "🏪 <b>ហាងនំខេកវិជ្ជដា</b> • ម៉ោង៖ <code>" + timeNow + "</code>\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    "🔢 មុខទំនិញសរុប៖ <b>" + products.length + " មុខ</b>\n" +
    "✅ ទំនិញមានស្តុក៖ <b>" + inStock.length + " មុខ</b>\n" +
    "⚠️ ជិតអស់/អស់ស្តុក៖ <b>" + lowStock.length + " មុខ</b>\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    lowStockText + sampleText;

  var inlineMarkup = {
    inline_keyboard: [
      [
        { text: "🔄 Refresh ស្តុក", callback_data: "cmd_stock" },
        { text: "📊 របាយការណ៍លក់", callback_data: "cmd_sales" }
      ],
      [
        { text: "💸 ចំណាយ", callback_data: "cmd_expenses" },
        { text: "📈 សង្ខេបថ្ងៃនេះ", callback_data: "cmd_summary" }
      ]
    ]
  };

  sendTelegramApi("sendMessage", {
    chat_id: chatId,
    text: text,
    parse_mode: "HTML",
    reply_markup: inlineMarkup
  });
}

function sendTelegramExpensesReport(ss, chatId) {
  var expenses = getExpensesData(ss);
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var timeNow = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");

  var todayExpenses = expenses.filter(function(e) { return e.date === today; });
  var totalExpKHR = todayExpenses.reduce(function(sum, e) { return sum + (e.amount || 0); }, 0);
  var totalExpUSD = (totalExpKHR / 4000).toFixed(2);

  var expListText = "";
  if (todayExpenses.length > 0) {
    expListText = todayExpenses.map(function(e, idx) {
      return "  " + (idx + 1) + ". <b>" + (e.title || "ចំណាយ") + "</b> — <code>" + (e.amount || 0).toLocaleString() + " ៛</code> (" + (e.cat || "ទូទៅ") + ")";
    }).join("\n");
  } else {
    expListText = "  <i>មិនទាន់មានការកត់ត្រាចំណាយនៅថ្ងៃនេះទេ</i>";
  }

  var text = "💸 <b>របាយការណ៍ចំណាយប្រចាំថ្ងៃ (" + today + ")</b>\n" +
    "🏪 <b>ហាងនំខេកវិជ្ជដា</b> • ម៉ោង៖ <code>" + timeNow + "</code>\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    "🧾 ចំនួនប្រតិបត្តិការចំណាយ៖ <b>" + todayExpenses.length + " លើក</b>\n" +
    "📉 ចំណាយសរុបថ្ងៃនេះ៖ <b>" + totalExpKHR.toLocaleString() + " ៛</b> (≈ $" + totalExpUSD + ")\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    "📋 <b>បញ្ជីចំណាយថ្ងៃនេះ៖</b>\n" + expListText;

  var inlineMarkup = {
    inline_keyboard: [
      [
        { text: "🔄 Refresh ចំណាយ", callback_data: "cmd_expenses" },
        { text: "📊 របាយការណ៍លក់", callback_data: "cmd_sales" }
      ],
      [
        { text: "📦 ពិនិត្យស្តុក", callback_data: "cmd_stock" },
        { text: "📈 សង្ខេបថ្ងៃនេះ", callback_data: "cmd_summary" }
      ]
    ]
  };

  sendTelegramApi("sendMessage", {
    chat_id: chatId,
    text: text,
    parse_mode: "HTML",
    reply_markup: inlineMarkup
  });
}

function sendTelegramSummaryReport(ss, chatId) {
  var sales = getSalesData(ss);
  var expenses = getExpensesData(ss);
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var timeNow = Utilities.formatDate(new Date(), "GMT+7", "HH:mm");

  var todaySales = sales.filter(function(s) { return s.date === today; });
  var totalSalesKHR = todaySales.reduce(function(sum, s) { return sum + (s.totalKHR || 0); }, 0);

  var todayExpenses = expenses.filter(function(e) { return e.date === today; });
  var totalExpKHR = todayExpenses.reduce(function(sum, e) { return sum + (e.amount || 0); }, 0);

  var netProfit = totalSalesKHR - totalExpKHR;
  var netUSD = (netProfit / 4000).toFixed(2);
  var isProfitable = netProfit >= 0;

  var text = "📈 <b>សង្ខេបចំណូល-ចំណាយ និងប្រាក់ចំណេញ (" + today + ")</b>\n" +
    "🏪 <b>ហាងនំខេកវិជ្ជដា</b> • ម៉ោង៖ <code>" + timeNow + "</code>\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    "💰 ចំណូលលក់ (Sales)៖ <b>+" + totalSalesKHR.toLocaleString() + " ៛</b>\n" +
    "📉 ការចំណាយ (Expenses)៖ <b>-" + totalExpKHR.toLocaleString() + " ៛</b>\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    (isProfitable ? "🟢" : "🔴") + " <b>ប្រាក់ចំណេញសុទ្ធ (Net Profit)៖</b>\n" +
    "👉 <b>" + netProfit.toLocaleString() + " ៛</b> (≈ $" + netUSD + ")\n" +
    "━━━━━━━━━━━━━━━━━━\n" +
    "🧾 ការលក់សរុប៖ <b>" + todaySales.length + " វិក្កយបត្រ</b>\n" +
    "📑 ចំណាយសរុប៖ <b>" + todayExpenses.length + " លើក</b>";

  var inlineMarkup = {
    inline_keyboard: [
      [
        { text: "🔄 Refresh សង្ខេប", callback_data: "cmd_summary" },
        { text: "📊 លក់ថ្ងៃនេះ", callback_data: "cmd_sales" }
      ],
      [
        { text: "📦 ស្តុកទំនិញ", callback_data: "cmd_stock" },
        { text: "💸 ចំណាយថ្ងៃនេះ", callback_data: "cmd_expenses" }
      ]
    ]
  };

  sendTelegramApi("sendMessage", {
    chat_id: chatId,
    text: text,
    parse_mode: "HTML",
    reply_markup: inlineMarkup
  });
}

function sendTelegramApi(method, payload) {
  try {
    const url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/" + method;
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(url, options);
    return JSON.parse(res.getContentText());
  } catch (e) {
    Logger.log("Telegram API Error: " + e.toString());
    return { ok: false, error: e.toString() };
  }
}

