// ============================================================
// DAILY PRICING ANALYSIS - doGet() API
// Add this script to MASTER CSV Google Sheet (Extensions > Apps Script)
// Deploy as Web App (Execute as: Me, Access: Anyone)
// ============================================================

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
    if (!sheet) throw new Error('Sheet "MASTER CSV" not found');

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      return jsonResp({ summary: {}, byPlatform: {}, byCompany: {}, byMonth: [], byCategory: {}, timestamp: new Date().toISOString() });
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headers.length; i++) headers[i] = String(headers[i] || '').trim();

    // Column index mapping (0-based)
    var COL = {};
    var colNames = {
      'UNIQUE': 'unique', 'S.NO': 'sno', 'YEAR': 'year', 'MONTH': 'month', 'DATE': 'date',
      'PLATFORM': 'platform', 'COMPANY': 'company', 'ORDER ID': 'orderId',
      'TRACKING ID': 'trackingId', 'COMPANY ID': 'companyId', 'SKU ID': 'skuId',
      'SUB ORDER ID': 'subOrderId', 'QUANTITY': 'quantity', 'COURIER': 'courier',
      'MASTER SKU': 'masterSku', 'SINGLE-COMBO': 'singleCombo', 'CATEGORY': 'category',
      'SUB CATEGORY': 'subCategory', 'COLOR': 'color', 'PRODUCT': 'product',
      'LINK': 'link', 'REQUIRED': 'required', 'GROSS RECEIVED': 'grossReceived',
      'TAX %': 'taxPct', 'TAX AMOUNT': 'taxAmount', 'NET RECEIVED': 'netReceived',
      'COP': 'cop', 'P/L': 'pl', 'ACTUAL PRICE': 'actualPrice',
      'SELLING PRICE': 'sellingPrice', 'ESTIMADED SELLING PRICE': 'estSellingPrice'
    };
    for (var i = 0; i < headers.length; i++) {
      var hUpper = headers[i].toUpperCase();
      for (var key in colNames) {
        if (hUpper === key || hUpper.indexOf(key) >= 0) {
          COL[colNames[key]] = i;
          break;
        }
      }
    }

    // Read all data in one batch
    var numRows = lastRow - 1;
    var rawData = sheet.getRange(2, 1, numRows, lastCol).getValues();

    // Aggregation
    var summary = { totalOrders: 0, totalQty: 0, grossReceived: 0, netReceived: 0, taxAmount: 0, cop: 0, pl: 0, avgOrderValue: 0 };
    var byPlatform = {};
    var byCompany = {};
    var byMonthMap = {};
    var byCategory = {};
    var bySku = {};
    var dailyMap = {};

    for (var r = 0; r < rawData.length; r++) {
      var row = rawData[r];
      if (!row[0] || String(row[0]).trim() === '') continue;

      var qty = toNum(row[COL.quantity] !== undefined ? row[COL.quantity] : 0);
      var gross = toNum(row[COL.grossReceived] !== undefined ? row[COL.grossReceived] : 0);
      var net = toNum(row[COL.netReceived] !== undefined ? row[COL.netReceived] : 0);
      var tax = toNum(row[COL.taxAmount] !== undefined ? row[COL.taxAmount] : 0);
      var cop = toNum(row[COL.cop] !== undefined ? row[COL.cop] : 0);
      var pl = toNum(row[COL.pl] !== undefined ? row[COL.pl] : 0);
      var platform = String(row[COL.platform] || '').trim() || 'Unknown';
      var company = String(row[COL.company] || '').trim() || 'Unknown';
      var category = String(row[COL.category] || '').trim() || 'Unknown';
      var sku = String(row[COL.masterSku] || '').trim() || 'Unknown';

      // Date handling
      var dateVal = row[COL.date];
      var dateStr = '';
      var monthKey = '';
      if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        monthKey = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM');
      } else if (dateVal) {
        var parsed = new Date(dateVal);
        if (!isNaN(parsed.getTime())) {
          dateStr = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          monthKey = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM');
        }
      }
      if (!monthKey) {
        var yr = String(row[COL.year] || '').trim();
        var mo = String(row[COL.month] || '').trim();
        if (yr && mo) monthKey = yr + '-' + monthToNum(mo);
      }

      // Summary totals
      summary.totalOrders++;
      summary.totalQty += qty;
      summary.grossReceived += gross;
      summary.netReceived += net;
      summary.taxAmount += tax;
      summary.cop += cop;
      summary.pl += pl;

      // By Platform
      if (!byPlatform[platform]) byPlatform[platform] = { orders: 0, qty: 0, gross: 0, net: 0, tax: 0, cop: 0, pl: 0 };
      byPlatform[platform].orders++;
      byPlatform[platform].qty += qty;
      byPlatform[platform].gross += gross;
      byPlatform[platform].net += net;
      byPlatform[platform].tax += tax;
      byPlatform[platform].cop += cop;
      byPlatform[platform].pl += pl;

      // By Company
      if (!byCompany[company]) byCompany[company] = { orders: 0, qty: 0, gross: 0, net: 0, tax: 0, cop: 0, pl: 0 };
      byCompany[company].orders++;
      byCompany[company].qty += qty;
      byCompany[company].gross += gross;
      byCompany[company].net += net;
      byCompany[company].tax += tax;
      byCompany[company].cop += cop;
      byCompany[company].pl += pl;

      // By Month
      if (monthKey) {
        if (!byMonthMap[monthKey]) byMonthMap[monthKey] = { month: monthKey, orders: 0, qty: 0, gross: 0, net: 0, tax: 0, cop: 0, pl: 0 };
        byMonthMap[monthKey].orders++;
        byMonthMap[monthKey].qty += qty;
        byMonthMap[monthKey].gross += gross;
        byMonthMap[monthKey].net += net;
        byMonthMap[monthKey].tax += tax;
        byMonthMap[monthKey].cop += cop;
        byMonthMap[monthKey].pl += pl;
      }

      // By Category
      if (!byCategory[category]) byCategory[category] = { orders: 0, qty: 0, gross: 0, net: 0, cop: 0, pl: 0 };
      byCategory[category].orders++;
      byCategory[category].qty += qty;
      byCategory[category].gross += gross;
      byCategory[category].net += net;
      byCategory[category].cop += cop;
      byCategory[category].pl += pl;

      // Daily (last 60 days)
      if (dateStr) {
        if (!dailyMap[dateStr]) dailyMap[dateStr] = { date: dateStr, orders: 0, qty: 0, gross: 0, net: 0, cop: 0, pl: 0 };
        dailyMap[dateStr].orders++;
        dailyMap[dateStr].qty += qty;
        dailyMap[dateStr].gross += gross;
        dailyMap[dateStr].net += net;
        dailyMap[dateStr].cop += cop;
        dailyMap[dateStr].pl += pl;
      }

      // Top SKUs (aggregate)
      if (!bySku[sku]) bySku[sku] = { sku: sku, platform: platform, orders: 0, qty: 0, gross: 0, net: 0, cop: 0, pl: 0 };
      bySku[sku].orders++;
      bySku[sku].qty += qty;
      bySku[sku].gross += gross;
      bySku[sku].net += net;
      bySku[sku].cop += cop;
      bySku[sku].pl += pl;
    }

    summary.avgOrderValue = summary.totalOrders > 0 ? summary.grossReceived / summary.totalOrders : 0;
    summary.margin = summary.grossReceived > 0 ? (summary.pl / summary.grossReceived * 100) : 0;

    // Sort months
    var byMonth = [];
    var monthKeys = Object.keys(byMonthMap).sort();
    for (var i = 0; i < monthKeys.length; i++) byMonth.push(byMonthMap[monthKeys[i]]);

    // Sort daily and take last 60
    var daily = [];
    var dailyKeys = Object.keys(dailyMap).sort();
    var startIdx = Math.max(0, dailyKeys.length - 60);
    for (var i = startIdx; i < dailyKeys.length; i++) daily.push(dailyMap[dailyKeys[i]]);

    // Top 20 SKUs by revenue
    var skuList = [];
    for (var k in bySku) skuList.push(bySku[k]);
    skuList.sort(function(a, b) { return b.gross - a.gross; });
    var topSkus = skuList.slice(0, 20);

    // Platform list for filters
    var platforms = Object.keys(byPlatform).sort();
    var companies = Object.keys(byCompany).sort();

    // Round all numbers
    roundObj(summary);
    for (var k in byPlatform) roundObj(byPlatform[k]);
    for (var k in byCompany) roundObj(byCompany[k]);
    for (var i = 0; i < byMonth.length; i++) roundObj(byMonth[i]);
    for (var k in byCategory) roundObj(byCategory[k]);
    for (var i = 0; i < daily.length; i++) roundObj(daily[i]);
    for (var i = 0; i < topSkus.length; i++) roundObj(topSkus[i]);

    return jsonResp({
      summary: summary,
      byPlatform: byPlatform,
      byCompany: byCompany,
      byMonth: byMonth,
      byCategory: byCategory,
      daily: daily,
      topSkus: topSkus,
      platforms: platforms,
      companies: companies,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    return jsonResp({ error: err.message });
  }
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function roundObj(obj) {
  for (var k in obj) {
    if (typeof obj[k] === 'number') {
      obj[k] = Math.round(obj[k] * 100) / 100;
    }
  }
}

function monthToNum(m) {
  var months = { 'january': '01', 'february': '02', 'march': '03', 'april': '04', 'may': '05', 'june': '06',
                 'july': '07', 'august': '08', 'september': '09', 'october': '10', 'november': '11', 'december': '12' };
  var ml = String(m).toLowerCase().trim();
  return months[ml] || '01';
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Financial Dashboard')
    .addItem('Test API', 'testApi')
    .addToUi();
}

function testApi() {
  var result = doGet();
  var json = JSON.parse(result.getContent());
  Logger.log('Total Orders: ' + json.summary.totalOrders);
  Logger.log('Gross Received: ' + json.summary.grossReceived);
  Logger.log('P/L: ' + json.summary.pl);
  Logger.log('Platforms: ' + json.platforms.join(', '));
  Logger.log('Months: ' + json.byMonth.length);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Orders: ' + json.summary.totalOrders + ' | Revenue: ₹' + json.summary.grossReceived.toLocaleString(),
    'API Test OK', 10
  );
}
