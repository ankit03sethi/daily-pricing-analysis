// ============================================================
// MASTER CSV - Apps Script (Sheets API - FAST)
// ============================================================
// Pulls columns A-N from 6 platform sheets, deduplicates by
// column A (keeps older entry).
// Columns O-U filled via SKU lookup from Master SKU FILE.
//   O=AC, P=AD, Q=AG, R=AJ, S=AL, T=AM, U=AN
// Column V = platform-specific price lookup × col M (Quantity).
// Column W filled via platform-specific lookup from each platform's source sheet.
// Column X filled via platform-specific TAX % lookup from PRICES sheet.
// Column Y = Tax Value = W - Z
// Column Z = Taxable Value = W / (100 + X) * 100
// Column AA = VLOOKUP(PRICES col F, key=result col O vs source col A) × col M (Quantity).
// Column AB = Z - AA
// Column AC filled from Meesho only (source col O).
//
// SETUP:
// 1. Paste in Extensions > Apps Script > Save
// 2. Click "Services" (+) on left > Add "Google Sheets API"
// 3. Reload sheet > use "Master CSV" menu
// ============================================================

var HEADERS = ['UNIQUE','S.NO','YEAR','MONTH','DATE','PLATFORM','COMPANY','ORDER ID','TRACKING ID','COMPANY ID','SKU ID','SUB ORDER ID','QUANTITY','COURIER'];

// Master SKU FILE for VLOOKUP (key = col AA, pull AC,AD,AG,AJ,AL,AM,AN)
var SKU_SHEET_ID = '1iSNEwTlqDWCBC0zYGLUUIJMl_F_mmtjlOwh4AlBjbXE';
var SKU_TAB_NAME = 'Master SKU FILE';

// PRICES sheet for platform-specific price lookup (key = col A)
var PRICES_SHEET_ID = '1MIdAFRSA1CLN3QYJX-b3DYVYdOfJXNzRTsHAd8HXPLA';
var PRICES_TAB_NAME = 'PRICES';
// Platform → price column in PRICES sheet
var PRICE_COLS = {
  'Firstcry':   'K',
  'FIRSTCRY':   'K',
  'Flipkart':   'U',
  'FLIPKART':   'U',
  'Meesho':     'AA',
  'MEESHO':     'AA',
  'Amazon':     'AF',
  'AMAZON':     'AF',
  'Amazon FBA': 'AF',
  'AMAZON FBA': 'AF',
  'Myntra':     'AK',
  'MYNTRA':     'AK'
};

// Column W sources — each platform has its own source sheet/col
// Myntra key = F+G+K in result (not col A), others key = col A in both
var MULTI_SHEET_ID = '1dUD8ryR3F3T6Em_CqE9Xd1n9EH57SVW52RAIoxUci-4';
var COL_V_SOURCES = {
  'MEESHO':     { id: '1kWKcsfIoc6aV4dhRXWhGisztwAnbRwvbFWc4kDhrlbk', tab: 'MEESHO MENFEST', valCol: 'N', abCol: 'O' },
  'FLIPKART':   { id: '13yoFXec4DgR30gEYRBgEf8_L7Zvf4SD45WQR6pwIWx8', tab: 'FLIPKART CSV',   valCol: 'X' },
  'FIRSTCRY':   { id: MULTI_SHEET_ID, tab: 'FIRSTCRY', valCol: 'Y' },
  'AMAZON':     { id: MULTI_SHEET_ID, tab: 'AMAZON',   valCol: 'S' },
  'AMAZON FBA': { id: MULTI_SHEET_ID, tab: 'AMAZON',   valCol: 'S' },
  'MYNTRA':     { id: MULTI_SHEET_ID, tab: 'MYNTRA',   valCol: 'K' }
};

var PLATFORMS = {
  amazonFBA: {
    id: '1n_CYru50cdpYNY343KJCN5Q_EXOjbS8dHcjQlNWvKpQ',
    tab: 'AMAZON FBA CSV',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','AY2:AY','BE2:BE','U2:U','L2:L','W2:W','BD2:BD']
  },
  myntra: {
    id: '1GWmarRmUDb5TFqSrnLHq3_tupY1O2afDDwGVXA8jbLM',
    tab: 'MYNTRA CSV',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','K2:K','Y2:Y','P2:P','S2:S','T2:T','AA2:AA','Z2:Z']
  },
  amazon: {
    id: '1jvELY2wMnzznmJe5KXHUBzVzrdU7NUI-h2m9KY7LpAg',
    tab: 'AMAZON CSV',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','I2:I','K2:K','L2:L','T2:T','N2:N','AC2:AC']
  },
  flipkart: {
    id: '13yoFXec4DgR30gEYRBgEf8_L7Zvf4SD45WQR6pwIWx8',
    tab: 'FLIPKART CSV',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','K2:K','AL2:AL','O2:O','P2:P','J2:J','Z2:Z','AS2:AS']
  },
  meesho: {
    id: '1kWKcsfIoc6aV4dhRXWhGisztwAnbRwvbFWc4kDhrlbk',
    tab: 'MEESHO MENFEST',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','J2:J','R2:R','L2:L','S2:S','M2:M','K2:K']
  },
  firstcry: {
    id: '167d95u9q2C7FCYR2IvKQFLkshKSEgMenWht3VKqnt9c',
    tab: 'FIRSTCRY CSV',
    cols: ['A2:A','B2:B','C2:C','D2:D','E2:E','F2:F','G2:G','H2:H','I2:I','L2:L','M2:M','J2:J','T2:T','W2:W']
  }
};

// ===================== MENU =====================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Master CSV')
    .addItem('>>> RUN EVERYTHING <<<', 'runEverything')
    .addItem('>>> CLEAR & RUN FRESH <<<', 'clearAndRun')
    .addSeparator()
    .addItem('Sync All Platforms', 'syncAll')
    .addItem('Run All Fills (O-AB, AC)', 'runAllFills')
    .addSeparator()
    .addItem('Sync Amazon FBA', 'sAmazonFBA')
    .addItem('Sync Amazon', 'sAmazon')
    .addItem('Sync Flipkart', 'sFlipkart')
    .addItem('Sync Meesho', 'sMeesho')
    .addItem('Sync Myntra', 'sMyntra')
    .addItem('Sync Firstcry', 'sFirstcry')
    .addSeparator()
    .addItem('Fill SKU Data (O-U)', 'fillSkuData')
    .addItem('Fill Prices (V)', 'fillPrices')
    .addItem('Fill Col W & AC', 'fillColV')
    .addItem('Fill Tax % (X)', 'fillTaxPct')
    .addItem('Calc Tax Y-Z', 'calcTaxValues')
    .addItem('Fill Col AA (COP)', 'fillColZ')
    .addItem('Calc AB (Z-AA)', 'calcColAA')
    .addItem('Remove Duplicates', 'removeDuplicates')
    .addSeparator()
    .addItem('Start Auto-Refresh (1 min)', 'setupAutoRefresh')
    .addItem('Stop Auto-Refresh', 'stopAutoRefresh')
    .addSeparator()
    .addItem('Setup Sheet', 'setupSheet')
    .addToUi();

  ui.createMenu('Financial Dashboard')
    .addItem('Test API', 'testFinancialApi')
    .addToUi();
}

function sAmazonFBA() { syncOne('amazonFBA'); }
function sAmazon()    { syncOne('amazon'); }
function sFlipkart()  { syncOne('flipkart'); }
function sMeesho()    { syncOne('meesho'); }
function sMyntra()    { syncOne('myntra'); }
function sFirstcry()  { syncOne('firstcry'); }

// ===================== FAST READ (Sheets API batchGet) =====================

function fastRead(sheetId, tabName, colRanges) {
  var ranges = [];
  for (var i = 0; i < colRanges.length; i++) {
    ranges.push("'" + tabName + "'!" + colRanges[i]);
  }
  var res = Sheets.Spreadsheets.Values.batchGet(sheetId, {ranges: ranges});
  var vr = res.valueRanges;

  var maxRows = 0;
  for (var v = 0; v < vr.length; v++) {
    var len = (vr[v].values || []).length;
    if (len > maxRows) maxRows = len;
  }

  var rows = [];
  for (var r = 0; r < maxRows; r++) {
    var row = [];
    for (var c = 0; c < vr.length; c++) {
      var vals = vr[c].values || [];
      row.push(r < vals.length && vals[r].length > 0 ? vals[r][0] : '');
    }
    rows.push(row);
  }
  return rows;
}

// ===================== SYNC ONE PLATFORM =====================

function syncOne(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('Run Setup Sheet first.'); return; }

  var p = PLATFORMS[name];
  var existing = getExisting(ms);

  var raw;
  try {
    raw = fastRead(p.id, p.tab, p.cols);
  } catch(e) {
    SpreadsheetApp.getUi().alert(name + ' error: ' + e.message);
    return;
  }

  var seen = {};
  var newRows = [];

  for (var i = 0; i < raw.length; i++) {
    var key = String(raw[i][0]).trim();

    // Build UNIQUE if empty
    if (!key) {
      var oid = String(raw[i][7]).trim();
      if (!oid) continue;
      key = String(raw[i][5]).trim() + String(raw[i][6]).trim() + oid + String(raw[i][10]).trim();
      raw[i][0] = key;
    }

    // Build S.NO if empty
    if (!String(raw[i][1]).trim()) {
      var tid = String(raw[i][8]).trim();
      if (tid.endsWith('.0')) tid = tid.slice(0, -2);
      raw[i][1] = String(raw[i][5]).trim() + String(raw[i][6]).trim() + tid;
    }

    // Skip duplicates (keep older = first)
    if (existing[key] || seen[key]) continue;
    seen[key] = true;
    newRows.push(raw[i]);
  }

  if (newRows.length > 0) {
    var lr = Math.max(ms.getLastRow(), 1);
    ms.getRange(lr + 1, 1, newRows.length, 14).setValues(newRows);
  }

  SpreadsheetApp.getUi().alert(name + ': ' + raw.length + ' pulled, ' + newRows.length + ' added');
}

// ===================== SYNC ALL =====================

function syncAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('Run Setup Sheet first.'); return; }

  var names = Object.keys(PLATFORMS);
  var totalP = 0, totalA = 0;

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var p = PLATFORMS[name];
    var existing = getExisting(ms);

    try {
      var raw = fastRead(p.id, p.tab, p.cols);
      totalP += raw.length;

      var seen = {};
      var newRows = [];

      for (var r = 0; r < raw.length; r++) {
        var key = String(raw[r][0]).trim();
        if (!key) {
          var oid = String(raw[r][7]).trim();
          if (!oid) continue;
          key = String(raw[r][5]).trim() + String(raw[r][6]).trim() + oid + String(raw[r][10]).trim();
          raw[r][0] = key;
        }
        if (!String(raw[r][1]).trim()) {
          var tid = String(raw[r][8]).trim();
          if (tid.endsWith('.0')) tid = tid.slice(0, -2);
          raw[r][1] = String(raw[r][5]).trim() + String(raw[r][6]).trim() + tid;
        }
        if (existing[key] || seen[key]) continue;
        seen[key] = true;
        newRows.push(raw[r]);
      }

      if (newRows.length > 0) {
        var lr = Math.max(ms.getLastRow(), 1);
        ms.getRange(lr + 1, 1, newRows.length, 14).setValues(newRows);
        totalA += newRows.length;
      }

      Logger.log(name + ': ' + raw.length + ' pulled, ' + newRows.length + ' added');
      SpreadsheetApp.flush();
    } catch(e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert('All done!\nPulled: ' + totalP + '\nAdded: ' + totalA + '\nSkipped: ' + (totalP - totalA));
}

// ===================== GET EXISTING UNIQUES =====================

function getExisting(ms) {
  var map = {};
  var lr = ms.getLastRow();
  if (lr < 2) return map;
  var vals = ms.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0]).trim();
    if (k) map[k] = true;
  }
  return map;
}

// ===================== REMOVE DUPLICATES =====================

function removeDuplicates() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  var lr = ms.getLastRow();
  if (lr <= 1) return;
  var lc = ms.getLastColumn();
  var data = ms.getRange(2, 1, lr - 1, lc).getValues();
  var seen = {}, keep = [], dupes = 0;
  for (var i = 0; i < data.length; i++) {
    var k = String(data[i][0]).trim();
    if (!k) continue;
    if (seen[k]) { dupes++; continue; }
    seen[k] = true;
    keep.push(data[i]);
  }
  if (dupes > 0) {
    ms.getRange(2, 1, lr - 1, lc).clearContent();
    if (keep.length > 0) ms.getRange(2, 1, keep.length, lc).setValues(keep);
    SpreadsheetApp.getUi().alert('Removed ' + dupes + ' duplicates.');
  } else {
    SpreadsheetApp.getUi().alert('No duplicates found.');
  }
}

// ===================== LOAD SKU MAP (key=AA, pull AC,AD,AG,AJ,AL,AM,AN) =====================

function loadSkuMap() {
  // Read columns AA, AC, AD, AG, AJ, AL, AM, AN from Master SKU FILE
  // AA=key, AC→O, AD→P, AG→Q, AJ→R, AL→S, AM→T, AN→U
  var ranges = [
    "'" + SKU_TAB_NAME + "'!AA:AA",  // key
    "'" + SKU_TAB_NAME + "'!AC:AC",  // → O (15)
    "'" + SKU_TAB_NAME + "'!AD:AD",  // → P (16)
    "'" + SKU_TAB_NAME + "'!AG:AG",  // → Q (17)
    "'" + SKU_TAB_NAME + "'!AJ:AJ",  // → R (18)
    "'" + SKU_TAB_NAME + "'!AL:AL",  // → S (19)
    "'" + SKU_TAB_NAME + "'!AM:AM",  // → T (20) ← NEW
    "'" + SKU_TAB_NAME + "'!AN:AN"   // → U (21)
  ];

  var res = Sheets.Spreadsheets.Values.batchGet(SKU_SHEET_ID, {ranges: ranges});
  var vr = res.valueRanges;

  var map = {};
  var keyVals = vr[0].values || [];

  for (var i = 1; i < keyVals.length; i++) { // skip header row
    var skuKey = String(keyVals[i][0] || '').trim();
    if (!skuKey) continue;

    var row = [];
    for (var c = 1; c < vr.length; c++) {
      var vals = vr[c].values || [];
      row.push(i < vals.length && vals[i].length > 0 ? vals[i][0] : '');
    }
    map[skuKey] = row; // [AC, AD, AG, AJ, AL, AM, AN] — 7 values
  }

  Logger.log('Loaded ' + Object.keys(map).length + ' SKUs from Master SKU FILE');
  return map;
}

// ===================== FILL SKU DATA (columns O-U, 7 cols) =====================

function fillSkuData() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to fill.'); return; }

  // Load SKU lookup map
  var skuMap = loadSkuMap();

  // Read column K (SKU ID) - column 11
  var skuIds = ms.getRange(2, 11, lr - 1, 1).getValues();

  // Build output for columns O-U (columns 15-21, 7 cols)
  var output = [];
  var found = 0, notFound = 0;

  for (var i = 0; i < skuIds.length; i++) {
    var skuId = String(skuIds[i][0]).trim();
    var match = skuMap[skuId];

    if (match) {
      output.push(match);
      found++;
    } else {
      output.push(['', '', '', '', '', '', '']);
      notFound++;
    }
  }

  // Write columns O-U (columns 15-21) starting from row 2
  ms.getRange(2, 15, output.length, 7).setValues(output);

  SpreadsheetApp.getUi().alert('SKU data filled!\nMatched: ' + found + '\nNot found: ' + notFound);
}

// ===================== FILL PRICES (column V = col 22) =====================
// Key = MASTER CSV col O matched against PRICES col A
// Price column chosen by platform in col F:
//   Firstcry→K, Flipkart→U, Meesho→AA, Amazon/AmazonFBA→AF, Myntra→AK

function loadPriceMap() {
  // Read key (A) + all 5 price columns at once
  var ranges = [
    "'" + PRICES_TAB_NAME + "'!A:A",   // key
    "'" + PRICES_TAB_NAME + "'!K:K",   // Firstcry price
    "'" + PRICES_TAB_NAME + "'!U:U",   // Flipkart price
    "'" + PRICES_TAB_NAME + "'!AA:AA", // Meesho price
    "'" + PRICES_TAB_NAME + "'!AF:AF", // Amazon price
    "'" + PRICES_TAB_NAME + "'!AK:AK"  // Myntra price
  ];

  var res = Sheets.Spreadsheets.Values.batchGet(PRICES_SHEET_ID, {ranges: ranges});
  var vr = res.valueRanges;
  var keyVals = vr[0].values || [];

  // map[key] = { Firstcry: val, Flipkart: val, Meesho: val, Amazon: val, Myntra: val }
  var map = {};
  for (var i = 1; i < keyVals.length; i++) {
    var k = String(keyVals[i][0] || '').trim();
    if (!k) continue;
    map[k] = {
      'Firstcry':  (i < (vr[1].values||[]).length && vr[1].values[i].length > 0) ? vr[1].values[i][0] : '',
      'Flipkart':  (i < (vr[2].values||[]).length && vr[2].values[i].length > 0) ? vr[2].values[i][0] : '',
      'Meesho':    (i < (vr[3].values||[]).length && vr[3].values[i].length > 0) ? vr[3].values[i][0] : '',
      'Amazon':    (i < (vr[4].values||[]).length && vr[4].values[i].length > 0) ? vr[4].values[i][0] : '',
      'Myntra':    (i < (vr[5].values||[]).length && vr[5].values[i].length > 0) ? vr[5].values[i][0] : ''
    };
  }
  Logger.log('Loaded ' + Object.keys(map).length + ' price entries');
  return map;
}

function fillPrices() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to fill.'); return; }

  var priceMap = loadPriceMap();

  var platforms = ms.getRange(2, 6, lr - 1, 1).getValues();   // col F
  var lookupKeys = ms.getRange(2, 15, lr - 1, 1).getValues(); // col O
  var colM = ms.getRange(2, 13, lr - 1, 1).getValues();       // col M (Quantity)

  var output = [];
  var found = 0, notFound = 0;

  for (var i = 0; i < lookupKeys.length; i++) {
    var key = String(lookupKeys[i][0]).trim();
    var plat = String(platforms[i][0]).trim();
    var qty = parseFloat(colM[i][0]) || 0;

    var normPlat = '';
    var pu = plat.toUpperCase();
    if (pu === 'FIRSTCRY') normPlat = 'Firstcry';
    else if (pu === 'FLIPKART') normPlat = 'Flipkart';
    else if (pu === 'MEESHO') normPlat = 'Meesho';
    else if (pu === 'AMAZON' || pu === 'AMAZON FBA') normPlat = 'Amazon';
    else if (pu === 'MYNTRA') normPlat = 'Myntra';

    var entry = priceMap[key];
    if (entry && normPlat && entry[normPlat] !== '') {
      var price = parseFloat(entry[normPlat]) || 0;
      output.push([Math.round(price * qty * 100) / 100]);
      found++;
    } else {
      output.push(['']);
      notFound++;
    }
  }

  // Write to column V (col 22) starting from row 2
  ms.getRange(2, 22, output.length, 1).setValues(output);

  SpreadsheetApp.getUi().alert('Prices×Qty filled (col V)!\nMatched: ' + found + '\nNot found: ' + notFound);
}

// ===================== FILL TAX % (column X = col 24) =====================
// Key = MASTER CSV col O matched against PRICES col A
// Tax % column chosen by platform in col F:
//   Firstcry→I, Flipkart→S, Meesho→Y, Amazon/AmazonFBA→AD, Myntra→AI

function loadTaxMap() {
  var ranges = [
    "'" + PRICES_TAB_NAME + "'!A:A",   // key
    "'" + PRICES_TAB_NAME + "'!I:I",   // Firstcry tax %
    "'" + PRICES_TAB_NAME + "'!S:S",   // Flipkart tax %
    "'" + PRICES_TAB_NAME + "'!Y:Y",   // Meesho tax %
    "'" + PRICES_TAB_NAME + "'!AD:AD", // Amazon tax %
    "'" + PRICES_TAB_NAME + "'!AI:AI"  // Myntra tax %
  ];

  var res = Sheets.Spreadsheets.Values.batchGet(PRICES_SHEET_ID, {ranges: ranges});
  var vr = res.valueRanges;
  var keyVals = vr[0].values || [];

  var map = {};
  for (var i = 1; i < keyVals.length; i++) {
    var k = String(keyVals[i][0] || '').trim();
    if (!k) continue;
    map[k] = {
      'Firstcry':  (i < (vr[1].values||[]).length && vr[1].values[i].length > 0) ? vr[1].values[i][0] : '',
      'Flipkart':  (i < (vr[2].values||[]).length && vr[2].values[i].length > 0) ? vr[2].values[i][0] : '',
      'Meesho':    (i < (vr[3].values||[]).length && vr[3].values[i].length > 0) ? vr[3].values[i][0] : '',
      'Amazon':    (i < (vr[4].values||[]).length && vr[4].values[i].length > 0) ? vr[4].values[i][0] : '',
      'Myntra':    (i < (vr[5].values||[]).length && vr[5].values[i].length > 0) ? vr[5].values[i][0] : ''
    };
  }
  Logger.log('Loaded ' + Object.keys(map).length + ' tax % entries');
  return map;
}

function fillTaxPct() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to fill.'); return; }

  var taxMap = loadTaxMap();

  var platforms = ms.getRange(2, 6, lr - 1, 1).getValues();    // col F
  var lookupKeys = ms.getRange(2, 15, lr - 1, 1).getValues();  // col O

  var output = [];
  var found = 0, notFound = 0;

  for (var i = 0; i < lookupKeys.length; i++) {
    var key = String(lookupKeys[i][0]).trim();
    var plat = String(platforms[i][0]).trim();

    var normPlat = '';
    var pu = plat.toUpperCase();
    if (pu === 'FIRSTCRY') normPlat = 'Firstcry';
    else if (pu === 'FLIPKART') normPlat = 'Flipkart';
    else if (pu === 'MEESHO') normPlat = 'Meesho';
    else if (pu === 'AMAZON' || pu === 'AMAZON FBA') normPlat = 'Amazon';
    else if (pu === 'MYNTRA') normPlat = 'Myntra';

    var entry = taxMap[key];
    if (entry && normPlat && entry[normPlat] !== '') {
      output.push([entry[normPlat]]);
      found++;
    } else {
      output.push(['']);
      notFound++;
    }
  }

  // Write to column X (col 24) starting from row 2
  ms.getRange(2, 24, output.length, 1).setValues(output);

  SpreadsheetApp.getUi().alert('Tax % filled (col X)!\nMatched: ' + found + '\nNot found: ' + notFound);
}

// ===================== FILL COL W & AC (platform source sheets) =====================
// Col W: Meesho→N, Flipkart→X, Firstcry→Y, Amazon/FBA→S, Myntra→K
// Col AC: Meesho only→O
// Key: col A in both sheets, EXCEPT Myntra where result key = F+G+K (built internally)

function loadColVMap(src) {
  // Read key (A) + value column, and optionally AB column
  var ranges = ["'" + src.tab + "'!A:A", "'" + src.tab + "'!" + src.valCol + ":" + src.valCol];
  if (src.abCol) {
    ranges.push("'" + src.tab + "'!" + src.abCol + ":" + src.abCol);
  }

  var res = Sheets.Spreadsheets.Values.batchGet(src.id, {ranges: ranges});
  var vr = res.valueRanges;
  var keyVals = vr[0].values || [];

  var map = {};
  for (var i = 1; i < keyVals.length; i++) {
    var k = String(keyVals[i][0] || '').trim();
    if (!k) continue;
    var val = (i < (vr[1].values||[]).length && vr[1].values[i].length > 0) ? vr[1].values[i][0] : '';
    var abVal = '';
    if (src.abCol && vr[2]) {
      abVal = (i < (vr[2].values||[]).length && vr[2].values[i].length > 0) ? vr[2].values[i][0] : '';
    }
    map[k] = { v: val, ab: abVal };
  }
  return map;
}

function fillColV() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to fill.'); return; }

  // Read cols A(1), F(6), G(7), K(11), M(13) from result sheet
  var colA = ms.getRange(2, 1, lr - 1, 1).getValues();   // UNIQUE key
  var colF = ms.getRange(2, 6, lr - 1, 1).getValues();   // Platform
  var colG = ms.getRange(2, 7, lr - 1, 1).getValues();   // Company
  var colK = ms.getRange(2, 11, lr - 1, 1).getValues();  // SKU ID
  var colM = ms.getRange(2, 13, lr - 1, 1).getValues();  // Quantity (for Myntra ×M)

  // Load maps for each platform (only load once per unique source)
  var loaded = {};
  var maps = {};
  var srcKeys = Object.keys(COL_V_SOURCES);
  for (var s = 0; s < srcKeys.length; s++) {
    var platName = srcKeys[s];
    var src = COL_V_SOURCES[platName];
    var cacheKey = src.id + '|' + src.tab;
    if (!loaded[cacheKey]) {
      loaded[cacheKey] = loadColVMap(src);
      Logger.log('Loaded ' + Object.keys(loaded[cacheKey]).length + ' rows from ' + src.tab);
    }
    maps[platName] = loaded[cacheKey];
  }

  var outputW = [];
  var outputAC = [];
  var foundW = 0, foundAC = 0;

  for (var i = 0; i < colA.length; i++) {
    var plat = String(colF[i][0]).trim().toUpperCase();
    var platMap = maps[plat];

    if (!platMap) {
      outputW.push(['']);
      outputAC.push(['']);
      continue;
    }

    // Determine lookup key
    var lookupKey;
    if (plat === 'MYNTRA') {
      lookupKey = String(colF[i][0]).trim() + String(colG[i][0]).trim() + String(colK[i][0]).trim();
    } else {
      lookupKey = String(colA[i][0]).trim();
    }

    var entry = platMap[lookupKey];

    if (entry && entry.v !== '') {
      // Myntra & Meesho: multiply by col M (Quantity)
      if (plat === 'MYNTRA' || plat === 'MEESHO') {
        var qty = parseFloat(colM[i][0]) || 0;
        var val = parseFloat(entry.v) || 0;
        outputW.push([Math.round(val * qty * 100) / 100]);
      } else {
        outputW.push([entry.v]);
      }
      foundW++;
    } else {
      outputW.push(['']);
    }

    // Col AC — only Meesho has abCol
    if (plat === 'MEESHO' && entry && entry.ab !== '') {
      var abVal = parseFloat(entry.ab) || 0;
      var abQty = parseFloat(colM[i][0]) || 0;
      outputAC.push([Math.round(abVal * abQty * 100) / 100]);
      foundAC++;
    } else {
      outputAC.push(['']);
    }
  }

  // Write col W (col 23) and col AC (col 29)
  ms.getRange(2, 23, outputW.length, 1).setValues(outputW);
  ms.getRange(2, 29, outputAC.length, 1).setValues(outputAC);

  SpreadsheetApp.getUi().alert('Col W filled: ' + foundW + ' matched\nCol AC (Meesho): ' + foundAC + ' matched');
}

// ===================== CALC TAX VALUES (columns Y & Z) =====================
// Z = Taxable Value = W / (100 + X) * 100
// Y = Tax Value = W - Z
// W = col 23 (invoice value incl. tax), X = col 24 (tax %)

function calcTaxValues() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to calc.'); return; }

  var colW = ms.getRange(2, 23, lr - 1, 1).getValues();  // invoice value
  var colX = ms.getRange(2, 24, lr - 1, 1).getValues();  // tax %

  var outputZ = [];  // taxable value
  var outputY = [];  // tax value
  var calc = 0;

  for (var i = 0; i < colW.length; i++) {
    var w = parseFloat(colW[i][0]);
    var x = parseFloat(colX[i][0]);

    if (!isNaN(w) && !isNaN(x) && (100 + x) !== 0) {
      var taxable = w / (100 + x) * 100;
      var taxVal = w - taxable;
      outputY.push([Math.round(taxVal * 100) / 100]);
      outputZ.push([Math.round(taxable * 100) / 100]);
      calc++;
    } else {
      outputY.push(['']);
      outputZ.push(['']);
    }
  }

  // Y = col 25 (Tax Value), Z = col 26 (Taxable Value)
  ms.getRange(2, 25, outputY.length, 1).setValues(outputY);
  ms.getRange(2, 26, outputZ.length, 1).setValues(outputZ);

  SpreadsheetApp.getUi().alert('Tax calc done!\nY (Tax Value) & Z (Taxable): ' + calc + ' rows calculated');
}

// ===================== RUN ALL FILLS =====================
// Runs: SKU(O-U) → Prices(V) → ColW&AC → Tax%(X) → TaxCalc(Y,Z) → ColAA → AB in one go

function runAllFills() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName('MASTER CSV');
  if (!ms) { ui.alert('ERROR: No MASTER CSV tab found.'); return; }
  var lr = ms.getLastRow();
  if (lr < 2) { ui.alert('ERROR: No data in MASTER CSV. Run Sync All first.'); return; }

  var rows = lr - 1;
  var log = [];
  log.push('Rows found: ' + rows);
  var skuF = 0, priceF = 0, vF = 0, abF = 0, taxF = 0, calcN = 0;

  // ---- STEP 1: Fill SKU Data (O-U, 7 cols) ----
  try {
    var skuMap = loadSkuMap();
    log.push('SKU map loaded: ' + Object.keys(skuMap).length + ' entries');
    var skuIds = ms.getRange(2, 11, rows, 1).getValues();
    var skuOut = [];
    for (var i = 0; i < skuIds.length; i++) {
      var match = skuMap[String(skuIds[i][0]).trim()];
      skuOut.push(match || ['', '', '', '', '', '', '']);
      if (match) skuF++;
    }
    ms.getRange(2, 15, skuOut.length, 7).setValues(skuOut);
    SpreadsheetApp.flush();
    log.push('Step 1 SKU: ' + skuF + ' matched');
  } catch(e) {
    log.push('Step 1 SKU ERROR: ' + e.message);
  }

  // ---- STEP 2: Fill Prices × Qty (V = col 22) ----
  try {
    var priceMap = loadPriceMap();
    log.push('Price map loaded: ' + Object.keys(priceMap).length + ' entries');
    var plats = ms.getRange(2, 6, rows, 1).getValues();
    var keys = ms.getRange(2, 15, rows, 1).getValues();
    var qtyM = ms.getRange(2, 13, rows, 1).getValues();
    for (var d = 0; d < Math.min(3, keys.length); d++) {
      log.push('  Row ' + (d+2) + ': plat=' + plats[d][0] + ', key=' + keys[d][0]);
    }
    var priceOut = [];
    for (var i = 0; i < keys.length; i++) {
      var pu2 = String(plats[i][0]).trim().toUpperCase();
      var np = normalizePlat(String(plats[i][0]).trim());
      var e = priceMap[String(keys[i][0]).trim()];
      if (e && np && e[np] !== '') {
        var price = parseFloat(e[np]) || 0;
        if (pu2 === 'MYNTRA' || pu2 === 'MEESHO') {
          var qty = parseFloat(qtyM[i][0]) || 0;
          priceOut.push([Math.round(price * qty * 100) / 100]);
        } else {
          priceOut.push([Math.round(price * 100) / 100]);
        }
        priceF++;
      }
      else { priceOut.push(['']); }
    }
    ms.getRange(2, 22, priceOut.length, 1).setValues(priceOut);
    SpreadsheetApp.flush();
    log.push('Step 2 Prices×Qty: ' + priceF + ' matched');
  } catch(e) {
    log.push('Step 2 Prices ERROR: ' + e.message);
  }

  // ---- STEP 3: Fill Col W (col 23) & AC (col 29) — Myntra W ×M ----
  try {
    var colA = ms.getRange(2, 1, rows, 1).getValues();
    var colF = ms.getRange(2, 6, rows, 1).getValues();
    var colG = ms.getRange(2, 7, rows, 1).getValues();
    var colK = ms.getRange(2, 11, rows, 1).getValues();
    var colM3 = ms.getRange(2, 13, rows, 1).getValues();  // Quantity (for Myntra ×M)

    var loaded = {}, maps = {};
    var srcKeys = Object.keys(COL_V_SOURCES);
    for (var s = 0; s < srcKeys.length; s++) {
      var src = COL_V_SOURCES[srcKeys[s]];
      var ck = src.id + '|' + src.tab;
      if (!loaded[ck]) {
        loaded[ck] = loadColVMap(src);
        log.push('ColW map ' + src.tab + ': ' + Object.keys(loaded[ck]).length + ' entries');
      }
      maps[srcKeys[s]] = loaded[ck];
    }

    var outW = [], outAC = [];
    for (var i = 0; i < colA.length; i++) {
      var pu = String(colF[i][0]).trim().toUpperCase();
      var pm = maps[pu];
      if (!pm) { outW.push(['']); outAC.push(['']); continue; }

      var lk = (pu === 'MYNTRA')
        ? String(colF[i][0]).trim() + String(colG[i][0]).trim() + String(colK[i][0]).trim()
        : String(colA[i][0]).trim();

      var en = pm[lk];
      if (en && en.v !== '') {
        if (pu === 'MYNTRA' || pu === 'MEESHO') {
          var qty = parseFloat(colM3[i][0]) || 0;
          var val = parseFloat(en.v) || 0;
          outW.push([Math.round(val * qty * 100) / 100]);
        } else {
          outW.push([en.v]);
        }
        vF++;
      } else { outW.push(['']); }
      if (pu === 'MEESHO' && en && en.ab !== '') {
        var abVal = parseFloat(en.ab) || 0;
        var abQty = parseFloat(colM3[i][0]) || 0;
        outAC.push([Math.round(abVal * abQty * 100) / 100]); abF++;
      } else { outAC.push(['']); }
    }
    ms.getRange(2, 23, outW.length, 1).setValues(outW);
    ms.getRange(2, 29, outAC.length, 1).setValues(outAC);
    SpreadsheetApp.flush();
    log.push('Step 3 ColW: ' + vF + ', AC: ' + abF);
  } catch(e) {
    log.push('Step 3 ColW ERROR: ' + e.message);
  }

  // ---- STEP 4: Fill Tax % (X = col 24) ----
  try {
    var taxMap = loadTaxMap();
    log.push('Tax map loaded: ' + Object.keys(taxMap).length + ' entries');
    var plats4 = ms.getRange(2, 6, rows, 1).getValues();
    var keys4 = ms.getRange(2, 15, rows, 1).getValues();
    var taxOut = [];
    for (var i = 0; i < keys4.length; i++) {
      var np = normalizePlat(String(plats4[i][0]).trim());
      var e = taxMap[String(keys4[i][0]).trim()];
      if (e && np && e[np] !== '') { taxOut.push([e[np]]); taxF++; }
      else { taxOut.push(['']); }
    }
    ms.getRange(2, 24, taxOut.length, 1).setValues(taxOut);
    SpreadsheetApp.flush();
    log.push('Step 4 Tax%: ' + taxF + ' matched');
  } catch(e) {
    log.push('Step 4 Tax% ERROR: ' + e.message);
  }

  // ---- STEP 5: Calc Tax Values (Y = Tax Value col 25, Z = Taxable Value col 26) ----
  try {
    var colWv = ms.getRange(2, 23, rows, 1).getValues();  // invoice (W)
    var colXv = ms.getRange(2, 24, rows, 1).getValues();  // tax % (X)
    var yOut = [], zOut = [];
    for (var i = 0; i < colWv.length; i++) {
      var w = parseFloat(colWv[i][0]);
      var x = parseFloat(colXv[i][0]);
      if (!isNaN(w) && !isNaN(x) && (100 + x) !== 0) {
        var taxable = w / (100 + x) * 100;
        yOut.push([Math.round((w - taxable) * 100) / 100]);
        zOut.push([Math.round(taxable * 100) / 100]);
        calcN++;
      } else { yOut.push(['']); zOut.push(['']); }
    }
    ms.getRange(2, 25, yOut.length, 1).setValues(yOut);  // Y = col 25
    ms.getRange(2, 26, zOut.length, 1).setValues(zOut);   // Z = col 26
    SpreadsheetApp.flush();
    log.push('TaxCalc Y,Z: ' + calcN + ' done');
  } catch(e) {
    log.push('TaxCalc ERROR: ' + e.message);
  }

  // ---- STEP 6: Col AA (PRICES col F × Qty, col 27) — key = col O ----
  var aaF = 0;
  try {
    var colFMap = loadPricesColFMap();
    log.push('ColF map loaded: ' + Object.keys(colFMap).length + ' entries');
    var colOz = ms.getRange(2, 15, rows, 1).getValues();
    var colMz = ms.getRange(2, 13, rows, 1).getValues();
    var aaOut = [];
    for (var i = 0; i < colOz.length; i++) {
      var key = String(colOz[i][0]).trim();
      var qty = parseFloat(colMz[i][0]) || 0;
      var val = colFMap[key];
      if (val !== undefined && val !== '') {
        var num = parseFloat(val) || 0;
        aaOut.push([Math.round(num * qty * 100) / 100]);
        aaF++;
      } else { aaOut.push(['']); }
    }
    ms.getRange(2, 27, aaOut.length, 1).setValues(aaOut);  // AA = col 27
    SpreadsheetApp.flush();
    log.push('Step 6 ColAA: ' + aaF + ' matched');
  } catch(e) {
    log.push('Step 6 ColAA ERROR: ' + e.message);
  }

  // ---- STEP 7: AB = Z - AA (col 28) ----
  var abCalc = 0;
  try {
    var colZab = ms.getRange(2, 26, rows, 1).getValues();   // Z (Taxable Value)
    var colAAab = ms.getRange(2, 27, rows, 1).getValues();  // AA
    var abOut = [];
    for (var i = 0; i < colZab.length; i++) {
      var zVal = parseFloat(colZab[i][0]);
      var aaVal = parseFloat(colAAab[i][0]);
      if (!isNaN(zVal) && !isNaN(aaVal)) {
        abOut.push([Math.round((zVal - aaVal) * 100) / 100]);
        abCalc++;
      } else { abOut.push(['']); }
    }
    ms.getRange(2, 28, abOut.length, 1).setValues(abOut);  // AB = col 28
    SpreadsheetApp.flush();
    log.push('Step 7 AB: ' + abCalc + ' done');
  } catch(e) {
    log.push('Step 7 AB ERROR: ' + e.message);
  }

  // Show full debug log
  ui.alert('RUN ALL FILLS LOG:\n' + log.join('\n'));
}

// ===================== FILL COL AA (VLOOKUP PRICES col F × col M) =====================
// Simple VLOOKUP: key = result col O matched against PRICES col A, pull col F
// Then multiply by col M (Quantity)

function loadPricesColFMap() {
  var ranges = [
    "'" + PRICES_TAB_NAME + "'!A:A",  // key
    "'" + PRICES_TAB_NAME + "'!F:F"   // value
  ];
  var res = Sheets.Spreadsheets.Values.batchGet(PRICES_SHEET_ID, {ranges: ranges});
  var vr = res.valueRanges;
  var keyVals = vr[0].values || [];

  var map = {};
  for (var i = 1; i < keyVals.length; i++) {
    var k = String(keyVals[i][0] || '').trim();
    if (!k) continue;
    var val = (i < (vr[1].values||[]).length && vr[1].values[i].length > 0) ? vr[1].values[i][0] : '';
    map[k] = val;
  }
  Logger.log('Loaded ' + Object.keys(map).length + ' PRICES col F entries');
  return map;
}

function fillColZ() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to fill.'); return; }

  var colFMap = loadPricesColFMap();

  var colO = ms.getRange(2, 15, lr - 1, 1).getValues();  // key = col O
  var colM = ms.getRange(2, 13, lr - 1, 1).getValues();   // Quantity

  var output = [];
  var found = 0;

  for (var i = 0; i < colO.length; i++) {
    var key = String(colO[i][0]).trim();
    var qty = parseFloat(colM[i][0]) || 0;
    var val = colFMap[key];

    if (val !== undefined && val !== '') {
      var num = parseFloat(val) || 0;
      output.push([Math.round(num * qty * 100) / 100]);
      found++;
    } else {
      output.push(['']);
    }
  }

  // Write to col AA (col 27)
  ms.getRange(2, 27, output.length, 1).setValues(output);

  SpreadsheetApp.getUi().alert('Col AA filled (PRICES F × Qty)!\nMatched: ' + found);
}

// ===================== CALC COL AB (Z - AA) =====================
// AB = col 28 = Z (Taxable Value, col 26) - AA (col 27)

function calcColAA() {
  var ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER CSV');
  if (!ms) { SpreadsheetApp.getUi().alert('No MASTER CSV tab found.'); return; }

  var lr = ms.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert('No data to calc.'); return; }

  var colZ = ms.getRange(2, 26, lr - 1, 1).getValues();   // Taxable Value
  var colAA = ms.getRange(2, 27, lr - 1, 1).getValues();  // Col AA

  var output = [];
  var calc = 0;

  for (var i = 0; i < colZ.length; i++) {
    var z = parseFloat(colZ[i][0]);
    var aa = parseFloat(colAA[i][0]);

    if (!isNaN(z) && !isNaN(aa)) {
      output.push([Math.round((z - aa) * 100) / 100]);
      calc++;
    } else {
      output.push(['']);
    }
  }

  // AB = col 28
  ms.getRange(2, 28, output.length, 1).setValues(output);

  SpreadsheetApp.getUi().alert('Col AB (Z-AA) done: ' + calc + ' rows calculated');
}

// ===================== HELPER: Normalize Platform =====================

function normalizePlat(plat) {
  var pu = plat.toUpperCase();
  if (pu === 'FIRSTCRY') return 'Firstcry';
  if (pu === 'FLIPKART') return 'Flipkart';
  if (pu === 'MEESHO') return 'Meesho';
  if (pu === 'AMAZON' || pu === 'AMAZON FBA') return 'Amazon';
  if (pu === 'MYNTRA') return 'Myntra';
  return '';
}

// ===================== RUN EVERYTHING =====================
// Sync All Platforms + All Fills + Tax Calc in one go

function runEverything() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName('MASTER CSV');
  if (!ms) {
    ms = ss.insertSheet('MASTER CSV');
  }

  var log = [];
  var totalP = 0, totalA = 0;

  // ===== PHASE 1: SYNC ALL PLATFORMS =====
  log.push('--- SYNC ---');
  var names = Object.keys(PLATFORMS);
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var p = PLATFORMS[name];
    var existing = getExisting(ms);
    try {
      var raw = fastRead(p.id, p.tab, p.cols);
      totalP += raw.length;
      var seen = {};
      var newRows = [];
      for (var r = 0; r < raw.length; r++) {
        var key = String(raw[r][0]).trim();
        if (!key) {
          var oid = String(raw[r][7]).trim();
          if (!oid) continue;
          key = String(raw[r][5]).trim() + String(raw[r][6]).trim() + oid + String(raw[r][10]).trim();
          raw[r][0] = key;
        }
        if (!String(raw[r][1]).trim()) {
          var tid = String(raw[r][8]).trim();
          if (tid.endsWith('.0')) tid = tid.slice(0, -2);
          raw[r][1] = String(raw[r][5]).trim() + String(raw[r][6]).trim() + tid;
        }
        if (existing[key] || seen[key]) continue;
        seen[key] = true;
        newRows.push(raw[r]);
      }
      if (newRows.length > 0) {
        var lr = Math.max(ms.getLastRow(), 1);
        ms.getRange(lr + 1, 1, newRows.length, 14).setValues(newRows);
        totalA += newRows.length;
      }
      log.push(name + ': +' + newRows.length);
      SpreadsheetApp.flush();
    } catch(e) {
      log.push(name + ' ERR: ' + e.message);
    }
  }
  log.push('Sync total: ' + totalP + ' pulled, ' + totalA + ' added');

  // Re-check row count after sync
  var lr = ms.getLastRow();
  if (lr < 2) {
    ui.alert('Sync done but no data found.\n' + log.join('\n'));
    return;
  }
  var rows = lr - 1;

  // ===== PHASE 2: ALL FILLS =====
  log.push('--- FILLS ---');
  var skuF = 0, priceF = 0, vF = 0, abF = 0, taxF = 0, calcN = 0;

  // Step 1: SKU (O-U, 7 cols)
  try {
    var skuMap = loadSkuMap();
    log.push('SKU map size: ' + Object.keys(skuMap).length);
    var skuIds = ms.getRange(2, 11, rows, 1).getValues();
    // Debug: show first 3 SKU keys from result sheet and first 3 from map
    var mapKeys = Object.keys(skuMap);
    for (var d = 0; d < Math.min(3, skuIds.length); d++) {
      log.push('  ResultK[' + (d+2) + ']: "' + String(skuIds[d][0]).trim() + '"');
    }
    for (var d = 0; d < Math.min(3, mapKeys.length); d++) {
      log.push('  SKUmap key[' + d + ']: "' + mapKeys[d] + '"');
    }
    var skuOut = [];
    for (var i = 0; i < skuIds.length; i++) {
      var match = skuMap[String(skuIds[i][0]).trim()];
      skuOut.push(match || ['', '', '', '', '', '', '']);
      if (match) skuF++;
    }
    ms.getRange(2, 15, skuOut.length, 7).setValues(skuOut);
    SpreadsheetApp.flush();
    log.push('SKU: ' + skuF + '/' + rows + ' matched');
  } catch(e) { log.push('SKU ERR: ' + e.message); }

  // Step 2: Prices × Qty (V = col 22)
  try {
    var priceMap = loadPriceMap();
    log.push('Price map size: ' + Object.keys(priceMap).length);
    var plats = ms.getRange(2, 6, rows, 1).getValues();
    var keys = ms.getRange(2, 15, rows, 1).getValues();  // col O (filled by Step 1)
    var qtyM = ms.getRange(2, 13, rows, 1).getValues();
    // Debug: show first 3 col O keys and platforms
    for (var d = 0; d < Math.min(3, keys.length); d++) {
      log.push('  O[' + (d+2) + ']="' + String(keys[d][0]).trim() + '" plat=' + plats[d][0] + ' qty=' + qtyM[d][0]);
    }
    var priceOut = [];
    for (var i = 0; i < keys.length; i++) {
      var np = normalizePlat(String(plats[i][0]).trim());
      var e = priceMap[String(keys[i][0]).trim()];
      var qty = parseFloat(qtyM[i][0]) || 0;
      if (e && np && e[np] !== '') {
        var price = parseFloat(e[np]) || 0;
        priceOut.push([Math.round(price * qty * 100) / 100]);
        priceF++;
      } else { priceOut.push(['']); }
    }
    ms.getRange(2, 22, priceOut.length, 1).setValues(priceOut);
    SpreadsheetApp.flush();
    log.push('Prices: ' + priceF + ' matched');
  } catch(e) { log.push('Prices ERR: ' + e.message); }

  // Step 3: Col W (col 23) & AC (col 29) — Myntra W ×M
  try {
    var colA = ms.getRange(2, 1, rows, 1).getValues();
    var colF = ms.getRange(2, 6, rows, 1).getValues();
    var colG = ms.getRange(2, 7, rows, 1).getValues();
    var colK = ms.getRange(2, 11, rows, 1).getValues();
    var colM3 = ms.getRange(2, 13, rows, 1).getValues();  // Quantity (for Myntra ×M)
    var loaded = {}, maps = {};
    var srcKeys = Object.keys(COL_V_SOURCES);
    for (var s = 0; s < srcKeys.length; s++) {
      var src = COL_V_SOURCES[srcKeys[s]];
      var ck = src.id + '|' + src.tab;
      if (!loaded[ck]) loaded[ck] = loadColVMap(src);
      maps[srcKeys[s]] = loaded[ck];
    }
    var outW = [], outAC = [];
    for (var i = 0; i < colA.length; i++) {
      var pu = String(colF[i][0]).trim().toUpperCase();
      var pm = maps[pu];
      if (!pm) { outW.push(['']); outAC.push(['']); continue; }
      var lk = (pu === 'MYNTRA')
        ? String(colF[i][0]).trim() + String(colG[i][0]).trim() + String(colK[i][0]).trim()
        : String(colA[i][0]).trim();
      var en = pm[lk];
      if (en && en.v !== '') {
        if (pu === 'MYNTRA' || pu === 'MEESHO') {
          var qty = parseFloat(colM3[i][0]) || 0;
          var val = parseFloat(en.v) || 0;
          outW.push([Math.round(val * qty * 100) / 100]);
        } else {
          outW.push([en.v]);
        }
        vF++;
      } else { outW.push(['']); }
      if (pu === 'MEESHO' && en && en.ab !== '') {
        var abVal = parseFloat(en.ab) || 0;
        var abQty = parseFloat(colM3[i][0]) || 0;
        outAC.push([Math.round(abVal * abQty * 100) / 100]); abF++;
      } else { outAC.push(['']); }
    }
    ms.getRange(2, 23, outW.length, 1).setValues(outW);
    ms.getRange(2, 29, outAC.length, 1).setValues(outAC);
    SpreadsheetApp.flush();
    log.push('ColW: ' + vF + ', AC: ' + abF);
  } catch(e) { log.push('ColW ERR: ' + e.message); }

  // Step 4: Tax % (X = col 24)
  try {
    var taxMap = loadTaxMap();
    var plats4 = ms.getRange(2, 6, rows, 1).getValues();
    var keys4 = ms.getRange(2, 15, rows, 1).getValues();
    var taxOut = [];
    for (var i = 0; i < keys4.length; i++) {
      var np = normalizePlat(String(plats4[i][0]).trim());
      var e = taxMap[String(keys4[i][0]).trim()];
      if (e && np && e[np] !== '') { taxOut.push([e[np]]); taxF++; }
      else { taxOut.push(['']); }
    }
    ms.getRange(2, 24, taxOut.length, 1).setValues(taxOut);
    SpreadsheetApp.flush();
    log.push('Tax%: ' + taxF + ' matched');
  } catch(e) { log.push('Tax% ERR: ' + e.message); }

  // Step 5: Tax Calc (Y = Tax Value col 25, Z = Taxable Value col 26)
  try {
    var colWv = ms.getRange(2, 23, rows, 1).getValues();  // invoice (W)
    var colXv = ms.getRange(2, 24, rows, 1).getValues();  // tax % (X)
    var yOut = [], zOut = [];
    for (var i = 0; i < colWv.length; i++) {
      var w = parseFloat(colWv[i][0]);
      var x = parseFloat(colXv[i][0]);
      if (!isNaN(w) && !isNaN(x) && (100 + x) !== 0) {
        var taxable = w / (100 + x) * 100;
        yOut.push([Math.round((w - taxable) * 100) / 100]);
        zOut.push([Math.round(taxable * 100) / 100]);
        calcN++;
      } else { yOut.push(['']); zOut.push(['']); }
    }
    ms.getRange(2, 25, yOut.length, 1).setValues(yOut);  // Y = col 25
    ms.getRange(2, 26, zOut.length, 1).setValues(zOut);   // Z = col 26
    SpreadsheetApp.flush();
    log.push('TaxCalc Y,Z: ' + calcN + ' done');
  } catch(e) { log.push('TaxCalc ERR: ' + e.message); }

  // Step 6: Col AA (PRICES col F × Qty, col 27) — key = col O (15)
  var aaF2 = 0;
  try {
    var colFMap = loadPricesColFMap();
    var colOz = ms.getRange(2, 15, rows, 1).getValues();
    var colMz = ms.getRange(2, 13, rows, 1).getValues();
    var aaOut = [];
    for (var i = 0; i < colOz.length; i++) {
      var key = String(colOz[i][0]).trim();
      var qty = parseFloat(colMz[i][0]) || 0;
      var val = colFMap[key];
      if (val !== undefined && val !== '') {
        var num = parseFloat(val) || 0;
        aaOut.push([Math.round(num * qty * 100) / 100]);
        aaF2++;
      } else { aaOut.push(['']); }
    }
    ms.getRange(2, 27, aaOut.length, 1).setValues(aaOut);  // AA = col 27
    SpreadsheetApp.flush();
    log.push('ColAA: ' + aaF2 + ' matched');
  } catch(e) { log.push('ColAA ERR: ' + e.message); }

  // Step 7: AB = Z - AA (col 28)
  var abCalc = 0;
  try {
    var colZab = ms.getRange(2, 26, rows, 1).getValues();   // Z (Taxable Value)
    var colAAab = ms.getRange(2, 27, rows, 1).getValues();  // AA
    var abOut = [];
    for (var i = 0; i < colZab.length; i++) {
      var zv = parseFloat(colZab[i][0]);
      var aav = parseFloat(colAAab[i][0]);
      if (!isNaN(zv) && !isNaN(aav)) {
        abOut.push([Math.round((zv - aav) * 100) / 100]);
        abCalc++;
      } else { abOut.push(['']); }
    }
    ms.getRange(2, 28, abOut.length, 1).setValues(abOut);  // AB = col 28
    SpreadsheetApp.flush();
    log.push('AB(Z-AA): ' + abCalc + ' done');
  } catch(e) { log.push('AB ERR: ' + e.message); }

  // Remove duplicates
  try {
    var lrFinal = ms.getLastRow();
    if (lrFinal > 1) {
      var lc = ms.getLastColumn();
      var data = ms.getRange(2, 1, lrFinal - 1, lc).getValues();
      var seen = {}, keep = [], dupes = 0;
      for (var i = 0; i < data.length; i++) {
        var k = String(data[i][0]).trim();
        if (!k) continue;
        if (seen[k]) { dupes++; continue; }
        seen[k] = true;
        keep.push(data[i]);
      }
      if (dupes > 0) {
        ms.getRange(2, 1, lrFinal - 1, lc).clearContent();
        if (keep.length > 0) ms.getRange(2, 1, keep.length, lc).setValues(keep);
        log.push('Dupes removed: ' + dupes);
      }
    }
  } catch(e) { log.push('Dedup ERR: ' + e.message); }

  ui.alert('ALL DONE!\n' + log.join('\n'));
}

// ===================== SETUP =====================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName('MASTER CSV') || ss.insertSheet('MASTER CSV');
  SpreadsheetApp.getUi().alert('MASTER CSV tab ready! Row 1 = blank, data from row 2.');
}

// ===================== CLEAR & RUN FRESH =====================

function clearAndRun() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('This will DELETE all data in MASTER CSV and re-sync everything from scratch.\n\nContinue?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = ss.getSheetByName('MASTER CSV');
  if (ms) {
    var lr = ms.getLastRow();
    if (lr > 1) {
      ms.getRange(2, 1, lr - 1, ms.getMaxColumns()).clearContent();
      SpreadsheetApp.flush();
    }
  }
  // Now run everything fresh
  runEverything();
}

// ===================== AUTO-REFRESH (1-minute trigger) =====================

function setupAutoRefresh() {
  // Remove any existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create new 1-minute trigger
  ScriptApp.newTrigger('autoRefresh')
    .timeBased()
    .everyMinutes(1)
    .create();
  SpreadsheetApp.getUi().alert('Auto-refresh ON! Runs every 1 minute.\nUse "Stop Auto-Refresh" to disable.');
}

function stopAutoRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  SpreadsheetApp.getUi().alert('Auto-refresh OFF. Removed ' + removed + ' trigger(s).');
}

function autoRefresh() {
  // Use lock to prevent overlapping runs
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('autoRefresh skipped — previous run still going.');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName('MASTER CSV');
    if (!ms) {
      ms = ss.insertSheet('MASTER CSV');
    }

    var totalP = 0, totalA = 0;

    // ===== SYNC ALL PLATFORMS =====
    var names = Object.keys(PLATFORMS);
    for (var i = 0; i < names.length; i++) {
      var name = names[i];
      var p = PLATFORMS[name];
      var existing = getExisting(ms);
      try {
        var raw = fastRead(p.id, p.tab, p.cols);
        totalP += raw.length;
        var seen = {};
        var newRows = [];
        for (var r = 0; r < raw.length; r++) {
          var key = String(raw[r][0]).trim();
          if (!key) {
            var oid = String(raw[r][7]).trim();
            if (!oid) continue;
            key = String(raw[r][5]).trim() + String(raw[r][6]).trim() + oid + String(raw[r][10]).trim();
            raw[r][0] = key;
          }
          if (!String(raw[r][1]).trim()) {
            var tid = String(raw[r][8]).trim();
            if (tid.endsWith('.0')) tid = tid.slice(0, -2);
            raw[r][1] = String(raw[r][5]).trim() + String(raw[r][6]).trim() + tid;
          }
          if (existing[key] || seen[key]) continue;
          seen[key] = true;
          newRows.push(raw[r]);
        }
        if (newRows.length > 0) {
          var lr = Math.max(ms.getLastRow(), 1);
          ms.getRange(lr + 1, 1, newRows.length, 14).setValues(newRows);
          totalA += newRows.length;
        }
        SpreadsheetApp.flush();
      } catch(e) {
        Logger.log('autoRefresh SYNC ' + name + ' ERR: ' + e.message);
      }
    }

    // Re-check row count
    var lr = ms.getLastRow();
    if (lr < 2) { lock.releaseLock(); return; }
    var rows = lr - 1;

    // ===== ALL FILLS (same as runEverything Phase 2) =====

    // Step 1: SKU (O-U, 7 cols)
    try {
      var skuMap = loadSkuMap();
      var skuIds = ms.getRange(2, 11, rows, 1).getValues();
      var skuOut = [];
      for (var i = 0; i < skuIds.length; i++) {
        var match = skuMap[String(skuIds[i][0]).trim()];
        skuOut.push(match || ['', '', '', '', '', '', '']);
      }
      ms.getRange(2, 15, skuOut.length, 7).setValues(skuOut);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh SKU ERR: ' + e.message); }

    // Step 2: Prices × Qty (V = col 22)
    try {
      var priceMap = loadPriceMap();
      var plats = ms.getRange(2, 6, rows, 1).getValues();
      var keys = ms.getRange(2, 15, rows, 1).getValues();
      var qtyM = ms.getRange(2, 13, rows, 1).getValues();
      var priceOut = [];
      for (var i = 0; i < keys.length; i++) {
        var pu2 = String(plats[i][0]).trim().toUpperCase();
        var np = normalizePlat(String(plats[i][0]).trim());
        var e = priceMap[String(keys[i][0]).trim()];
        if (e && np && e[np] !== '') {
          var price = parseFloat(e[np]) || 0;
          if (pu2 === 'MYNTRA' || pu2 === 'MEESHO') {
            var qty = parseFloat(qtyM[i][0]) || 0;
            priceOut.push([Math.round(price * qty * 100) / 100]);
          } else {
            priceOut.push([Math.round(price * 100) / 100]);
          }
        } else { priceOut.push(['']); }
      }
      ms.getRange(2, 22, priceOut.length, 1).setValues(priceOut);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh Prices ERR: ' + e.message); }

    // Step 3: Col W (col 23) & AC (col 29) — Myntra ×M
    try {
      var colA = ms.getRange(2, 1, rows, 1).getValues();
      var colF = ms.getRange(2, 6, rows, 1).getValues();
      var colG = ms.getRange(2, 7, rows, 1).getValues();
      var colK = ms.getRange(2, 11, rows, 1).getValues();
      var colM3 = ms.getRange(2, 13, rows, 1).getValues();
      var loaded = {}, maps = {};
      var srcKeys = Object.keys(COL_V_SOURCES);
      for (var s = 0; s < srcKeys.length; s++) {
        var src = COL_V_SOURCES[srcKeys[s]];
        var ck = src.id + '|' + src.tab;
        if (!loaded[ck]) loaded[ck] = loadColVMap(src);
        maps[srcKeys[s]] = loaded[ck];
      }
      var outW = [], outAC = [];
      for (var i = 0; i < colA.length; i++) {
        var pu = String(colF[i][0]).trim().toUpperCase();
        var pm = maps[pu];
        if (!pm) { outW.push(['']); outAC.push(['']); continue; }
        var lk = (pu === 'MYNTRA')
          ? String(colF[i][0]).trim() + String(colG[i][0]).trim() + String(colK[i][0]).trim()
          : String(colA[i][0]).trim();
        var en = pm[lk];
        if (en && en.v !== '') {
          if (pu === 'MYNTRA') {
            var qty = parseFloat(colM3[i][0]) || 0;
            var val = parseFloat(en.v) || 0;
            outW.push([Math.round(val * qty * 100) / 100]);
          } else {
            outW.push([en.v]);
          }
        } else { outW.push(['']); }
        if (pu === 'MEESHO' && en && en.ab !== '') {
          var abVal = parseFloat(en.ab) || 0;
          var abQty = parseFloat(colM3[i][0]) || 0;
          outAC.push([Math.round(abVal * abQty * 100) / 100]);
        } else { outAC.push(['']); }
      }
      ms.getRange(2, 23, outW.length, 1).setValues(outW);
      ms.getRange(2, 29, outAC.length, 1).setValues(outAC);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh ColW ERR: ' + e.message); }

    // Step 4: Tax % (X = col 24)
    try {
      var taxMap = loadTaxMap();
      var plats4 = ms.getRange(2, 6, rows, 1).getValues();
      var keys4 = ms.getRange(2, 15, rows, 1).getValues();
      var taxOut = [];
      for (var i = 0; i < keys4.length; i++) {
        var np = normalizePlat(String(plats4[i][0]).trim());
        var e = taxMap[String(keys4[i][0]).trim()];
        if (e && np && e[np] !== '') { taxOut.push([e[np]]); }
        else { taxOut.push(['']); }
      }
      ms.getRange(2, 24, taxOut.length, 1).setValues(taxOut);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh Tax% ERR: ' + e.message); }

    // Step 5: Tax Calc (Y = col 25, Z = col 26)
    try {
      var colWv = ms.getRange(2, 23, rows, 1).getValues();
      var colXv = ms.getRange(2, 24, rows, 1).getValues();
      var yOut = [], zOut = [];
      for (var i = 0; i < colWv.length; i++) {
        var w = parseFloat(colWv[i][0]);
        var x = parseFloat(colXv[i][0]);
        if (!isNaN(w) && !isNaN(x) && (100 + x) !== 0) {
          var taxable = w / (100 + x) * 100;
          yOut.push([Math.round((w - taxable) * 100) / 100]);
          zOut.push([Math.round(taxable * 100) / 100]);
        } else { yOut.push(['']); zOut.push(['']); }
      }
      ms.getRange(2, 25, yOut.length, 1).setValues(yOut);
      ms.getRange(2, 26, zOut.length, 1).setValues(zOut);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh TaxCalc ERR: ' + e.message); }

    // Step 6: Col AA (PRICES col F × Qty, col 27)
    try {
      var colFMap = loadPricesColFMap();
      var colOz = ms.getRange(2, 15, rows, 1).getValues();
      var colMz = ms.getRange(2, 13, rows, 1).getValues();
      var aaOut = [];
      for (var i = 0; i < colOz.length; i++) {
        var key = String(colOz[i][0]).trim();
        var qty = parseFloat(colMz[i][0]) || 0;
        var val = colFMap[key];
        if (val !== undefined && val !== '') {
          var num = parseFloat(val) || 0;
          aaOut.push([Math.round(num * qty * 100) / 100]);
        } else { aaOut.push(['']); }
      }
      ms.getRange(2, 27, aaOut.length, 1).setValues(aaOut);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh ColAA ERR: ' + e.message); }

    // Step 7: AB = Z - AA (col 28)
    try {
      var colZab = ms.getRange(2, 26, rows, 1).getValues();
      var colAAab = ms.getRange(2, 27, rows, 1).getValues();
      var abOut = [];
      for (var i = 0; i < colZab.length; i++) {
        var zv = parseFloat(colZab[i][0]);
        var aav = parseFloat(colAAab[i][0]);
        if (!isNaN(zv) && !isNaN(aav)) {
          abOut.push([Math.round((zv - aav) * 100) / 100]);
        } else { abOut.push(['']); }
      }
      ms.getRange(2, 28, abOut.length, 1).setValues(abOut);
      SpreadsheetApp.flush();
    } catch(e) { Logger.log('autoRefresh AB ERR: ' + e.message); }

    // Dedup
    try {
      var lrFinal = ms.getLastRow();
      if (lrFinal > 1) {
        var lc = ms.getLastColumn();
        var data = ms.getRange(2, 1, lrFinal - 1, lc).getValues();
        var seen = {}, keep = [], dupes = 0;
        for (var i = 0; i < data.length; i++) {
          var k = String(data[i][0]).trim();
          if (!k) continue;
          if (seen[k]) { dupes++; continue; }
          seen[k] = true;
          keep.push(data[i]);
        }
        if (dupes > 0) {
          ms.getRange(2, 1, lrFinal - 1, lc).clearContent();
          if (keep.length > 0) ms.getRange(2, 1, keep.length, lc).setValues(keep);
        }
      }
    } catch(e) { Logger.log('autoRefresh Dedup ERR: ' + e.message); }

    Logger.log('autoRefresh complete: synced ' + totalA + ' new rows');
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// DAILY PRICING ANALYSIS - doGet() API
// Returns aggregated financial data as JSON for the dashboard
// Deploy as Web App (Execute as: Me, Access: Anyone)
// ============================================================

function doGet(e) {
  try {
    // DIAGNOSTIC: ?test=1 returns minimal response to verify deployment
    if (e && e.parameter && e.parameter.test === '1') {
      return ContentService.createTextOutput(JSON.stringify({ok:true,ts:new Date().toISOString()})).setMimeType(ContentService.MimeType.JSON);
    }
    var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var tab = 'MASTER CSV';
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var rowsMode = (e && e.parameter && e.parameter.mode === 'rows');
    var daysParam = (e && e.parameter && e.parameter.days !== undefined) ? parseInt(e.parameter.days, 10) : -1;
    var filterStartStr = '', filterEndStr = '';
    if (daysParam === 0) {
      var fd = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      filterStartStr = Utilities.formatDate(fd, tz, 'yyyy-MM-dd');
      filterEndStr = Utilities.formatDate(new Date(fd.getTime()+86399000), tz, 'yyyy-MM-dd');
    } else if (daysParam === 1) {
      var yd = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
      filterStartStr = Utilities.formatDate(yd, tz, 'yyyy-MM-dd');
      filterEndStr = filterStartStr;
    } else if (daysParam > 1) {
      filterStartStr = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysParam), tz, 'yyyy-MM-dd');
      filterEndStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    }
    // ROWS MODE: only fetch columns needed for pricing drill-down
    if (rowsMode) {
      var colRangesR = ['A2:A','E2:E','F2:F','G2:G','M2:M','O2:O','Q2:Q','V2:V','W2:W','Y2:Y','Z2:Z','AA2:AA','AB2:AB','H2:H','I2:I','J2:J','K2:K','L2:L','N2:N','P2:P','R2:R','S2:S','T2:T','U2:U','X2:X','AC2:AC'];
      var rangesR = [];
      for (var i = 0; i < colRangesR.length; i++) rangesR.push("'" + tab + "'!" + colRangesR[i]);
      var resR = Sheets.Spreadsheets.Values.batchGet(ssId, {ranges: rangesR, valueRenderOption: 'UNFORMATTED_VALUE'});
      var vrR = resR.valueRanges;
      var maxR = 0;
      for (var v = 0; v < vrR.length; v++) { var len = (vrR[v].values || []).length; if (len > maxR) maxR = len; }
      var cr = [];
      for (var c = 0; c < vrR.length; c++) cr.push(vrR[c].values || []);
      // cr: 0=unique,1=date,2=platform,3=company,4=qty,5=masterSku,6=category,7=required,8=gross,9=taxAmt,10=net,11=cop,12=pl,13=orderId,14=trackingId,15=companyId,16=skuId,17=subOrderId,18=courier,19=singleCombo,20=subCategory,21=color,22=product,23=link,24=taxPct,25=actualPrice
      var rType = (e.parameter.type || 'short').toLowerCase();
      var rPlat = e.parameter.platform || '';
      var rComp = e.parameter.company || '';
      var matchRows = [];
      for (var r = 0; r < maxR; r++) {
        var unique = r < cr[0].length && cr[0][r].length > 0 ? cr[0][r][0] : '';
        if (!unique) continue;
        var qty = _toNum(r < cr[4].length && cr[4][r].length > 0 ? cr[4][r][0] : 0);
        var required = _toNum(r < cr[7].length && cr[7][r].length > 0 ? cr[7][r][0] : 0);
        var gross = _toNum(r < cr[8].length && cr[8][r].length > 0 ? cr[8][r][0] : 0);
        if (required <= 0 || gross === required) continue;
        var pType = gross < required ? 'short' : 'excess';
        if (pType !== rType) continue;
        var platform = String(r < cr[2].length && cr[2][r].length > 0 ? cr[2][r][0] : '').trim() || 'Unknown';
        var company = String(r < cr[3].length && cr[3][r].length > 0 ? cr[3][r][0] : '').trim() || 'Unknown';
        if (rPlat && platform !== rPlat) continue;
        if (rComp && company !== rComp) continue;
        var dateVal = r < cr[1].length && cr[1][r].length > 0 ? cr[1][r][0] : '';
        var dateStr = '';
        if (typeof dateVal === 'number') { dateStr = Utilities.formatDate(new Date(new Date(1899,11,30).getTime()+dateVal*86400000), tz, 'yyyy-MM-dd'); }
        else if (dateVal instanceof Date && !isNaN(dateVal.getTime())) { dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd'); }
        else if (dateVal) { var pp = new Date(dateVal); if (!isNaN(pp.getTime())) dateStr = Utilities.formatDate(pp, tz, 'yyyy-MM-dd'); }
        if (filterStartStr && (!dateStr || dateStr < filterStartStr || dateStr > filterEndStr)) continue;
        var diff = Math.round((gross - required) * 100) / 100;
        var tax = _toNum(r < cr[9].length && cr[9][r].length > 0 ? cr[9][r][0] : 0);
        var net = _toNum(r < cr[10].length && cr[10][r].length > 0 ? cr[10][r][0] : 0);
        var cop = _toNum(r < cr[11].length && cr[11][r].length > 0 ? cr[11][r][0] : 0);
        var pl = _toNum(r < cr[12].length && cr[12][r].length > 0 ? cr[12][r][0] : 0);
        var taxPct = _toNum(r < cr[24].length && cr[24][r].length > 0 ? cr[24][r][0] : 0);
        var actualPrice = _toNum(r < cr[25].length && cr[25][r].length > 0 ? cr[25][r][0] : 0);
        matchRows.push({u:String(unique),oi:String(r<cr[13].length&&cr[13][r].length>0?cr[13][r][0]:''),ti:String(r<cr[14].length&&cr[14][r].length>0?cr[14][r][0]:''),ci:String(r<cr[15].length&&cr[15][r].length>0?cr[15][r][0]:''),si:String(r<cr[16].length&&cr[16][r].length>0?cr[16][r][0]:''),so:String(r<cr[17].length&&cr[17][r].length>0?cr[17][r][0]:''),q:qty,cr:String(r<cr[18].length&&cr[18][r].length>0?cr[18][r][0]:''),sk:String(r<cr[5].length&&cr[5][r].length>0?cr[5][r][0]:''),sc:String(r<cr[19].length&&cr[19][r].length>0?cr[19][r][0]:''),ca:String(r<cr[6].length&&cr[6][r].length>0?cr[6][r][0]:''),sb:String(r<cr[20].length&&cr[20][r].length>0?cr[20][r][0]:''),co:String(r<cr[21].length&&cr[21][r].length>0?cr[21][r][0]:''),pr:String(r<cr[22].length&&cr[22][r].length>0?cr[22][r][0]:''),rq:required,gr:gross,tp:taxPct,ta:tax,n:net,cp:cop,p:pl,ap:actualPrice,sp:0,ep:0,d:diff,dt:dateStr,pf:platform,cm:company});
      }
      matchRows.sort(function(a,b){return Math.abs(b.d)-Math.abs(a.d);});
      if (matchRows.length > 500) matchRows = matchRows.slice(0, 500);
      return ContentService.createTextOutput(JSON.stringify({rows:matchRows})).setMimeType(ContentService.MimeType.JSON);
    }
    // MAIN MODE: only fetch 11 summary columns (no detail columns)
    var colRanges = ['A2:A','C2:C','D2:D','E2:E','F2:F','G2:G','M2:M','O2:O','Q2:Q','V2:V','W2:W','Y2:Y','Z2:Z','AA2:AA','AB2:AB'];
    var ranges = [];
    for (var i = 0; i < colRanges.length; i++) ranges.push("'" + tab + "'!" + colRanges[i]);
    var res = Sheets.Spreadsheets.Values.batchGet(ssId, {ranges: ranges, valueRenderOption: 'UNFORMATTED_VALUE'});
    var vr = res.valueRanges;
    var maxRows = 0;
    for (var v = 0; v < vr.length; v++) { var len = (vr[v].values || []).length; if (len > maxRows) maxRows = len; }
    if (maxRows === 0) {
      return _jsonResp({summary:{},byPlatform:{},byCompany:{},byMonth:[],byCategory:{},daily:[],topSkus:[],platforms:[],companies:[],pricing:{short:{count:0,totalDiff:0,byPlatform:{}},excess:{count:0,totalDiff:0,byPlatform:{}}},timestamp:new Date().toISOString()});
    }
    var cols = [];
    for (var c = 0; c < vr.length; c++) cols.push(vr[c].values || []);
    // cols: 0=unique,1=year,2=month,3=date,4=platform,5=company,6=qty,7=masterSku,8=category,9=required,10=gross,11=taxAmt,12=net,13=cop,14=pl
    var summary = {totalOrders:0,totalQty:0,grossReceived:0,netReceived:0,taxAmount:0,cop:0,pl:0};
    var byPlatform={},byCompany={},byMonthMap={},byCategory={},bySku={},dailyMap={};
    var seenOrders={},seenByPlatform={},seenByCompany={},seenByMonth={},seenByCategory={},seenByDaily={};
    var pricing={short:{count:0,totalDiff:0,byPlatform:{}},excess:{count:0,totalDiff:0,byPlatform:{}}};
    for (var r = 0; r < maxRows; r++) {
      var unique = r < cols[0].length && cols[0][r].length > 0 ? cols[0][r][0] : '';
      if (!unique || String(unique).trim() === '') continue;
      var qty = _toNum(r < cols[6].length && cols[6][r].length > 0 ? cols[6][r][0] : 0);
      var required = _toNum(r < cols[9].length && cols[9][r].length > 0 ? cols[9][r][0] : 0);
      var gross = _toNum(r < cols[10].length && cols[10][r].length > 0 ? cols[10][r][0] : 0);
      var tax = _toNum(r < cols[11].length && cols[11][r].length > 0 ? cols[11][r][0] : 0);
      var net = _toNum(r < cols[12].length && cols[12][r].length > 0 ? cols[12][r][0] : 0);
      var cop = _toNum(r < cols[13].length && cols[13][r].length > 0 ? cols[13][r][0] : 0);
      var pl = _toNum(r < cols[14].length && cols[14][r].length > 0 ? cols[14][r][0] : 0);
      var platform = String(r < cols[4].length && cols[4][r].length > 0 ? cols[4][r][0] : '').trim() || 'Unknown';
      var company = String(r < cols[5].length && cols[5][r].length > 0 ? cols[5][r][0] : '').trim() || 'Unknown';
      var category = String(r < cols[8].length && cols[8][r].length > 0 ? cols[8][r][0] : '').trim() || 'Unknown';
      var sku = String(r < cols[7].length && cols[7][r].length > 0 ? cols[7][r][0] : '').trim() || 'Unknown';
      var dateVal = r < cols[3].length && cols[3][r].length > 0 ? cols[3][r][0] : '';
      var dateStr = ''; var monthKey = '';
      if (typeof dateVal === 'number') {
        var d = new Date(new Date(1899,11,30).getTime() + dateVal * 86400000);
        dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        monthKey = Utilities.formatDate(d, tz, 'yyyy-MM');
      } else if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
        dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
        monthKey = Utilities.formatDate(dateVal, tz, 'yyyy-MM');
      } else if (dateVal) {
        var parsed = new Date(dateVal);
        if (!isNaN(parsed.getTime())) { dateStr = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd'); monthKey = Utilities.formatDate(parsed, tz, 'yyyy-MM'); }
      }
      if (filterStartStr && (!dateStr || dateStr < filterStartStr || dateStr > filterEndStr)) continue;
      if (!monthKey) {
        var yr = String(r < cols[1].length && cols[1][r].length > 0 ? cols[1][r][0] : '').trim();
        var mo = String(r < cols[2].length && cols[2][r].length > 0 ? cols[2][r][0] : '').trim();
        if (yr && mo) monthKey = yr + '-' + _monthToNum(mo);
      }
      if (!seenOrders[unique]) { seenOrders[unique]=true; summary.totalOrders++; }
      summary.totalQty+=qty; summary.grossReceived+=gross; summary.netReceived+=net; summary.taxAmount+=tax; summary.cop+=cop; summary.pl+=pl;
      if (!byPlatform[platform]) { byPlatform[platform]={orders:0,qty:0,gross:0,net:0,tax:0,cop:0,pl:0}; seenByPlatform[platform]={}; }
      if (!seenByPlatform[platform][unique]) { seenByPlatform[platform][unique]=true; byPlatform[platform].orders++; }
      byPlatform[platform].qty+=qty; byPlatform[platform].gross+=gross; byPlatform[platform].net+=net; byPlatform[platform].tax+=tax; byPlatform[platform].cop+=cop; byPlatform[platform].pl+=pl;
      if (!byCompany[company]) { byCompany[company]={orders:0,qty:0,gross:0,net:0,tax:0,cop:0,pl:0}; seenByCompany[company]={}; }
      if (!seenByCompany[company][unique]) { seenByCompany[company][unique]=true; byCompany[company].orders++; }
      byCompany[company].qty+=qty; byCompany[company].gross+=gross; byCompany[company].net+=net; byCompany[company].tax+=tax; byCompany[company].cop+=cop; byCompany[company].pl+=pl;
      if (monthKey) { if (!byMonthMap[monthKey]) { byMonthMap[monthKey]={month:monthKey,orders:0,qty:0,gross:0,net:0,tax:0,cop:0,pl:0}; seenByMonth[monthKey]={}; } if (!seenByMonth[monthKey][unique]) { seenByMonth[monthKey][unique]=true; byMonthMap[monthKey].orders++; } byMonthMap[monthKey].qty+=qty; byMonthMap[monthKey].gross+=gross; byMonthMap[monthKey].net+=net; byMonthMap[monthKey].tax+=tax; byMonthMap[monthKey].cop+=cop; byMonthMap[monthKey].pl+=pl; }
      if (!byCategory[category]) { byCategory[category]={orders:0,qty:0,gross:0,net:0,cop:0,pl:0}; seenByCategory[category]={}; }
      if (!seenByCategory[category][unique]) { seenByCategory[category][unique]=true; byCategory[category].orders++; }
      byCategory[category].qty+=qty; byCategory[category].gross+=gross; byCategory[category].net+=net; byCategory[category].cop+=cop; byCategory[category].pl+=pl;
      if (dateStr) {
        if (!dailyMap[dateStr]) { dailyMap[dateStr]={d:dateStr,o:0,q:0,g:0,n:0,t:0,c:0,p:0}; seenByDaily[dateStr]={}; }
        if (!seenByDaily[dateStr][unique]) { seenByDaily[dateStr][unique]=true; dailyMap[dateStr].o++; }
        dailyMap[dateStr].q+=qty; dailyMap[dateStr].g+=gross; dailyMap[dateStr].n+=net; dailyMap[dateStr].t+=tax; dailyMap[dateStr].c+=cop; dailyMap[dateStr].p+=pl;
      }
      if (!bySku[sku]) bySku[sku]={sku:sku,platform:platform,orders:0,qty:0,gross:0,net:0,cop:0,pl:0};
      bySku[sku].orders++; bySku[sku].qty+=qty; bySku[sku].gross+=gross; bySku[sku].net+=net; bySku[sku].cop+=cop; bySku[sku].pl+=pl;
      // Pricing: only counts and diffs, NO row objects
      if (required > 0 && gross !== required) {
        var pType = gross < required ? 'short' : 'excess';
        var diff = Math.round((gross - required) * 100) / 100;
        pricing[pType].count++;
        pricing[pType].totalDiff += diff;
        if (!pricing[pType].byPlatform[platform]) pricing[pType].byPlatform[platform] = {count:0,diff:0,byCompany:{}};
        pricing[pType].byPlatform[platform].count++;
        pricing[pType].byPlatform[platform].diff += diff;
        if (!pricing[pType].byPlatform[platform].byCompany[company]) pricing[pType].byPlatform[platform].byCompany[company] = {count:0,diff:0};
        pricing[pType].byPlatform[platform].byCompany[company].count++;
        pricing[pType].byPlatform[platform].byCompany[company].diff += diff;
      }
    }
    summary.avgOrderValue = summary.totalOrders > 0 ? summary.grossReceived / summary.totalOrders : 0;
    summary.margin = summary.grossReceived > 0 ? (summary.pl / summary.grossReceived * 100) : 0;
    var byMonth = []; var mKeys = Object.keys(byMonthMap).sort();
    for (var i = 0; i < mKeys.length; i++) byMonth.push(byMonthMap[mKeys[i]]);
    var daily = []; var dailyKeys = Object.keys(dailyMap).sort();
    for (var i = 0; i < dailyKeys.length; i++) daily.push(dailyMap[dailyKeys[i]]);
    var skuList = []; for (var k in bySku) skuList.push(bySku[k]);
    skuList.sort(function(a,b){return b.gross-a.gross;}); var topSkus = skuList.slice(0, 20);
    var platforms = Object.keys(byPlatform).sort(); var companies = Object.keys(byCompany).sort();
    _roundObj(summary); for(var k in byPlatform)_roundObj(byPlatform[k]); for(var k in byCompany)_roundObj(byCompany[k]);
    for(var i=0;i<byMonth.length;i++)_roundObj(byMonth[i]); for(var k in byCategory)_roundObj(byCategory[k]);
    for(var i=0;i<topSkus.length;i++)_roundObj(topSkus[i]);
    // Round daily
    for(var i=0;i<daily.length;i++){var dd=daily[i];dd.o=Math.round(dd.o);dd.q=Math.round(dd.q);dd.g=Math.round(dd.g*100)/100;dd.n=Math.round(dd.n*100)/100;dd.t=Math.round(dd.t*100)/100;dd.c=Math.round(dd.c*100)/100;dd.p=Math.round(dd.p*100)/100;}
    // Round pricing diffs
    pricing.short.totalDiff=Math.round(pricing.short.totalDiff*100)/100;
    pricing.excess.totalDiff=Math.round(pricing.excess.totalDiff*100)/100;
    for(var pt in pricing){for(var pp in pricing[pt].byPlatform){pricing[pt].byPlatform[pp].diff=Math.round(pricing[pt].byPlatform[pp].diff*100)/100;for(var pc in pricing[pt].byPlatform[pp].byCompany){pricing[pt].byPlatform[pp].byCompany[pc].diff=Math.round(pricing[pt].byPlatform[pp].byCompany[pc].diff*100)/100;}}}
    var result = {summary:summary,byPlatform:byPlatform,byCompany:byCompany,byMonth:byMonth,byCategory:byCategory,daily:daily,topSkus:topSkus,platforms:platforms,companies:companies,pricing:pricing,timestamp:new Date().toISOString()};
    var jsonStr = JSON.stringify(result);
    // Size check: if > 5MB, strip daily to save space
    if (jsonStr.length > 5000000) {
      result.daily = daily.slice(-90); // only last 90 days
      result._truncated = true;
      result._fullSize = jsonStr.length;
      jsonStr = JSON.stringify(result);
    }
    // If STILL > 5MB, strip more
    if (jsonStr.length > 5000000) {
      result.daily = daily.slice(-30);
      result.byCategory = {};
      result.topSkus = topSkus.slice(0, 10);
      jsonStr = JSON.stringify(result);
    }
    return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
  } catch (err) { return _jsonResp({ error: err.message }); }
}

function _jsonResp(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function _toNum(v) { if (v===''||v===null||v===undefined) return 0; var n=Number(v); return isNaN(n)?0:n; }
function _roundObj(obj) { for (var k in obj) { if (typeof obj[k]==='number') obj[k]=Math.round(obj[k]*100)/100; } }
function _monthToNum(m) { var months={'january':'01','february':'02','march':'03','april':'04','may':'05','june':'06','july':'07','august':'08','september':'09','october':'10','november':'11','december':'12'}; return months[String(m).toLowerCase().trim()]||'01'; }

function testFinancialApi() {
  var result = doGet();
  var json = JSON.parse(result.getContent());
  Logger.log('Total Orders: ' + json.summary.totalOrders);
  Logger.log('Gross Received: ' + json.summary.grossReceived);
  Logger.log('P/L: ' + json.summary.pl);
  Logger.log('Platforms: ' + json.platforms.join(', '));
  Logger.log('Months: ' + json.byMonth.length);
  SpreadsheetApp.getActiveSpreadsheet().toast('Orders: ' + json.summary.totalOrders + ' | Revenue: ' + json.summary.grossReceived, 'API Test OK', 10);
}
