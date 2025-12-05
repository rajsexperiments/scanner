/**
 * StockLens Backend for Google Sheets
 * 
 * Optimized version with concurrency protection and streamlined features.
 */

// --- CONFIGURATION ---
const LOG_SHEET_NAME = 'Scanned Inventory Log';
const PRODUCT_SHEET_NAME = 'Master Product List';
const SUMMARY_SHEET_NAME = 'Live Inventory Dashboard';
const USER_SHEET_NAME = 'Users';
const B2B_CLIENTS_SHEET_NAME = 'B2B Clients';
const CACHE_EXPIRATION_SECONDS = 30; // 30 sec (reduced for fresher data)
const LOCK_TIMEOUT_MS = 3000; // 3 seconds

// --- WEB APP ENTRY POINTS ---

/**
 * Handles POST requests to the web app.
 */
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let result;
    
    switch (action) {
      case 'addScan':
        result = addScan(request.payload);
        break;
      case 'addProduct':
        result = addProduct(request.payload);
        break;
      case 'deleteProduct':
        result = deleteProduct(request.payload);
        break;
      case 'clearLogs':
        result = clearLogs();
        break;
      default:
        throw new Error('Invalid action for POST request: ' + action);
    }
    
    return createJsonResponse({ success: true, data: result });
  } catch (error) {
    Logger.log('[ERROR] doPost: ' + error.message);
    return createJsonResponse({ success: false, error: error.message });
  }
}

/**
 * Handles GET requests to the web app.
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;
    
    switch (action) {
      case 'getLogs':
        result = withCache('logs', getLogs);
        break;
      case 'getProducts':
        result = withCache('products', getProducts);
        break;
      case 'getSummary':
        result = withCache('summary', getSummary);
        break;
      case 'getUsers':
        result = withCache('users', getUsers);
        break;
      case 'getB2BClients':
        result = withCache('b2b_clients', getB2BClients);
        break;
      default:
        throw new Error('Invalid action for GET request: ' + action);
    }
    
    return createJsonResponse({ success: true, data: result });
  } catch (error) {
    Logger.log('[ERROR] doGet: ' + error.message);
    return createJsonResponse({ success: false, error: error.message });
  }
}

// --- HELPER FUNCTIONS ---

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetAndCreate(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  
  return sheet;
}

/**
 * Execute a function with lock protection to prevent race conditions.
 */
function withLock(fn) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
    const result = fn();
    return result;
  } catch (error) {
    Logger.log('[ERROR] Lock acquisition failed: ' + error.message);
    throw new Error('System is busy. Please try again.');
  } finally {
    lock.releaseLock();
  }
}

// --- CACHING LOGIC ---

function withCache(key, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  
  if (cached != null) {
    return JSON.parse(cached);
  }
  
  const result = fetchFunction();
  cache.put(key, JSON.stringify(result), CACHE_EXPIRATION_SECONDS);
  return result;
}

function clearCache(keys) {
  const cache = CacheService.getScriptCache();
  if (Array.isArray(keys)) {
    cache.removeAll(keys);
  } else {
    cache.remove(keys);
  }
}

// --- CORE LOGIC: SCANS ---

function addScan(payload) {
  return withLock(function() {
    if (!payload || !payload.serialNumber || !payload.scanEvent || !payload.location) {
      throw new Error('Serial number, scan event, and location are required.');
    }
    
    const sheet = getSheetAndCreate(LOG_SHEET_NAME, ['Timestamp', 'SerialNumber', 'ScanEvent', 'Location', 'B2BClientId']);
    const timestamp = new Date();
    const b2bClientId = payload.b2bClientId || '';
    
    sheet.appendRow([timestamp, payload.serialNumber, payload.scanEvent, payload.location, b2bClientId]);
    
    clearCache(['logs', 'summary']);
    
    return {
      serialNumber: payload.serialNumber,
      timestamp: timestamp.toISOString(),
      scanEvent: payload.scanEvent,
      location: payload.location,
      b2bClientId: b2bClientId
    };
  });
}

function getLogs() {
  const sheet = getSheetAndCreate(LOG_SHEET_NAME, ['Timestamp', 'SerialNumber', 'ScanEvent', 'Location', 'B2BClientId']);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).trim());
  const timestampIdx = headers.indexOf('Timestamp');
  const serialIdx = headers.indexOf('SerialNumber');
  const eventIdx = headers.indexOf('ScanEvent');
  const locationIdx = headers.indexOf('Location');
  const clientIdx = headers.indexOf('B2BClientId');
  
  return data.slice(1).map(row => ({
    timestamp: new Date(row[timestampIdx]).toISOString(),
    serialNumber: row[serialIdx],
    scanEvent: row[eventIdx],
    location: row[locationIdx] || 'N/A',
    b2bClientId: row[clientIdx] || undefined
  })).reverse();
}

function clearLogs() {
  return withLock(function() {
    const sheet = getSheetAndCreate(LOG_SHEET_NAME);
    
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    clearCache(['logs', 'summary']);
    
    return { message: 'Logs cleared successfully.' };
  });
}

// --- CORE LOGIC: USERS & CLIENTS ---

function getUsers() {
  const sheet = getSheetAndCreate(USER_SHEET_NAME, ['email', 'name', 'role', 'location', 'password']);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).trim());
  
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => { obj[header] = row[i]; });
    return obj;
  }).filter(u => u.email);
}

function getB2BClients() {
  const sheet = getSheetAndCreate(B2B_CLIENTS_SHEET_NAME, ['clientId', 'clientName', 'contactPerson', 'address']);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).trim());
  
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => { obj[header] = row[i]; });
    return obj;
  }).filter(c => c.clientId);
}

// --- CORE LOGIC: PRODUCTS ---

function getProducts() {
  const sheet = getSheetAndCreate(PRODUCT_SHEET_NAME, ['id', 'name', 'category', 'unitOfMeasure', 'unitCost', 'supplierName', 'reorderLevel', 'reorderQuantity', 'storageLocation', 'shelfLifeDays', 'isPerishable']);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).trim());
  const idIndex = headers.indexOf('id');
  
  if (idIndex === -1) throw new Error('Product sheet must have an "id" column.');
  
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      if (header === 'isPerishable') {
        obj[header] = row[i] === true || String(row[i]).toLowerCase() === 'true';
      } else if (['unitCost', 'reorderLevel', 'reorderQuantity', 'shelfLifeDays'].includes(header)) {
        obj[header] = parseFloat(row[i]) || 0;
      } else {
        obj[header] = row[i];
      }
    });
    return obj;
  }).filter(p => p.id);
}

function addProduct(product) {
  return withLock(function() {
    if (!product.id) throw new Error('Product ID is required.');
    
    const sheet = getSheetAndCreate(PRODUCT_SHEET_NAME, ['id', 'name', 'category', 'unitOfMeasure', 'unitCost', 'supplierName', 'reorderLevel', 'reorderQuantity', 'storageLocation', 'shelfLifeDays', 'isPerishable']);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idIndex = headers.indexOf('id');
    const data = sheet.getDataRange().getValues();
    const existingRowIndex = data.findIndex(row => row[idIndex] === product.id);
    const newRow = headers.map(header => product[header] !== undefined ? product[header] : "");
    
    if (existingRowIndex > -1) {
      sheet.getRange(existingRowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    
    clearCache(['products', 'summary']);
    
    return getProducts();
  });
}

function deleteProduct(payload) {
  return withLock(function() {
    if (!payload || !payload.productId) {
      throw new Error('Product ID is required for deletion.');
    }
    
    const sheet = getSheetAndCreate(PRODUCT_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idIndex = headers.indexOf('id');
    const data = sheet.getDataRange().getValues();
    const rowIndexToDelete = data.findIndex(row => row[idIndex] === payload.productId);
    
    if (rowIndexToDelete > -1) {
      sheet.deleteRow(rowIndexToDelete + 1);
    } else {
      throw new Error('Product not found.');
    }
    
    clearCache(['products', 'summary']);
    
    return getProducts();
  });
}

// --- CORE LOGIC: SUMMARY ---

function getSummary() {
  const sheet = getSheetAndCreate(SUMMARY_SHEET_NAME, ['productId', 'productName', 'count']);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).trim());
  const productIdIdx = headers.indexOf('productId');
  const productNameIdx = headers.indexOf('productName');
  const countIdx = headers.indexOf('count');
  
  return data.slice(1).map(row => ({
    productId: row[productIdIdx],
    productName: row[productNameIdx],
    count: parseInt(row[countIdx], 10) || 0
  })).filter(item => item.productId && item.count > 0);
}
