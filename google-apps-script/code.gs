/**
 * StockLens Backend for Google Sheets - COMPLETE VERSION WITH AUTO-SYNC
 * 
 * This version includes automatic product synchronization between Master Product List
 * and Live Inventory Dashboard, plus all previous functionality with comprehensive logging.
 */


// --- CONFIGURATION ---
const LOG_SHEET_NAME = 'Scanned Inventory Log';
const PRODUCT_SHEET_NAME = 'Master Product List';
const LIVE_INVENTORY_DASHBOARD_SHEET_NAME = 'Live Inventory Dashboard';
const USER_SHEET_NAME = 'Users';
const CAKE_STATUS_SHEET_NAME = 'Cake Status Dashboard';
const LIVE_OPS_SHEET_NAME = 'Live Operations Dashboard';
const B2B_CLIENTS_SHEET_NAME = 'B2B Clients';
const WEEKLY_SALES_REPORT_SHEET_NAME = 'Weekly Sales Report';
const CACHE_EXPIRATION_SECONDS = 10;


// --- WEB APP ENTRY POINTS ---
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
    Logger.log('ERROR in doPost: ' + error.message);
    Logger.log(error.stack);
    return createJsonResponse({ success: false, error: error.message });
  }
}


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
      case 'getCakeStatus':
        result = withCache('cake_status', getCakeStatus);
        break;
      case 'getLiveOperationsData':
        result = withCache('live_ops', getLiveOperationsData);
        break;
      case 'getB2BClients':
        result = withCache('b2b_clients', getB2BClients);
        break;
      case 'generateWeeklySalesReport':
        result = generateWeeklySalesReport();
        break;
      case 'syncProducts':
        result = syncProductsToLiveInventoryDashboard();
        break;
      default:
        throw new Error('Invalid action for GET request: ' + action);
    }
    return createJsonResponse({ success: true, data: result });
  } catch (error) {
    Logger.log('ERROR in doGet: ' + error.message);
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
    }
  }
  return sheet;
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
  try {
    if (!payload || !payload.serialNumber || !payload.scanEvent || !payload.location) {
      throw new Error('Serial number, scan event, and location are required.');
    }
    
    Logger.log('=== addScan called ===');
    Logger.log('Payload: ' + JSON.stringify(payload));
    
    const sheet = getSheetAndCreate(LOG_SHEET_NAME, ['Timestamp', 'SerialNumber', 'scanEvent', 'Location', 'clientId']);
    const timestamp = new Date();
    const clientId = payload.clientId || '';
    
    sheet.appendRow([timestamp, payload.serialNumber, payload.scanEvent, payload.location, clientId]);
    Logger.log('Scan logged successfully. ClientID written: ' + (clientId || 'None'));

    updateCakeStatus(payload.serialNumber, payload.scanEvent, payload.location, timestamp);
    Logger.log('Cake status updated');
    
    updateLiveInventoryDashboard(payload);
    Logger.log('Live Inventory Dashboard update attempted');
    
    clearCache(['logs', 'summary', 'cake_status', 'live_ops', 'weekly_report']);
    
    return {
      serialNumber: payload.serialNumber,
      timestamp: timestamp.toISOString(),
      scanEvent: payload.scanEvent,
      location: payload.location,
      clientId: clientId
    };
  } catch (error) {
    Logger.log('ERROR in addScan: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}


/**
 * Updates the Live Inventory Dashboard with proper error handling
 */
function updateLiveInventoryDashboard(payload) {
  try {
    const { serialNumber, scanEvent } = payload;
    
    Logger.log('=== updateLiveInventoryDashboard ===');
    Logger.log('SerialNumber: ' + serialNumber);
    Logger.log('ScanEvent: ' + scanEvent);
    
    // Get all products to find matching product ID
    const allProducts = getProducts();
    Logger.log('Total products loaded: ' + allProducts.length);
    
    // Find the product whose ID matches the start of the serial number
    const product = allProducts.find(p => serialNumber.startsWith(p.id));

    if (!product) {
      Logger.log('WARNING: No matching product found for SerialNumber: ' + serialNumber);
      return;
    }
    
    const productId = product.id;
    Logger.log('Matched Product ID: ' + productId);

    const sheet = getSheetAndCreate(LIVE_INVENTORY_DASHBOARD_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      Logger.log('WARNING: Live Inventory Dashboard has no data rows. Auto-syncing products...');
      syncProductsToLiveInventoryDashboard();
      return updateLiveInventoryDashboard(payload); // Retry after sync
    }
    
    const dataRange = sheet.getRange("A2:A" + lastRow);
    const productIdsInSheet = dataRange.getValues().flat();
    Logger.log('Product IDs in sheet: ' + productIdsInSheet.join(', '));
    
    const rowIndex = productIdsInSheet.indexOf(productId) + 2;
    
    if (rowIndex < 2) {
      Logger.log('WARNING: Product ID ' + productId + ' not found in Live Inventory Dashboard. Auto-syncing...');
      syncProductsToLiveInventoryDashboard();
      return updateLiveInventoryDashboard(payload); // Retry after sync
    }
    
    Logger.log('Found product at row: ' + rowIndex);

    // Column positions match your exact sheet layout
    const columnMap = {
      'In Warehouse': 3,        // Column C
      'Boutique Stock': 4,      // Column D
      'Marche Stock': 5,        // Column E
      'Saleya Stock': 6,        // Column F
      'B2B Delivered': 7        // Column G
    };
    
    const updateCell = (colName, change) => {
      const colIndex = columnMap[colName];
      if (colIndex) {
        const cell = sheet.getRange(rowIndex, colIndex);
        const currentValue = Number(cell.getValue()) || 0;
        const newValue = Math.max(0, currentValue + change); // Prevent negative inventory
        cell.setValue(newValue);
        Logger.log('Updated ' + colName + ': ' + currentValue + ' → ' + newValue);
      }
    };

    // Handle all scan events with proper logging
    Logger.log('Processing scanEvent: ' + scanEvent);
    
    switch (scanEvent) {
      case 'PRODUCTION_SCAN':
        updateCell('In Warehouse', 1);
        break;
      case 'BOUTIQUE_STOCK_SCAN':
        updateCell('In Warehouse', -1);
        updateCell('Boutique Stock', 1);
        break;
      case 'MARCHE_STOCK_SCAN':
        updateCell('In Warehouse', -1);
        updateCell('Marche Stock', 1);
        break;
      case 'SALEYA_STOCK_SCAN':
        updateCell('In Warehouse', -1);
        updateCell('Saleya Stock', 1);
        break;
      case 'DELIVERY_B2B':
        updateCell('In Warehouse', -1);
        updateCell('B2B Delivered', 1);
        break;
      case 'SALE_BOUTIQUE':
        updateCell('Boutique Stock', -1);
        break;
      case 'SALE_MARCHE':
        updateCell('Marche Stock', -1);
        break;
      case 'SALE_SALEYA':
        updateCell('Saleya Stock', -1);
        break;
      default:
        Logger.log('WARNING: Unhandled scanEvent: ' + scanEvent);
    }
    
    Logger.log('=== Update complete ===');
    
  } catch (error) {
    Logger.log('ERROR in updateLiveInventoryDashboard: ' + error.message);
    Logger.log(error.stack);
  }
}


/**
 * Syncs Master Product List with Live Inventory Dashboard
 * Automatically adds new products to dashboard with zero initial counts
 */
function syncProductsToLiveInventoryDashboard() {
  try {
    Logger.log('=== Syncing Products to Dashboard ===');
    
    const allProducts = getProducts();
    const dashboardSheet = getSheetAndCreate(LIVE_INVENTORY_DASHBOARD_SHEET_NAME, 
      ['Product ID', 'Product Name', 'In Warehouse', 'Boutique Stock', 'Marche Stock', 'Saleya Stock', 'B2B Delivered']);
    
    // Get existing product IDs in dashboard
    const lastRow = dashboardSheet.getLastRow();
    const existingIds = lastRow >= 2 
      ? dashboardSheet.getRange("A2:A" + lastRow).getValues().flat()
      : [];
    
    Logger.log('Existing products in dashboard: ' + existingIds.length);
    Logger.log('Products in Master List: ' + allProducts.length);
    
    let newProductsAdded = 0;
    
    // Add missing products
    allProducts.forEach(product => {
      if (!existingIds.includes(product.id)) {
        // Add new row: ProductID, ProductName, InWarehouse=0, Boutique=0, Marche=0, Saleya=0, B2B=0
        dashboardSheet.appendRow([
          product.id,
          product.name,
          0, // In Warehouse
          0, // Boutique Stock
          0, // Marche Stock
          0, // Saleya Stock
          0  // B2B Delivered
        ]);
        newProductsAdded++;
        Logger.log('Added new product: ' + product.id + ' - ' + product.name);
      }
    });
    
    Logger.log('=== Sync complete: ' + newProductsAdded + ' new products added ===');
    
    clearCache(['summary']); // Clear summary cache after sync
    
    return { 
      newProductsAdded: newProductsAdded, 
      totalProducts: allProducts.length,
      message: newProductsAdded > 0 
        ? newProductsAdded + ' new product(s) added to dashboard' 
        : 'Dashboard is already up to date'
    };
    
  } catch (error) {
    Logger.log('ERROR in syncProductsToLiveInventoryDashboard: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}


// --- CORE LOGIC: PRODUCTS ---
function getProducts() {
  const sheet = getSheetAndCreate(PRODUCT_SHEET_NAME, ['id', 'name', 'category', 'unitOfMeasure', 'unitCost', 'supplierName', 'reorderLevel', 'reorderQuantity', 'storageLocation', 'shelfLifeDays', 'isPerishable']);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    Logger.log('WARNING: No products found in Master Product List');
    return [];
  }
  
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idIndex = headers.indexOf('id');
  
  if (idIndex === -1) {
    Logger.log('ERROR: "id" column not found in Master Product List');
    throw new Error('Product sheet must have an "id" column.');
  }
  
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      if (header === 'isperishable') {
        obj[header] = row[i] === true || String(row[i]).toLowerCase() === 'true';
      } else if (['unitcost', 'reorderlevel', 'reorderquantity', 'shelflifedays'].includes(header)) {
        obj[header] = parseFloat(row[i]) || 0;
      } else {
        obj[header] = row[i];
      }
    });
    return obj;
  }).filter(p => p.id);
}


function addProduct(payload) {
  try {
    if (!payload || !payload.id || !payload.name) {
      throw new Error('Product ID and name are required.');
    }
    
    Logger.log('=== addProduct called ===');
    Logger.log('Payload: ' + JSON.stringify(payload));
    
    const sheet = getSheetAndCreate(PRODUCT_SHEET_NAME, ['id', 'name', 'category', 'unitOfMeasure', 'unitCost', 'supplierName', 'reorderLevel', 'reorderQuantity', 'storageLocation', 'shelfLifeDays', 'isPerishable']);
    
    // Check if product already exists
    const data = sheet.getDataRange().getValues();
    const existingIds = data.slice(1).map(row => row[0]);
    
    if (existingIds.includes(payload.id)) {
      throw new Error('Product with ID ' + payload.id + ' already exists.');
    }
    
    // Add product to Master Product List
    sheet.appendRow([
      payload.id,
      payload.name,
      payload.category || '',
      payload.unitOfMeasure || 'unit',
      payload.unitCost || 0,
      payload.supplierName || '',
      payload.reorderLevel || 0,
      payload.reorderQuantity || 0,
      payload.storageLocation || '',
      payload.shelfLifeDays || 0,
      payload.isPerishable || false
    ]);
    
    Logger.log('Product added to Master Product List: ' + payload.id);
    
    // Auto-sync to Live Inventory Dashboard
    const syncResult = syncProductsToLiveInventoryDashboard();
    Logger.log('Auto-sync result: ' + JSON.stringify(syncResult));
    
    clearCache(['products', 'summary']);
    
    return {
      id: payload.id,
      name: payload.name,
      message: 'Product added successfully and synced to dashboard',
      syncResult: syncResult
    };
    
  } catch (error) {
    Logger.log('ERROR in addProduct: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}


function deleteProduct(payload) {
  try {
    if (!payload || !payload.id) {
      throw new Error('Product ID is required.');
    }
    
    Logger.log('=== deleteProduct called ===');
    Logger.log('Product ID: ' + payload.id);
    
    const sheet = getSheetAndCreate(PRODUCT_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // Find product row
    const rowIndex = data.findIndex((row, index) => index > 0 && row[0] === payload.id);
    
    if (rowIndex === -1) {
      throw new Error('Product with ID ' + payload.id + ' not found.');
    }
    
    // Delete from Master Product List
    sheet.deleteRow(rowIndex + 1);
    Logger.log('Product deleted from Master Product List');
    
    // Delete from Live Inventory Dashboard
    const dashboardSheet = getSheetAndCreate(LIVE_INVENTORY_DASHBOARD_SHEET_NAME);
    const dashboardData = dashboardSheet.getDataRange().getValues();
    const dashboardRowIndex = dashboardData.findIndex((row, index) => index > 0 && row[0] === payload.id);
    
    if (dashboardRowIndex > 0) {
      dashboardSheet.deleteRow(dashboardRowIndex + 1);
      Logger.log('Product deleted from Live Inventory Dashboard');
    }
    
    clearCache(['products', 'summary']);
    
    return {
      id: payload.id,
      message: 'Product deleted successfully from all sheets'
    };
    
  } catch (error) {
    Logger.log('ERROR in deleteProduct: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}


// --- LOGS ---
function getLogs() {
  const sheet = getSheetAndCreate(LOG_SHEET_NAME, ['Timestamp', 'SerialNumber', 'scanEvent', 'Location', 'clientId']);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const timestampIdx = headers.indexOf('timestamp');
  const serialIdx = headers.indexOf('serialnumber');
  const eventIdx = headers.indexOf('scanevent');
  const locationIdx = headers.indexOf('location');
  const clientIdx = headers.indexOf('clientid');
  
  return data.slice(1).map(row => ({
    timestamp: new Date(row[timestampIdx]).toISOString(),
    serialNumber: row[serialIdx],
    scanEvent: row[eventIdx],
    location: row[locationIdx] || 'N/A',
    clientId: row[clientIdx] || undefined
  })).reverse();
}


function clearLogs() {
  const sheet = getSheetAndCreate(LOG_SHEET_NAME);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  const cakeStatusSheet = getSheetAndCreate(CAKE_STATUS_SHEET_NAME);
  if (cakeStatusSheet.getLastRow() > 1) {
    cakeStatusSheet.getRange(2, 1, cakeStatusSheet.getLastRow() - 1, cakeStatusSheet.getLastColumn()).clearContent();
  }
  clearCache(['logs', 'summary', 'cake_status', 'live_ops']);
  return { message: 'Logs cleared successfully.' };
}


// --- CAKE STATUS ---
function updateCakeStatus(serialNumber, scanEvent, location, timestamp) {
  const sheet = getSheetAndCreate(CAKE_STATUS_SHEET_NAME, ['SerialNumber', 'CurrentLocation', 'Status', 'LastUpdate']);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const serialIndex = headers.indexOf('SerialNumber');
  let status = scanEvent.replace(/_/g, ' ').toUpperCase();
  let rowIndex = data.findIndex(row => row[serialIndex] === serialNumber);
  
  if (rowIndex === -1) {
    sheet.appendRow([serialNumber, location, status, timestamp]);
  } else {
    sheet.getRange(rowIndex + 1, 1, 1, 4).setValues([[serialNumber, location, status, timestamp]]);
  }
}


function getCakeStatus() {
  const sheet = getSheetAndCreate(CAKE_STATUS_SHEET_NAME, ['SerialNumber', 'CurrentLocation', 'Status', 'LastUpdate']);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => ({
    serialNumber: row[headers.indexOf('SerialNumber')],
    currentLocation: row[headers.indexOf('CurrentLocation')],
    status: row[headers.indexOf('Status')],
    lastUpdate: (date => date.getTime() ? date.toISOString() : null)(new Date(row[headers.indexOf('LastUpdate')]))
  }));
}


// --- SUMMARY ---
function getSummary() {
  const sheet = getSheetAndCreate(LIVE_INVENTORY_DASHBOARD_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => ({
    productId: row[headers.indexOf('Product ID')],
    productName: row[headers.indexOf('Product Name')],
    inWarehouse: parseInt(row[headers.indexOf('In Warehouse')], 10) || 0,
    boutiqueStock: parseInt(row[headers.indexOf('Boutique Stock')], 10) || 0,
    marcheStock: parseInt(row[headers.indexOf('Marche Stock')], 10) || 0,
    saleyaStock: parseInt(row[headers.indexOf('Saleya Stock')], 10) || 0,
    b2bDelivered: parseInt(row[headers.indexOf('B2B Delivered')], 10) || 0
  })).filter(item => item.productId);
}


// --- USERS ---
function getUsers() {
  const sheet = getSheetAndCreate(USER_SHEET_NAME, ['email', 'name', 'role', 'location', 'password']);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.trim());
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => { obj[header] = row[i]; });
    return obj;
  }).filter(u => u.email);
}


// --- B2B CLIENTS ---
function getB2BClients() {
  const sheet = getSheetAndCreate(B2B_CLIENTS_SHEET_NAME, ['clientId', 'clientName', 'contactPerson', 'address']);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.trim());
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => { obj[header] = row[i]; });
    return obj;
  }).filter(c => c.clientId);
}


// --- LIVE OPERATIONS ---
function getLiveOperationsData() {
  const sheet = getSheetAndCreate(LIVE_OPS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  
  const headers = data[0].map(h => h.toString().trim());
  const productIds = headers.slice(2);
  const result = {};
  
  data.slice(1).forEach(row => {
    const metric = row[0].toString().trim();
    if (!metric) return;
    
    const formattedMetric = metric.replace(/\s+/g, '_').toLowerCase();
    result[formattedMetric] = {
      total: row[1] || 0,
      by_product: {}
    };
    
    productIds.forEach((productId, i) => {
      if (productId) {
        result[formattedMetric].by_product[productId] = row[i + 2] || 0;
      }
    });
  });
  
  return result;
}


// --- WEEKLY SALES REPORT ---
function generateWeeklySalesReport() {
  try {
    Logger.log('=== Generating Weekly Sales Report ===');
    
    const logs = getLogs();
    const products = getProducts();
    const reportSheet = getSheetAndCreate(WEEKLY_SALES_REPORT_SHEET_NAME);
    
    // Clear existing data
    if (reportSheet.getLastRow() > 0) {
      reportSheet.clear();
    }
    
    // Calculate date range (last 7 days)
    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - (7 * 24 * 60 * 60 * 1000));
    
    Logger.log('Report period: ' + startDate.toISOString() + ' to ' + endDate.toISOString());
    
    // Filter logs for sales in the last 7 days
    const salesLogs = logs.filter(log => {
      const logDate = new Date(log.timestamp);
      return logDate >= startDate && 
             logDate <= endDate && 
             (log.scanEvent === 'SALE_BOUTIQUE' || 
              log.scanEvent === 'SALE_MARCHE' || 
              log.scanEvent === 'SALE_SALEYA' ||
              log.scanEvent === 'DELIVERY_B2B');
    });
    
    Logger.log('Sales logs found: ' + salesLogs.length);
    
    // Aggregate sales by product and location
    const salesByProduct = {};
    
    salesLogs.forEach(log => {
      const product = products.find(p => log.serialNumber.startsWith(p.id));
      if (!product) return;
      
      if (!salesByProduct[product.id]) {
        salesByProduct[product.id] = {
          productName: product.name,
          boutique: 0,
          marche: 0,
          saleya: 0,
          b2b: 0,
          total: 0
        };
      }
      
      if (log.scanEvent === 'SALE_BOUTIQUE') salesByProduct[product.id].boutique++;
      else if (log.scanEvent === 'SALE_MARCHE') salesByProduct[product.id].marche++;
      else if (log.scanEvent === 'SALE_SALEYA') salesByProduct[product.id].saleya++;
      else if (log.scanEvent === 'DELIVERY_B2B') salesByProduct[product.id].b2b++;
      
      salesByProduct[product.id].total++;
    });
    
    // Write report
    const reportData = [
      ['Weekly Sales Report'],
      ['Period: ' + startDate.toLocaleDateString() + ' to ' + endDate.toLocaleDateString()],
      [],
      ['Product ID', 'Product Name', 'Boutique', 'Marche', 'Saleya', 'B2B', 'Total Sales']
    ];
    
    Object.keys(salesByProduct).forEach(productId => {
      const data = salesByProduct[productId];
      reportData.push([
        productId,
        data.productName,
        data.boutique,
        data.marche,
        data.saleya,
        data.b2b,
        data.total
      ]);
    });
    
    reportSheet.getRange(1, 1, reportData.length, 7).setValues(reportData);
    reportSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    reportSheet.getRange(4, 1, 1, 7).setFontWeight('bold').setBackground('#f3f3f3');
    reportSheet.autoResizeColumns(1, 7);
    
    Logger.log('=== Weekly Sales Report generated successfully ===');
    
    return { 
      message: 'Weekly sales report generated successfully',
      period: {
        start: startDate.toISOString(),
        end: endDate.toISOString()
      },
      totalSales: salesLogs.length,
      productsReported: Object.keys(salesByProduct).length
    };
    
  } catch (error) {
    Logger.log('ERROR in generateWeeklySalesReport: ' + error.message);
    Logger.log(error.stack);
    throw error;
  }
}
