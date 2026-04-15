/**
 * ============================================================
 * 15B to BRZ Inventory — Google Apps Script Backend
 * ============================================================
 * 
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Open your Google Spreadsheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any existing code in Code.gs
 * 4. Paste this entire file
 * 5. Click Deploy → New Deployment
 * 6. Select "Web app" as the type
 * 7. Set "Execute as" → Me
 * 8. Set "Who has access" → Anyone
 * 9. Click Deploy and authorize
 * 10. Copy the Web App URL and paste it into app.js (SCRIPT_URL variable)
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SYSTEM_PIN = '123456'; // Default security PIN

// Sheet name mapping
const SHEETS = {
  'all_inventory': 'ALL INVENTORY',
  'figurine_list': 'FIGURINE LIST',
  'item_from_warehouse': 'ITEM FROM WAREHOUSE',
  'all_box': 'ALL BOX',
  'lipat_bahay': 'LIPAT BAHAY'
};

// Column headers for each sheet
const HEADERS = {
  'ALL INVENTORY': ['CODE', 'NAME OF ITEM', 'PLACE OF ITEM', 'DESCRIPTION', 'UPDATE'],
  'FIGURINE LIST': ['FIGURINE', 'FIGURINE PLACE', 'FIGURINE UPDATE', 'PICTURES', 'DATE DEPARTURE'],
  'ITEM FROM WAREHOUSE': ['# NUMBER OF ITEM', 'ITEM NAME', 'COLOR OF ITEM', 'PICTURES'],
  'ALL BOX': ['BOX NUMBER', 'BOX NAME', 'BOX PLACE', 'BOX DESCRIPTION', 'BOX UPDATE', 'PICTURES', 'DATE DEPARTURE', 'COMPANY'],
  'LIPAT BAHAY': ['ITEM/BOX NAME', 'ORIGIN ROOM', 'DESTINATION ROOM', 'STATUS', 'NOTES', 'PICTURES']
};

/**
 * Handle GET requests — fetch data from sheets
 */
function doGet(e) {
  try {
    const params = e.parameter;
    const action = params.action || 'fetch';
    const authPin = params.auth || '';

    // Verify PIN
    if (authPin !== SYSTEM_PIN) {
      return jsonResponse({ error: 'Unauthorized. Invalid PIN.' }, 401);
    }

    if (action === 'auth_check') {
      return jsonResponse({ success: true, message: 'Authenticated' });
    }

    const sheetKey = params.sheet || 'all_inventory';
    const sheetName = SHEETS[sheetKey];

    if (!sheetName) {
      return jsonResponse({ error: 'Invalid sheet name' }, 400);
    }

    if (action === 'fetch') {
      return fetchData(sheetName);
    } else if (action === 'stats') {
      return fetchStats();
    }

    return jsonResponse({ error: 'Invalid action' }, 400);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

/**
 * Handle POST requests — add, update, delete data
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const authPin = body.auth || '';

    // Verify PIN
    if (authPin !== SYSTEM_PIN) {
      return jsonResponse({ error: 'Unauthorized. Invalid PIN.' }, 401);
    }

    // Handle image upload (doesn't need sheet validation)
    if (action === 'upload_image') {
      return uploadImage(body.imageData, body.fileName);
    }

    const sheetKey = body.sheet;
    const sheetName = SHEETS[sheetKey];

    if (!sheetName) {
      return jsonResponse({ error: 'Invalid sheet name' }, 400);
    }

    switch (action) {
      case 'add':
        return addRow(sheetName, body.data);
      case 'update':
        return updateRow(sheetName, body.row, body.data);
      case 'delete':
        return deleteRow(sheetName, body.row);
      case 'init_database':
        return initDatabase();
      case 'create_sheet':
        return createCustomSheet(body.sheetName, body.headers);
      default:
        return jsonResponse({ error: 'Invalid action' }, 400);
    }
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

/**
 * Initialize all standard sheets if they don't exist
 */
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];

  Object.keys(SHEETS).forEach(key => {
    const name = SHEETS[key];
    let sheet = ss.getSheetByName(name);
    let created = false;

    if (!sheet) {
      sheet = ss.insertSheet(name);
      created = true;
      
      // Add headers
      const headers = HEADERS[name] || [];
      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // Format headers: Bold, nice background, frozen top row
        sheet.getRange(1, 1, 1, headers.length)
             .setFontWeight('bold')
             .setBackground('#f3f4f6')
             .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
        sheet.setFrozenRows(1);
      }
    }
    results.push({ name: name, status: created ? 'Created' : 'Already Exists' });
  });

  return jsonResponse({ success: true, results: results });
}

/**
 * Create a custom sheet with user-defined headers
 */
function createCustomSheet(name, headerList) {
  if (!name) return jsonResponse({ error: 'Sheet name is required' }, 400);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(name)) {
    return jsonResponse({ error: 'A sheet named "' + name + '" already exists' }, 400);
  }

  const sheet = ss.insertSheet(name);
  const headers = headerList || [];
  
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#f3f4f6')
         .setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);
  }

  return jsonResponse({ success: true, message: 'Custom sheet "' + name + '" created successfully' });
}

/**
 * Upload an image to Google Drive and return the viewable URL
 */
function uploadImage(base64Data, fileName) {
  if (!base64Data) {
    return jsonResponse({ error: 'No image data provided' }, 400);
  }

  const folderName = '15B-BRZ-Inventory-Images';
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  
  // Determine content type from base64 header or default to png
  let contentType = 'image/png';
  let cleanData = base64Data;
  
  if (base64Data.includes(',')) {
    const parts = base64Data.split(',');
    const header = parts[0];
    cleanData = parts[1];
    
    if (header.includes('image/jpeg')) contentType = 'image/jpeg';
    else if (header.includes('image/png')) contentType = 'image/png';
    else if (header.includes('image/webp')) contentType = 'image/webp';
    else if (header.includes('image/gif')) contentType = 'image/gif';
  }
  
  const decoded = Utilities.base64Decode(cleanData);
  const blob = Utilities.newBlob(decoded, contentType, fileName || 'image_' + Date.now());
  const file = folder.createFile(blob);
  
  // Make publicly viewable
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  const fileId = file.getId();
  const viewUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
  
  return jsonResponse({
    success: true,
    url: viewUrl,
    fileId: fileId
  });
}

/**
 * Fetch all data from a specific sheet
 */
function fetchData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse({ error: 'Sheet not found: ' + sheetName }, 404);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return jsonResponse({ headers: HEADERS[sheetName] || [], data: [], sheetName: sheetName });
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = dataRange.getValues();

  // Convert to array of objects
  const rows = data.map((row, index) => {
    const obj = { _rowIndex: index + 2 }; // Store actual row number (1-indexed, skip header)
    headers.forEach((header, colIndex) => {
      obj[header] = row[colIndex] !== undefined ? row[colIndex].toString() : '';
    });
    return obj;
  });

  return jsonResponse({
    headers: headers,
    data: rows,
    sheetName: sheetName,
    totalRows: rows.length
  });
}

/**
 * Fetch stats for all sheets
 */
function fetchStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stats = {};

  Object.keys(SHEETS).forEach(key => {
    const sheetName = SHEETS[key];
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      stats[key] = { 
        count: Math.max(0, lastRow - 1),
        sheetName: sheetName 
      };
    }
  });

  return jsonResponse({ stats: stats });
}

/**
 * Add a new row to a sheet
 */
function addRow(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse({ error: 'Sheet not found' }, 404);
  }

  const headers = HEADERS[sheetName];
  if (!headers) {
    return jsonResponse({ error: 'No headers defined for sheet' }, 400);
  }

  // Build row array matching header order
  const rowData = headers.map(header => data[header] || '');
  sheet.appendRow(rowData);

  return jsonResponse({ 
    success: true, 
    message: 'Row added successfully',
    row: sheet.getLastRow()
  });
}

/**
 * Update an existing row in a sheet
 */
function updateRow(sheetName, rowIndex, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse({ error: 'Sheet not found' }, 404);
  }

  const headers = HEADERS[sheetName];
  const rowData = headers.map(header => data[header] !== undefined ? data[header] : '');

  const range = sheet.getRange(rowIndex, 1, 1, headers.length);
  range.setValues([rowData]);

  return jsonResponse({ 
    success: true, 
    message: 'Row updated successfully' 
  });
}

/**
 * Delete a row from a sheet
 */
function deleteRow(sheetName, rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse({ error: 'Sheet not found' }, 404);
  }

  sheet.deleteRow(rowIndex);

  return jsonResponse({ 
    success: true, 
    message: 'Row deleted successfully' 
  });
}

/**
 * Create a JSON response with CORS headers
 */
function jsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
