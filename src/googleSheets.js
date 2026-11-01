// googleSheets.js - Direct integration with Google Sheets backend via Google Apps Script Web App

export const GOOGLE_APPS_SCRIPT_CODE = `/**
 * ==============================================================================
 * OREL ASSET MANAGEMENT - GOOGLE APPS SCRIPT BACKEND
 * ==============================================================================
 * This script runs inside your Google Spreadsheet and serves as the backend database.
 * Every asset category (e.g., "Land and building", "Machinery", "Tools", "Furniture and Fittings", "AC", "Fans")
 * is stored in its OWN SEPARATE TAB inside this SAME spreadsheet.
 * 
 * HOW TO DEPLOY:
 * 1. In your Google Sheet, click: Extensions > Apps Script
 * 2. Delete any existing code, and paste this entire file.
 * 3. Click the blue "Deploy" button at top right > "New deployment"
 * 4. Select type: "Web app" (click gear icon next to 'Select type')
 * 5. Description: "Asset Management Backend"
 * 6. Execute as: "Me (your email)"
 * 7. Who has access: "Anyone"  <-- CRITICAL for web app access!
 * 8. Click "Deploy", authorize permissions when prompted, and copy the Web App URL.
 * 9. Paste the Web App URL into your Asset Management Web App Settings!
 * ==============================================================================
 */

var SPREADSHEET_ID = '1qYuBmn_7H59AorWcWQraP6mNtgupw6-vkAPH8qPZVjg';

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (action === 'ping') {
    return createJsonResponse({
      status: 'success',
      message: 'Successfully connected to Google Sheet: ' + ss.getName(),
      sheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      tabsCount: ss.getSheets().length
    });
  }

  if (action === 'getData') {
    var sheets = ss.getSheets();
    var tabsData = [];
    var allAssets = [];

    sheets.forEach(function(sh) {
      var tabName = sh.getName();
      // Skip hidden sheets or temp sheets if any
      if (tabName.indexOf('_') === 0) return;

      var data = sh.getDataRange().getValues();
      if (!data || data.length === 0) return;

      var headers = data[0];
      var tabAssets = [];

      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        if (!row[0] && !row[1]) continue; // Skip completely empty row

        var assetObj = { categoryName: tabName };
        for (var c = 0; c < headers.length; c++) {
          var headerKey = headers[c];
          if (headerKey) {
            assetObj[headerKey] = row[c];
          }
        }
        tabAssets.push(assetObj);
        allAssets.push(assetObj);
      }

      tabsData.push({
        tabName: tabName,
        headers: headers,
        rowCount: tabAssets.length
      });
    });

    return createJsonResponse({
      status: 'success',
      tabs: tabsData,
      assets: allAssets,
      timestamp: new Date().toISOString()
    });
  }

  return createJsonResponse({ status: 'error', message: 'Unknown GET action: ' + action });
}

function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : '';
    if (!contents) {
      return createJsonResponse({ status: 'error', message: 'No POST data received' });
    }

    var payload = JSON.parse(contents);
    var action = payload.action;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. FULL SYNC (Push all categories & assets to their respective tabs)
    if (action === 'syncAll') {
      var tabs = payload.tabs || [];
      var results = [];

      tabs.forEach(function(t) {
        var tabName = t.tabName;
        var headers = t.headers || [];
        var rows = t.rows || [];

        var sh = ss.getSheetByName(tabName);
        if (!sh) {
          sh = ss.insertSheet(tabName);
        }

        sh.clearContents();

        // Write header row
        if (headers.length > 0) {
          sh.appendRow(headers);
          // Style header row nicely
          var headerRange = sh.getRange(1, 1, 1, headers.length);
          headerRange.setFontWeight('bold');
          headerRange.setBackground('#0f172a');
          headerRange.setFontColor('#ffffff');
        }

        // Write data rows
        if (rows.length > 0) {
          var matrix = [];
          rows.forEach(function(r) {
            var rowArray = [];
            headers.forEach(function(h) {
              rowArray.push(r[h] !== undefined && r[h] !== null ? r[h] : '');
            });
            matrix.push(rowArray);
          });
          if (matrix.length > 0) {
            sh.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
          }
        }

        // Auto-fit column widths
        for (var col = 1; col <= Math.min(headers.length, 15); col++) {
          sh.autoResizeColumn(col);
        }

        results.push({ tabName: tabName, rowsSaved: rows.length });
      });

      return createJsonResponse({
        status: 'success',
        message: 'Synchronized ' + results.length + ' category tabs in sheet',
        details: results
      });
    }

    // 2. SAVE SINGLE ASSET (Create or update row in its category tab)
    if (action === 'saveAsset') {
      var tabName = payload.tabName;
      var asset = payload.asset;
      var headers = payload.headers || [];

      var sh = ss.getSheetByName(tabName);
      if (!sh) {
        sh = ss.insertSheet(tabName);
        if (headers.length > 0) {
          sh.appendRow(headers);
          sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
        }
      }

      // Read current tab headers and append any newly added custom fields
      var curLastCol = Math.max(1, sh.getLastColumn());
      var curHeaders = sh.getRange(1, 1, 1, curLastCol).getValues()[0];
      
      headers.forEach(function(h) {
        if (curHeaders.indexOf(h) === -1 && h.trim() !== '') {
          var nextCol = curHeaders.length + 1;
          sh.getRange(1, nextCol).setValue(h).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
          curHeaders.push(h);
        }
      });

      // Search if Asset ID already exists in this tab
      var assetId = asset['Asset ID'] || asset['asset_id'] || asset['id'];
      var data = sh.getDataRange().getValues();
      var targetRow = -1;

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(assetId).trim()) {
          targetRow = i + 1;
          break;
        }
      }

      var rowValues = curHeaders.map(function(h) {
        return asset[h] !== undefined && asset[h] !== null ? asset[h] : '';
      });

      if (targetRow > 0) {
        sh.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sh.appendRow(rowValues);
      }

      return createJsonResponse({
        status: 'success',
        message: 'Saved asset ' + assetId + ' to tab "' + tabName + '"',
        row: targetRow > 0 ? targetRow : sh.getLastRow()
      });
    }

    // 3. DELETE ASSET
    if (action === 'deleteAsset') {
      var tabName = payload.tabName;
      var assetId = payload.assetId;
      var sh = ss.getSheetByName(tabName);
      if (!sh) {
        return createJsonResponse({ status: 'error', message: 'Tab not found: ' + tabName });
      }

      var data = sh.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][0]).trim() === String(assetId).trim()) {
          sh.deleteRow(r + 1);
          return createJsonResponse({ status: 'success', message: 'Deleted asset ' + assetId + ' from ' + tabName });
        }
      }

      return createJsonResponse({ status: 'error', message: 'Asset ID not found: ' + assetId });
    }

    return createJsonResponse({ status: 'error', message: 'Invalid action: ' + action });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
`;

export const GoogleSheets = {
  // Test connection to Google Apps Script Web App
  async testConnection(scriptUrl) {
    if (!scriptUrl || !scriptUrl.trim()) {
      throw new Error('Please enter your Google Apps Script Web App URL');
    }

    const cleanUrl = scriptUrl.trim();
    const testUrl = `${cleanUrl}${cleanUrl.includes('?') ? '&' : '?'}action=ping`;

    try {
      const response = await fetch(testUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP Error ${response.status}: Failed to reach Google Apps Script`);
      }

      const result = await response.json();
      if (result.status === 'success') {
        return result;
      } else {
        throw new Error(result.message || 'Verification failed');
      }
    } catch (err) {
      console.error('Google Sheets connection test failed:', err);
      // Give clear troubleshooting tips
      throw new Error(
        `Connection failed: ${err.message}. Make sure your Apps Script is deployed as Web App with "Who has access: Anyone".`
      );
    }
  },

  // Pull all data from all tabs inside the Google Sheet
  async fetchRemoteData(scriptUrl) {
    const cleanUrl = scriptUrl.trim();
    const fetchUrl = `${cleanUrl}${cleanUrl.includes('?') ? '&' : '?'}action=getData`;

    const response = await fetch(fetchUrl, {
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch Google Sheet data (Status ${response.status})`);
    }

    const result = await response.json();
    return result;
  },

  // Push all categories and their assets to the Google Sheet (creates 1 tab per category)
  async pushAllToGoogleSheet(scriptUrl, categories, assets) {
    const cleanUrl = scriptUrl.trim();

    // Group assets by category
    const tabsPayload = categories.map(cat => {
      const tabName = cat.name.trim();
      const catAssets = assets.filter(a => a.categoryId === cat.id || a.categoryName === cat.name);
      
      // Determine headers from fields
      const headers = cat.fields.map(f => f.label);

      // Map rows
      const rows = catAssets.map(a => {
        const rowObj = {};
        cat.fields.forEach(f => {
          let val = a[f.id] !== undefined ? a[f.id] : (a[f.label] !== undefined ? a[f.label] : '');
          if ((val === '' || val === undefined) && f.id === 'asset_number' && a.asset_id) val = a.asset_id;
          if ((val === '' || val === undefined) && f.id === 'purchase_price' && a.purchase_cost) val = a.purchase_cost;
          rowObj[f.label] = val;
        });
        return rowObj;
      });

      return {
        tabName: tabName,
        headers: headers,
        rows: rows
      };
    });

    const payload = {
      action: 'syncAll',
      tabs: tabsPayload
    };

    // Google Apps Script requires text/plain for CORS preflight avoidance or standard POST
    const response = await fetch(cleanUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    const result = await response.json();
    return result;
  },

  // Save or update an individual asset to its category tab
  async saveSingleAsset(scriptUrl, category, asset) {
    if (!scriptUrl) return null;

    const tabName = category.name.trim();
    const headers = category.fields.map(f => f.label);

    const assetRow = {};
    category.fields.forEach(f => {
      let val = asset[f.id] !== undefined ? asset[f.id] : (asset[f.label] !== undefined ? asset[f.label] : '');
      if ((val === '' || val === undefined) && f.id === 'asset_number' && asset.asset_id) val = asset.asset_id;
      if ((val === '' || val === undefined) && f.id === 'purchase_price' && asset.purchase_cost) val = asset.purchase_cost;
      assetRow[f.label] = val;
    });

    const payload = {
      action: 'saveAsset',
      tabName: tabName,
      headers: headers,
      asset: assetRow
    };

    const response = await fetch(scriptUrl.trim(), {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    return await response.json();
  },

  // Delete an individual asset from its category tab
  async deleteSingleAsset(scriptUrl, categoryName, assetId) {
    if (!scriptUrl) return null;

    const payload = {
      action: 'deleteAsset',
      tabName: categoryName.trim(),
      assetId: assetId
    };

    const response = await fetch(scriptUrl.trim(), {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    return await response.json();
  }
};
