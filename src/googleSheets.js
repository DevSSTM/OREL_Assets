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
      apiVersion: 2,
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
      if (headers.indexOf('Asset Number') < 0) return;
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
        assetObj.categoryName = assetObj['Category'] || tabName;
        assetObj.photos = ['Photo 1', 'Photo 2'].map(function(key) {
          var id = photoFileId(assetObj[key]);
          if (!id) return '';
          try {
            var blob = DriveApp.getFileById(id).getBlob();
            return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
          } catch (error) { return ''; }
        });
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var contents = e.postData ? e.postData.contents : '';
    if (!contents) {
      return createJsonResponse({ status: 'error', message: 'No POST data received' });
    }

    var payload = JSON.parse(contents);
    var action = payload.action;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // All writes locate the existing record by its original number, across legacy tabs.
    if (action === 'syncAll') {
      return createJsonResponse({ status: 'error', message: 'Bulk overwrite is disabled. Save individual assets instead.' });
    }
    if (action === 'saveAsset' || action === 'deleteAsset') {
      var originalId = String(payload.originalId || payload.assetId || (payload.asset && payload.asset['Asset Number']) || '').trim();
      if (!originalId) throw new Error('Asset Number is required.');
      var matches = [];
      ss.getSheets().forEach(function(sheet) {
        var data = sheet.getDataRange().getValues();
        var column = data[0].indexOf('Asset Number');
        if (column < 0) return;
        for (var r = 1; r < data.length; r++) {
          if (String(data[r][column]).trim() === originalId) matches.push({ sheet: sheet, row: r + 1 });
        }
      });
      if (matches.length > 1) throw new Error('Duplicate asset numbers found in the spreadsheet. Resolve them before editing or deleting.');
      if (action === 'deleteAsset') {
        if (!matches.length) throw new Error('Asset not found in the spreadsheet.');
        var deleteHeaders = matches[0].sheet.getRange(1, 1, 1, matches[0].sheet.getLastColumn()).getValues()[0];
        var revisionColumn = deleteHeaders.indexOf('Revision');
        var currentRevision = revisionColumn < 0 ? '' : matches[0].sheet.getRange(matches[0].row, revisionColumn + 1).getValue();
        if (String(currentRevision || '') !== String(payload.expectedRevision || '')) throw new Error('This asset was changed by another user. Pull the latest data before deleting.');
        matches[0].sheet.deleteRow(matches[0].row);
        return createJsonResponse({ status: 'success', message: 'Asset deleted.' });
      }
      if (payload.isEdit && !matches.length) throw new Error('The original asset was not found. Refresh the data before editing.');
      if (!payload.isEdit && matches.length) throw new Error('An asset with this number already exists.');
      var newId = String(payload.asset['Asset Number'] || '').trim();
      if (!newId) throw new Error('Asset Number is required.');
      if (newId !== originalId) {
        var collision = false;
        ss.getSheets().forEach(function(sheet) {
          var data = sheet.getDataRange().getValues();
          var column = data[0].indexOf('Asset Number');
          if (column >= 0 && data.slice(1).some(function(row) { return String(row[column]).trim() === newId; })) collision = true;
        });
        if (collision) throw new Error('The new asset number already exists.');
      }
      var sheet = matches.length ? matches[0].sheet : ss.getSheetByName('Asset Upload');
      if (!sheet) sheet = ss.insertSheet('Asset Upload');
      var headers = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
      (payload.headers || []).concat(['Photo 1', 'Photo 2', 'Revision']).forEach(function(header) {
        if (header && headers.indexOf(header) < 0) { headers.push(header); sheet.getRange(1, headers.length).setValue(header); }
      });
      var values = matches.length ? sheet.getRange(matches[0].row, 1, 1, headers.length).getValues()[0] : headers.map(function() { return ''; });
      if (matches.length && String(values[headers.indexOf('Revision')] || '') !== String(payload.expectedRevision || '')) {
        throw new Error('This asset was changed by another user. Pull the latest data before editing.');
      }
      Object.keys(payload.asset).forEach(function(key) {
        var column = headers.indexOf(key);
        if (column >= 0) values[column] = payload.asset[key];
      });
      var photos = payload.photos || ['', ''];
      for (var p = 0; p < 2; p++) {
        var photoColumn = headers.indexOf('Photo ' + (p + 1));
        var photo = photos[p] || '';
        if (!photo) { values[photoColumn] = ''; continue; }
        var match = /^data:image\\/(jpeg|png|webp);base64,(.+)$/.exec(photo);
        if (!match) throw new Error('Invalid photo format.');
        var bytes = Utilities.base64Decode(match[2]);
        var existingId = photoFileId(values[photoColumn]);
        var same = false;
        if (existingId) {
          try { same = Utilities.base64Encode(DriveApp.getFileById(existingId).getBlob().getBytes()) === match[2]; } catch (ignored) {}
        }
        if (!same) {
          var folders = DriveApp.getFoldersByName('OREL Asset Photos');
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('OREL Asset Photos');
          var fileId = folder.createFile(Utilities.newBlob(bytes, 'image/' + match[1], newId + '-photo-' + (p + 1))).getId();
          values[photoColumn] = 'https://drive.google.com/open?id=' + fileId;
        }
      }
      var target = matches.length ? matches[0].row : Math.max(2, sheet.getLastRow() + 1);
      var revision = Utilities.getUuid();
      values[headers.indexOf('Revision')] = revision;
      sheet.getRange(target, 1, 1, headers.length).setValues([values]);
      return createJsonResponse({ status: 'success', message: 'Asset saved.', revision: revision });
    }

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
      var assetId = asset['Asset Number'] || asset['Asset ID'] || asset['asset_id'] || asset['id'];
      var data = sh.getDataRange().getValues();
      var targetRow = -1;
      var idColumn = curHeaders.indexOf('Asset Number');
      if (idColumn < 0) idColumn = curHeaders.indexOf('Asset ID');
      if (!assetId || idColumn < 0) {
        return createJsonResponse({ status: 'error', message: 'Asset Number is required.' });
      }

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idColumn]).trim() === String(assetId).trim()) {
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
      var idColumn = data[0].indexOf('Asset Number');
      if (idColumn < 0) idColumn = data[0].indexOf('Asset ID');
      if (idColumn < 0) return createJsonResponse({ status: 'error', message: 'Asset Number column not found.' });
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][idColumn]).trim() === String(assetId).trim()) {
          sh.deleteRow(r + 1);
          return createJsonResponse({ status: 'success', message: 'Deleted asset ' + assetId + ' from ' + tabName });
        }
      }

      return createJsonResponse({ status: 'error', message: 'Asset ID not found: ' + assetId });
    }

    return createJsonResponse({ status: 'error', message: 'Invalid action: ' + action });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function photoFileId(value) {
  var text = String(value || '');
  var match = /[?&]id=([^&]+)/.exec(text);
  return match ? match[1] : text;
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
    if (!response.ok || result.status !== 'success') throw new Error(result.message || 'Failed to fetch assets.');
    return result;
  },

  // Push all categories and their assets to the Google Sheet (creates 1 tab per category)
  async pushAllToGoogleSheet(scriptUrl, categories, assets) {
    throw new Error('Bulk overwrite is disabled. Save assets individually instead.');
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
  async saveSingleAsset(scriptUrl, category, asset, originalId = '') {
    if (!scriptUrl) return null;
    const connection = await this.testConnection(scriptUrl);
    if (connection.apiVersion !== 2) throw new Error('Update and redeploy the Apps Script shown in connection settings before saving.');

    const tabName = category.name.trim();
    const headers = category.fields.map(f => f.label);

    const assetRow = {};
    category.fields.forEach(f => {
      let val = asset[f.id] !== undefined ? asset[f.id] : (asset[f.label] !== undefined ? asset[f.label] : '');
      if ((val === '' || val === undefined) && f.id === 'asset_number' && asset.asset_id) val = asset.asset_id;
      if ((val === '' || val === undefined) && f.id === 'purchase_price' && asset.purchase_cost) val = asset.purchase_cost;
      assetRow[f.label] = val;
    });
    assetRow.Category = category.name;

    const payload = {
      action: 'saveAsset',
      tabName: tabName,
      headers: headers,
      asset: assetRow,
      originalId: originalId || asset.asset_number,
      isEdit: Boolean(originalId),
      expectedRevision: asset._revision || '',
      photos: asset.photos || ['', '']
    };

    const response = await fetch(scriptUrl.trim(), {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    const result = await response.json();
    if (!response.ok || result.status !== 'success') throw new Error(result.message || 'Failed to save asset to Google Sheets.');
    return result;
  },

  // Delete an individual asset from its category tab
  async deleteSingleAsset(scriptUrl, categoryName, assetId, revision = '') {
    if (!scriptUrl) return null;
    const connection = await this.testConnection(scriptUrl);
    if (connection.apiVersion !== 2) throw new Error('Update and redeploy the Apps Script shown in connection settings before deleting.');

    const payload = {
      action: 'deleteAsset',
      tabName: categoryName.trim(),
      assetId: assetId,
      expectedRevision: revision
    };

    const response = await fetch(scriptUrl.trim(), {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    const result = await response.json();
    if (!response.ok || result.status !== 'success') throw new Error(result.message || 'Failed to delete asset from Google Sheets.');
    return result;
  }
};
