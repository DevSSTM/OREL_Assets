// excelHandler.js - Excel template parsing and multi-tab workbook generation using SheetJS (XLSX)
// 4 Main Asset Categories: Machinery, Land and building, Tools and equipments, Furniture and fittings
import * as XLSX from 'xlsx';

// Canonical 23 Fields from Asset_Bulk_Upload_Template.xlsx
export const TEMPLATE_COLUMNS = [
  'Asset Name',
  'Asset Number',
  'Title',
  'Asset Type',
  'Category',
  'Brand',
  'Model',
  'Serial No',
  'Barcode',
  'QR Code',
  'Description',
  'Company',
  'Department',
  'Location',
  'Purchase Date',
  'Purchase Price',
  'Vendor',
  'Invoice No',
  'Warranty Expiry',
  'Warranty Provider',
  'Depreciation Method',
  'Useful Life (Months)',
  'Residual Value'
];

export const MAIN_4_CATEGORIES = [
  'Machinery',
  'Land and building',
  'Tools and equipments',
  'Furniture and fittings'
];

export const HEADER_TO_KEY_MAP = {
  'asset name': 'asset_name',
  'asset number': 'asset_number',
  'asset id': 'asset_number',
  'title': 'title',
  'asset type': 'asset_type',
  'category': 'category',
  'brand': 'brand',
  'model': 'model',
  'serial no': 'serial_no',
  'serial number': 'serial_no',
  'barcode': 'barcode',
  'qr code': 'qr_code',
  'description': 'description',
  'company': 'company',
  'department': 'department',
  'location': 'location',
  'building': 'building',
  'floor': 'floor',
  'purchase date': 'purchase_date',
  'purchase price': 'purchase_price',
  'cost': 'purchase_price',
  'purchase cost': 'purchase_price',
  'vendor': 'vendor',
  'invoice no': 'invoice_no',
  'invoice number': 'invoice_no',
  'warranty expiry': 'warranty_expiry',
  'warranty provider': 'warranty_provider',
  'depreciation method': 'depreciation_method',
  'useful life (months)': 'useful_life_months',
  'useful life': 'useful_life_months',
  'residual value': 'residual_value',
  'status': 'status'
};

/**
 * Resolves any raw category or asset type to one of the 4 Main Categories:
 * 1. Machinery
 * 2. Land and building
 * 3. Tools and equipments
 * 4. Furniture and fittings
 */
export function resolveMainCategory(rawCategory = '', rawAssetType = '') {
  const cat = String(rawCategory || '').toLowerCase().trim();
  const type = String(rawAssetType || '').toLowerCase().trim();

  // 1. Check Category match
  if (cat.includes('land') || cat.includes('build') || cat.includes('property') || cat.includes('real estate') || cat.includes('facility')) {
    return 'Land and building';
  }
  if (cat.includes('machin') || cat.includes('plant') || cat.includes('vehicle') || cat.includes('fleet') || cat.includes('generator') || cat.includes('motor')) {
    return 'Machinery';
  }
  if (cat.includes('tool') || cat.includes('equip') || cat.includes('it') || cat.includes('tech') || cat.includes('computer') || cat.includes('instrument')) {
    return 'Tools and equipments';
  }
  if (cat.includes('furnit') || cat.includes('fitting') || cat.includes('chair') || cat.includes('desk') || cat.includes('table') || cat.includes('fixture')) {
    return 'Furniture and fittings';
  }

  // 2. Fallback to Asset Type mapping
  if (type.includes('laptop') || type.includes('desktop') || type.includes('monitor') || type.includes('printer') || 
      type.includes('mobile') || type.includes('phone') || type.includes('tool') || type.includes('instrument') || 
      type.includes('gauge') || type.includes('server') || type.includes('network') || type.includes('device') || type.includes('kit')) {
    return 'Tools and equipments';
  }
  if (type.includes('vehicle') || type.includes('car') || type.includes('van') || type.includes('truck') || 
      type.includes('forklift') || type.includes('machin') || type.includes('cnc') || type.includes('milling') || 
      type.includes('generator') || type.includes('plant') || type.includes('compressor') || type.includes('pump')) {
    return 'Machinery';
  }
  if (type.includes('chair') || type.includes('desk') || type.includes('table') || type.includes('cabinet') || 
      type.includes('fixture') || type.includes('furnit') || type.includes('partition') || type.includes('lighting')) {
    return 'Furniture and fittings';
  }
  if (type.includes('build') || type.includes('land') || type.includes('office') || type.includes('warehouse') || type.includes('campus')) {
    return 'Land and building';
  }

  return 'Tools and equipments';
}

/**
 * Optimizes worksheet range bounds to skip trailing empty rows/cols in huge Excel files
 */
function getWorksheetSafeRange(worksheet) {
  let maxR = 0;
  let maxC = 0;
  let hasCells = false;

  for (const cellAddress in worksheet) {
    if (cellAddress[0] === '!') continue;
    try {
      const decoded = XLSX.utils.decode_cell(cellAddress);
      if (decoded.r > maxR) maxR = decoded.r;
      if (decoded.c > maxC) maxC = decoded.c;
      hasCells = true;
    } catch (e) {
      // ignore
    }
  }

  if (!hasCells) return null;
  return XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
}

export const ExcelHandler = {
  /**
   * Parse an uploaded Excel template file (.xlsx, .xls, .csv)
   * Automatically groups all asset rows under the 4 Main Categories:
   * Machinery, Land and building, Tools and equipments, Furniture and fittings
   */
  async parseTemplateFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });

          const sheetNames = workbook.SheetNames || [];
          const isStandardBulkTemplate = sheetNames.some(s => s.toLowerCase() === 'asset upload');
          const hasLookupValues = sheetNames.some(s => s.toLowerCase() === 'lookup values');

          // Extract Lookup Values if present
          const lookupValues = {};
          if (hasLookupValues) {
            const lookupSheet = workbook.Sheets['Lookup Values'] || workbook.Sheets['lookup values'];
            const safeRef = getWorksheetSafeRange(lookupSheet);
            if (safeRef) lookupSheet['!ref'] = safeRef;
            const lookupJson = XLSX.utils.sheet_to_json(lookupSheet, { header: 1 });
            for (let r = 1; r < lookupJson.length; r++) {
              const row = lookupJson[r];
              if (row && row[0] && row[1]) {
                const fieldName = String(row[0]).trim();
                const examples = String(row[1]).split(',').map(s => s.trim()).filter(Boolean);
                lookupValues[fieldName] = examples;
              }
            }
          }

          // Case 1: Standard Bulk Template (Sheet "Asset Upload")
          if (isStandardBulkTemplate) {
            const assetUploadSheet = workbook.Sheets['Asset Upload'] || workbook.Sheets['asset upload'];
            const safeRef = getWorksheetSafeRange(assetUploadSheet);
            if (safeRef) assetUploadSheet['!ref'] = safeRef;

            const jsonData = XLSX.utils.sheet_to_json(assetUploadSheet, { header: 1, defval: '' });
            const rawHeaders = (jsonData[0] || []).map(h => String(h).trim()).filter(Boolean);

            const allRows = [];
            for (let r = 1; r < jsonData.length; r++) {
              const row = jsonData[r];
              if (!row || !row.some(c => String(c).trim() !== '')) continue;

              const rowObj = {};
              rawHeaders.forEach((h, cIdx) => {
                const val = row[cIdx] !== undefined ? String(row[cIdx]).trim() : '';
                rowObj[h] = val;
                const normalizedKey = HEADER_TO_KEY_MAP[h.toLowerCase()];
                if (normalizedKey) {
                  rowObj[normalizedKey] = val;
                }
              });

              // Assign to one of the 4 Main Categories
              const mainCat = resolveMainCategory(rowObj.category || rowObj.Category, rowObj.asset_type || rowObj['Asset Type']);
              rowObj.category = mainCat;
              rowObj.categoryName = mainCat;
              rowObj.Category = mainCat;

              // Ensure essential identifiers exist
              rowObj.asset_name = rowObj.asset_name || rowObj['Asset Name'] || 'Unnamed Asset';
              rowObj.asset_number = rowObj.asset_number || rowObj['Asset Number'] || '';
              rowObj.asset_id = rowObj.asset_number || rowObj.id;
              rowObj.purchase_price = rowObj.purchase_price !== undefined ? rowObj.purchase_price : rowObj['Purchase Price'];
              rowObj.purchase_cost = rowObj.purchase_price;
              rowObj.status = rowObj.status || 'Active / In Use';

              allRows.push(rowObj);
            }

            // Group rows by the 4 Main Categories
            const groupedTabsMap = new Map();
            MAIN_4_CATEGORIES.forEach(mCat => groupedTabsMap.set(mCat, []));

            allRows.forEach(row => {
              const catName = row.category || 'Tools and equipments';
              if (!groupedTabsMap.has(catName)) {
                groupedTabsMap.set(catName, []);
              }
              groupedTabsMap.get(catName).push(row);
            });

            const detectedTabs = [];
            groupedTabsMap.forEach((rows, catName) => {
              detectedTabs.push({
                tabName: catName,
                headers: rawHeaders.length > 0 ? rawHeaders : TEMPLATE_COLUMNS,
                rows: rows
              });
            });

            resolve({
              fileName: file.name,
              isStandardBulkTemplate: true,
              totalTabs: detectedTabs.length,
              totalRows: allRows.length,
              lookupValues,
              tabs: detectedTabs,
              rawAssets: allRows
            });
            return;
          }

          // Case 2: Multi-Tab Workbook
          const detectedTabs = [];
          workbook.SheetNames.forEach((sheetName) => {
            if (sheetName.toLowerCase().includes('lookup')) return;

            const worksheet = workbook.Sheets[sheetName];
            const safeRef = getWorksheetSafeRange(worksheet);
            if (safeRef) worksheet['!ref'] = safeRef;

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            const rawHeaders = (jsonData[0] || []).map(h => String(h).trim()).filter(Boolean);
            const headers = rawHeaders.length > 0 ? rawHeaders : TEMPLATE_COLUMNS;

            const rows = [];
            for (let r = 1; r < jsonData.length; r++) {
              const row = jsonData[r];
              if (!row || !row.some(cell => String(cell).trim() !== '')) continue;

              const rowObj = {};
              headers.forEach((header, colIndex) => {
                const val = row[colIndex] !== undefined ? String(row[colIndex]).trim() : '';
                rowObj[header] = val;
                const normalizedKey = HEADER_TO_KEY_MAP[header.toLowerCase()];
                if (normalizedKey) {
                  rowObj[normalizedKey] = val;
                }
              });

              // Target category
              const mainCat = resolveMainCategory(sheetName, rowObj.asset_type || rowObj['Asset Type']);
              rowObj.category = mainCat;
              rowObj.categoryName = mainCat;
              rowObj.asset_name = rowObj.asset_name || rowObj['Asset Name'] || 'Asset';
              rowObj.asset_number = rowObj.asset_number || rowObj['Asset Number'] || '';
              rowObj.asset_id = rowObj.asset_number || rowObj.id;
              rowObj.purchase_price = rowObj.purchase_price !== undefined ? rowObj.purchase_price : rowObj['Purchase Price'];
              rowObj.purchase_cost = rowObj.purchase_price;
              rowObj.status = rowObj.status || 'Active / In Use';

              rows.push(rowObj);
            }

            detectedTabs.push({
              tabName: sheetName,
              headers: headers,
              rows: rows
            });
          });

          const totalRowsCount = detectedTabs.reduce((acc, t) => acc + t.rows.length, 0);

          resolve({
            fileName: file.name,
            isStandardBulkTemplate: false,
            totalTabs: detectedTabs.length,
            totalRows: totalRowsCount,
            lookupValues,
            tabs: detectedTabs
          });
        } catch (error) {
          reject(new Error(`Failed to parse Excel file: ${error.message}`));
        }
      };

      reader.onerror = (err) => {
        reject(new Error(`File read error: ${err}`));
      };

      reader.readAsArrayBuffer(file);
    });
  },

  /**
   * Generates the official OREL Asset Bulk Upload Template (.xlsx) matching
   * Asset_Bulk_Upload_Template.xlsx:
   * Sheet 1: Lookup Values (4 Main Categories, Asset Types, Depreciation Method, Company, Date Format)
   * Sheet 2: Asset Upload (23 standard columns with guidance rows or exported assets)
   */
  exportStandardTemplateWorkbook(assets = [], includeExistingAssets = false) {
    const workbook = XLSX.utils.book_new();

    // 1. Lookup Values Sheet
    const lookupRows = [
      ['Field', 'Example Values'],
      ['Main Categories', 'Machinery, Land and building, Tools and equipments, Furniture and fittings'],
      ['Asset Type', 'Laptop, Desktop, Monitor, Printer, Mobile, Vehicle, CNC / Milling Unit, Generator, Industrial Machinery, Power Tool, Testing Kit, Building, Furniture'],
      ['Depreciation Method', 'Straight Line, Reducing Balance, Units of Production'],
      ['Company', 'Orel Corp'],
      ['Date Format', 'YYYY-MM-DD']
    ];
    const lookupWorksheet = XLSX.utils.aoa_to_sheet(lookupRows);
    lookupWorksheet['!cols'] = [{ wch: 22 }, { wch: 85 }];
    XLSX.utils.book_append_sheet(workbook, lookupWorksheet, 'Lookup Values');

    // 2. Asset Upload Sheet
    let uploadData = [];

    if (includeExistingAssets && assets.length > 0) {
      uploadData = assets.map(a => ({
        'Asset Name': a.asset_name || a['Asset Name'] || '',
        'Asset Number': a.asset_number || a.asset_id || a['Asset Number'] || '',
        'Title': a.title || a.Title || '',
        'Asset Type': a.asset_type || a['Asset Type'] || '',
        'Category': a.category || a.categoryName || a.Category || 'Machinery',
        'Brand': a.brand || a.Brand || '',
        'Model': a.model || a.Model || '',
        'Serial No': a.serial_no || a['Serial No'] || a.serial_number || '',
        'Barcode': a.barcode || a.Barcode || '',
        'QR Code': a.qr_code || a['QR Code'] || '',
        'Description': a.description || a.Description || '',
        'Company': a.company || a.Company || 'Orel Corp',
        'Department': a.department || a.Department || '',
        'Location': a.location || a.Location || '',
        'Purchase Date': a.purchase_date || a['Purchase Date'] || '',
        'Purchase Price': a.purchase_price !== undefined ? a.purchase_price : (a.purchase_cost || ''),
        'Vendor': a.vendor || a.Vendor || '',
        'Invoice No': a.invoice_no || a['Invoice No'] || '',
        'Warranty Expiry': a.warranty_expiry || a['Warranty Expiry'] || '',
        'Warranty Provider': a.warranty_provider || a['Warranty Provider'] || '',
        'Depreciation Method': a.depreciation_method || a['Depreciation Method'] || 'Straight Line',
        'Useful Life (Months)': a.useful_life_months || a['Useful Life (Months)'] || 36,
        'Residual Value': a.residual_value !== undefined ? a.residual_value : (a['Residual Value'] || '')
      }));
    } else {
      // 4 Guide Rows (One for each of the 4 Main Categories)
      uploadData = [
        {
          'Asset Name': 'Haas 5-Axis Precision CNC Milling Center',
          'Asset Number': 'AST-MC-2001',
          'Title': 'Component Precision Machining Line 1',
          'Asset Type': 'CNC / Milling Unit',
          'Category': 'Machinery',
          'Brand': 'Haas Automation',
          'Model': 'VF-4SS',
          'Serial No': 'HAAS-2023-9941',
          'Barcode': '88492039001',
          'QR Code': 'QR-AST-MC-2001',
          'Description': '30 kW 12,000 RPM inline direct-drive spindle',
          'Company': 'Orel Corp',
          'Department': 'Operations',
          'Location': 'Main Factory, Machining Bay 1',
          'Purchase Date': '2023-01-20',
          'Purchase Price': 165000,
          'Vendor': 'Haas Asia Pacific Automation',
          'Invoice No': 'HAAS-INV-2023-014',
          'Warranty Expiry': '2025-01-20',
          'Warranty Provider': 'Haas Certified Industrial Warranty',
          'Depreciation Method': 'Units of Production',
          'Useful Life (Months)': 120,
          'Residual Value': 25000
        },
        {
          'Asset Name': 'Corporate Headquarters & Innovation Campus',
          'Asset Number': 'AST-LB-1001',
          'Title': 'Central Corporate Administration Campus',
          'Asset Type': 'Office Building',
          'Category': 'Land and building',
          'Brand': 'OREL Properties',
          'Model': '5-Story Commercial Campus',
          'Serial No': 'DEED-CA-88921-A',
          'Barcode': '88492069001',
          'QR Code': 'QR-AST-LB-1001',
          'Description': 'Executive corporate tower with solar rooftop array',
          'Company': 'Orel Corp',
          'Department': 'Executive',
          'Location': '450 Innovation Blvd, Silicon District, Colombo',
          'Purchase Date': '2021-03-15',
          'Purchase Price': 3450000,
          'Vendor': 'Prime Lands Residencies PLC',
          'Invoice No': 'INV-PLR-2021-08',
          'Warranty Expiry': '2041-03-15',
          'Warranty Provider': 'Architectural Structural Guarantee',
          'Depreciation Method': 'Straight Line',
          'Useful Life (Months)': 360,
          'Residual Value': 800000
        },
        {
          'Asset Name': 'Apple MacBook Pro 16" M2 Max',
          'Asset Number': 'AST-TE-3001',
          'Title': 'Lead Architect Dev Machine',
          'Asset Type': 'Laptop',
          'Category': 'Tools and equipments',
          'Brand': 'Apple',
          'Model': 'MacBook Pro 16" (Space Black)',
          'Serial No': 'C02G901XYZ8',
          'Barcode': '88492019001',
          'QR Code': 'QR-AST-TE-3001',
          'Description': 'Apple M2 Max chip, 32GB RAM, 1TB SSD',
          'Company': 'Orel Corp',
          'Department': 'IT & Software',
          'Location': 'Level 4, Tech Lab, Colombo HQ',
          'Purchase Date': '2023-04-15',
          'Purchase Price': 2850,
          'Vendor': 'Softlogic Information Systems',
          'Invoice No': 'INV-2023-4410',
          'Warranty Expiry': '2026-04-15',
          'Warranty Provider': 'AppleCare Enterprise',
          'Depreciation Method': 'Straight Line',
          'Useful Life (Months)': 36,
          'Residual Value': 450
        },
        {
          'Asset Name': 'Steelcase Executive Boardroom Conference Table (16-Seater)',
          'Asset Number': 'AST-FF-4001',
          'Title': 'Executive Boardroom Central Suite',
          'Asset Type': 'Conference Table',
          'Category': 'Furniture and fittings',
          'Brand': 'Steelcase',
          'Model': 'Potrero415 Architectural Suite',
          'Serial No': 'SC-CONF-8812',
          'Barcode': '88492049001',
          'QR Code': 'QR-AST-FF-4001',
          'Description': 'Solid walnut finish with integrated AV cable wells',
          'Company': 'Orel Corp',
          'Department': 'Executive',
          'Location': 'Corporate HQ, Level 5, Boardroom Alpha',
          'Purchase Date': '2022-11-15',
          'Purchase Price': 12500,
          'Vendor': 'Steelcase Architectural Furnishings',
          'Invoice No': 'SC-SL-2022-887',
          'Warranty Expiry': '2032-11-15',
          'Warranty Provider': 'Steelcase 10-Year Global Warranty',
          'Depreciation Method': 'Straight Line',
          'Useful Life (Months)': 120,
          'Residual Value': 1500
        }
      ];
    }

    // Templates contain headers only, so importing a blank template never creates examples.
    if (!includeExistingAssets || !assets.length) uploadData = [];
    const uploadWorksheet = XLSX.utils.json_to_sheet(uploadData, { header: TEMPLATE_COLUMNS });
    uploadWorksheet['!cols'] = TEMPLATE_COLUMNS.map(h => ({
      wch: Math.max(h.length + 4, 16)
    }));

    XLSX.utils.book_append_sheet(workbook, uploadWorksheet, 'Asset Upload');

    const fileName = includeExistingAssets
      ? `OREL_Assets_Export_${new Date().toISOString().split('T')[0]}.xlsx`
      : `Asset_Bulk_Upload_Template.xlsx`;

    XLSX.writeFile(workbook, fileName);
    return fileName;
  },

  /**
   * Generates a 4-tab Excel workbook matching the 4 main categories:
   * Machinery, Land and building, Tools and equipments, Furniture and fittings
   */
  exportTemplateWorkbook(categories, assets = [], includeExistingAssets = false) {
    const workbook = XLSX.utils.book_new();

    categories.forEach((cat) => {
      const tabName = cat.name.substring(0, 31);
      const headers = cat.fields.map(f => f.label);

      let sheetData = [];

      if (includeExistingAssets) {
        const catAssets = assets.filter(a => a.categoryId === cat.id || a.categoryName === cat.name);
        if (catAssets.length > 0) {
          sheetData = catAssets.map(a => {
            const row = {};
            cat.fields.forEach(f => {
              row[f.label] = a[f.id] !== undefined ? a[f.id] : (a[f.label] !== undefined ? a[f.label] : '');
            });
            return row;
          });
        } else {
          const emptyRow = {};
          headers.forEach(h => { emptyRow[h] = ''; });
          sheetData = [emptyRow];
        }
      } else {
        sheetData = [];
      }

      const worksheet = XLSX.utils.json_to_sheet(sheetData, { header: headers });
      worksheet['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 4, 16) }));
      XLSX.utils.book_append_sheet(workbook, worksheet, tabName);
    });

    const fileName = includeExistingAssets
      ? `OREL_Assets_4MainTabs_${new Date().toISOString().split('T')[0]}.xlsx`
      : `OREL_Asset_4MainCategories_Template.xlsx`;

    XLSX.writeFile(workbook, fileName);
    return fileName;
  }
};
