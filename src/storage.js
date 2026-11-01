// storage.js - Local state and persistence management for Asset Management App
// 4 Main Asset Categories: Machinery, Land and building, Tools and equipments, Furniture and fittings
// Aligned with OREL Asset Bulk Upload Template (23 Fields)

const STORAGE_KEYS = {
  CATEGORIES: 'orel_asset_categories_v4',
  ASSETS: 'orel_asset_items_v4',
  SETTINGS: 'orel_asset_settings_v4'
};

export const MAIN_4_CATEGORIES = [
  'Machinery',
  'Land and building',
  'Tools and equipments',
  'Furniture and fittings'
];

/**
 * Standard 23 Template Fields matching Asset_Bulk_Upload_Template.xlsx exactly
 * plus lifecycle 'Status'
 */
export const STANDARD_TEMPLATE_FIELDS = [
  // 1. Identification
  { id: 'asset_name', label: 'Asset Name', type: 'text', required: true, isDefault: true, group: 'Identification', placeholder: 'e.g. Haas 5-Axis Precision CNC Milling Center' },
  { id: 'asset_number', label: 'Asset Number', type: 'text', required: true, isDefault: true, group: 'Identification', placeholder: 'e.g. AST-MC-2001' },
  { id: 'title', label: 'Title', type: 'text', required: false, isDefault: true, group: 'Identification', placeholder: 'e.g. Component Precision Machining Line 1' },
  { 
    id: 'asset_type', 
    label: 'Asset Type', 
    type: 'select', 
    options: [
      'Industrial Machinery',
      'CNC / Milling Unit',
      'Production Line',
      'Generator / Power Unit',
      'Commercial Vehicle / Van',
      'Forklift / Heavy Equipment',
      'Office Building',
      'Warehouse / Logistics Hub',
      'Plant Facility',
      'Land Parcel',
      'Laptop',
      'Desktop',
      'Monitor',
      'Printer',
      'Mobile Phone',
      'Server / Network Equipment',
      'Power Tool',
      'Hand Tool',
      'Testing & Calibration Kit',
      'Workstation / Desk',
      'Ergonomic Chair',
      'Conference Table',
      'Filing / Storage Cabinet',
      'HVAC / Electrical Fixture',
      'Other'
    ], 
    required: true, 
    isDefault: true, 
    group: 'Identification' 
  },
  { 
    id: 'category', 
    label: 'Category', 
    type: 'select', 
    options: MAIN_4_CATEGORIES, 
    required: true, 
    isDefault: true, 
    group: 'Identification' 
  },
  { 
    id: 'status', 
    label: 'Status', 
    type: 'select', 
    options: ['Active / In Use', 'Available / In Stock', 'In Stock', 'Not Currently in Use', 'Under Maintenance', 'In Storage', 'Disposed / Written Off'], 
    required: true, 
    isDefault: true, 
    group: 'Identification' 
  },

  // 2. Specifications & Hardware
  { id: 'brand', label: 'Brand', type: 'text', required: false, isDefault: true, group: 'Specifications', placeholder: 'e.g. Haas, Apple, Toyota, Steelcase, Norbar' },
  { id: 'model', label: 'Model', type: 'text', required: false, isDefault: true, group: 'Specifications', placeholder: 'e.g. VF-4SS, MacBook Pro 16", HiAce 2.8L' },
  { id: 'serial_no', label: 'Serial No', type: 'text', required: false, isDefault: true, group: 'Specifications', placeholder: 'e.g. HAAS-2023-9941, SN-884920' },
  { id: 'barcode', label: 'Barcode', type: 'text', required: false, isDefault: true, group: 'Specifications', placeholder: 'e.g. 88492039001' },
  { id: 'qr_code', label: 'QR Code', type: 'text', required: false, isDefault: true, group: 'Specifications', placeholder: 'e.g. QR-AST-MC-2001' },
  { id: 'description', label: 'Description', type: 'textarea', required: false, isDefault: true, group: 'Specifications', placeholder: 'Technical specifications, rating, dimensions, accessories, condition...' },

  // 3. Organization & Location
  { id: 'company', label: 'Company', type: 'select', options: ['Orel Corp', 'Orel IT', 'Orel Manufacturing', 'Orel Energy', 'Orel Logistics'], required: true, isDefault: true, group: 'Organization' },
  { id: 'department', label: 'Department', type: 'select', options: ['Operations', 'Engineering', 'IT & Software', 'Logistics & Transport', 'Finance & Accounting', 'Facilities & Maintenance', 'Human Resources', 'Executive'], required: false, isDefault: true, group: 'Organization' },
  { id: 'location', label: 'Location', type: 'text', required: true, isDefault: true, group: 'Organization', placeholder: 'e.g. Main Plant Bay 1, Level 4 Tech Lab, Central Yard' },

  // 4. Procurement & Financials
  { id: 'purchase_date', label: 'Purchase Date', type: 'date', required: false, isDefault: true, group: 'Financials' },
  { id: 'purchase_price', label: 'Purchase Price', type: 'currency', required: true, isDefault: true, group: 'Financials', placeholder: '0.00' },
  { id: 'vendor', label: 'Vendor', type: 'text', required: false, isDefault: true, group: 'Financials', placeholder: 'e.g. Haas Automation Asia, Softlogic IT, Toyota Lanka' },
  { id: 'invoice_no', label: 'Invoice No', type: 'text', required: false, isDefault: true, group: 'Financials', placeholder: 'e.g. INV-2023-8841' },

  // 5. Warranty & Depreciation
  { id: 'warranty_expiry', label: 'Warranty Expiry', type: 'date', required: false, isDefault: true, group: 'Warranty & Depreciation' },
  { id: 'warranty_provider', label: 'Warranty Provider', type: 'text', required: false, isDefault: true, group: 'Warranty & Depreciation', placeholder: 'e.g. Haas Certified Support, AppleCare, Toyota Care' },
  { id: 'depreciation_method', label: 'Depreciation Method', type: 'select', options: ['Straight Line', 'Reducing Balance', 'Units of Production'], required: false, isDefault: true, group: 'Warranty & Depreciation' },
  { id: 'useful_life_months', label: 'Useful Life (Months)', type: 'number', required: false, isDefault: true, group: 'Warranty & Depreciation', placeholder: '36' },
  { id: 'residual_value', label: 'Residual Value', type: 'currency', required: false, isDefault: true, group: 'Warranty & Depreciation', placeholder: '0.00' }
];

export const DEFAULT_CATEGORIES = [
  {
    id: 'machinery',
    name: 'Machinery',
    tabName: 'Machinery',
    description: 'Industrial machinery, CNC milling centers, cleanroom lines, generators & commercial fleet',
    icon: 'Cog',
    color: '#3b82f6', // Blue
    fields: JSON.parse(JSON.stringify(STANDARD_TEMPLATE_FIELDS))
  },
  {
    id: 'land-and-building',
    name: 'Land and building',
    tabName: 'Land and building',
    description: 'Real estate, land parcels, corporate facilities, warehouses & infrastructure structures',
    icon: 'Building2',
    color: '#10b981', // Emerald
    fields: JSON.parse(JSON.stringify(STANDARD_TEMPLATE_FIELDS))
  },
  {
    id: 'tools-and-equipments',
    name: 'Tools and equipments',
    tabName: 'Tools and equipments',
    description: 'IT equipment (laptops, monitors), power tools, testing kits, measurement gauges & workshop equipment',
    icon: 'Wrench',
    color: '#f59e0b', // Amber
    fields: JSON.parse(JSON.stringify(STANDARD_TEMPLATE_FIELDS))
  },
  {
    id: 'furniture-and-fittings',
    name: 'Furniture and fittings',
    tabName: 'Furniture and fittings',
    description: 'Executive workstations, ergonomic task chairs, conference tables, storage units & HVAC fixtures',
    icon: 'Armchair',
    color: '#8b5cf6', // Purple
    fields: JSON.parse(JSON.stringify(STANDARD_TEMPLATE_FIELDS))
  }
];

export const INITIAL_ASSETS = [
  // 1. Machinery
  {
    id: 'AST-MC-2001',
    categoryId: 'machinery',
    categoryName: 'Machinery',
    asset_name: 'Haas 5-Axis Precision CNC Milling Center',
    asset_number: 'AST-MC-2001',
    asset_id: 'AST-MC-2001',
    title: 'Component Precision Machining Line 1',
    asset_type: 'CNC / Milling Unit',
    category: 'Machinery',
    brand: 'Haas Automation',
    model: 'VF-4SS Super-Speed CNC',
    serial_no: 'HAAS-2023-9941',
    barcode: '88492039001',
    qr_code: 'QR-AST-MC-2001',
    description: '30 kW 12,000 RPM inline direct-drive spindle with high-speed side-mount tool changer',
    company: 'Orel Corp',
    department: 'Operations',
    location: 'Main Factory, Machining Bay 1',
    purchase_date: '2023-01-20',
    purchase_price: 165000,
    purchase_cost: 165000,
    vendor: 'Haas Asia Pacific Automation',
    invoice_no: 'HAAS-INV-2023-014',
    warranty_expiry: '2025-01-20',
    warranty_provider: 'Haas Certified Industrial Warranty',
    depreciation_method: 'Units of Production',
    useful_life_months: 120,
    residual_value: 25000,
    status: 'Active / In Use'
  },
  {
    id: 'AST-MC-2002',
    categoryId: 'machinery',
    categoryName: 'Machinery',
    asset_name: 'Toyota HiAce Commercial Cargo Van 2.8L',
    asset_number: 'AST-MC-2002',
    asset_id: 'AST-MC-2002',
    title: 'Inter-Branch Daily Dispatch Van',
    asset_type: 'Commercial Vehicle / Van',
    category: 'Machinery',
    brand: 'Toyota',
    model: 'HiAce High Roof 2.8L Turbo Diesel',
    serial_no: 'JTE-HA200-99410',
    barcode: '88492029001',
    qr_code: 'QR-AST-MC-2002',
    description: 'Automatic transmission, high-capacity payload 1,400kg, GPS tracking installed',
    company: 'Orel Corp',
    department: 'Logistics & Transport',
    location: 'Central Logistics Terminal, Bay 2',
    purchase_date: '2022-09-10',
    purchase_price: 48500,
    purchase_cost: 48500,
    vendor: 'Toyota Lanka (Pvt) Ltd',
    invoice_no: 'INV-TL-2022-098',
    warranty_expiry: '2025-09-10',
    warranty_provider: 'Toyota 3-Yr / 100,000km Warranty',
    depreciation_method: 'Reducing Balance',
    useful_life_months: 84,
    residual_value: 12000,
    status: 'Active / In Use'
  },
  {
    id: 'AST-MC-2003',
    categoryId: 'machinery',
    categoryName: 'Machinery',
    asset_name: 'Automated Surface SMT Pick & Place System',
    asset_number: 'AST-MC-2003',
    asset_id: 'AST-MC-2003',
    title: 'Cleanroom PCB Assembly Line',
    asset_type: 'Production Line',
    category: 'Machinery',
    brand: 'Yamaha Robotics',
    model: 'YSM20R High-Speed SMT',
    serial_no: 'YM-SMT-7720-X',
    barcode: '88492039003',
    qr_code: 'QR-AST-MC-2003',
    description: 'Dual-beam modular mounter with 90,000 CPH high-precision component placement',
    company: 'Orel Corp',
    department: 'Engineering',
    location: 'Cleanroom Production Line 1',
    purchase_date: '2022-11-05',
    purchase_price: 210000,
    purchase_cost: 210000,
    vendor: 'Yamaha Industrial Robotics',
    invoice_no: 'INV-YM-2022-551',
    warranty_expiry: '2025-11-05',
    warranty_provider: 'Yamaha Robotics Certified Service',
    depreciation_method: 'Straight Line',
    useful_life_months: 120,
    residual_value: 30000,
    status: 'Active / In Use'
  },

  // 2. Land and building
  {
    id: 'AST-LB-1001',
    categoryId: 'land-and-building',
    categoryName: 'Land and building',
    asset_name: 'Corporate Headquarters & Innovation Campus',
    asset_number: 'AST-LB-1001',
    asset_id: 'AST-LB-1001',
    title: 'Central Corporate Administration Campus',
    asset_type: 'Office Building',
    category: 'Land and building',
    brand: 'OREL Properties',
    model: '5-Story Commercial Campus (45,000 sq ft)',
    serial_no: 'DEED-CA-88921-A',
    barcode: '88492069001',
    qr_code: 'QR-AST-LB-1001',
    description: 'Executive corporate tower with integrated solar rooftop array, high-density zone C-4',
    company: 'Orel Corp',
    department: 'Executive',
    location: '450 Innovation Blvd, Silicon District, Colombo',
    purchase_date: '2021-03-15',
    purchase_price: 3450000,
    purchase_cost: 3450000,
    vendor: 'Prime Lands Residencies PLC',
    invoice_no: 'INV-PLR-2021-08',
    warranty_expiry: '2041-03-15',
    warranty_provider: 'Architectural Structural Guarantee',
    depreciation_method: 'Straight Line',
    useful_life_months: 360,
    residual_value: 800000,
    status: 'Active / In Use'
  },
  {
    id: 'AST-LB-1002',
    categoryId: 'land-and-building',
    categoryName: 'Land and building',
    asset_name: 'Regional Logistics Center & Central Warehouse',
    asset_number: 'AST-LB-1002',
    asset_id: 'AST-LB-1002',
    title: 'Central Supply Chain Hub',
    asset_type: 'Warehouse / Logistics Hub',
    category: 'Land and building',
    brand: 'OREL Properties',
    model: '68,000 sq ft High-Bay Storage Facility',
    serial_no: 'DEED-TX-44012-B',
    barcode: '88492069002',
    qr_code: 'QR-AST-LB-1002',
    description: 'Heavy industrial grade M-1 logistics complex with 8 automated loading bays',
    company: 'Orel Corp',
    department: 'Logistics & Transport',
    location: 'Plot 12, Expressway Logistics Park, Biyagama',
    purchase_date: '2022-08-10',
    purchase_price: 1800000,
    purchase_cost: 1800000,
    vendor: 'Commercial Logistics Parks Ltd',
    invoice_no: 'INV-CLP-2022-44',
    warranty_expiry: '2042-08-10',
    warranty_provider: 'Industrial Facilities Guarantee',
    depreciation_method: 'Straight Line',
    useful_life_months: 360,
    residual_value: 400000,
    status: 'Active / In Use'
  },

  // 3. Tools and equipments
  {
    id: 'AST-TE-3001',
    categoryId: 'tools-and-equipments',
    categoryName: 'Tools and equipments',
    asset_name: 'Apple MacBook Pro 16" M2 Max',
    asset_number: 'AST-TE-3001',
    asset_id: 'AST-TE-3001',
    title: 'Lead Architect Dev Machine',
    asset_type: 'Laptop',
    category: 'Tools and equipments',
    brand: 'Apple',
    model: 'MacBook Pro 16" (Space Black, 32GB / 1TB)',
    serial_no: 'C02G901XYZ8',
    barcode: '88492019001',
    qr_code: 'QR-AST-TE-3001',
    description: 'Apple M2 Max chip with 12-core CPU and 38-core GPU, 32GB Unified Memory, 1TB SSD',
    company: 'Orel Corp',
    department: 'IT & Software',
    location: 'Level 4, Tech Lab, Colombo HQ',
    purchase_date: '2023-04-15',
    purchase_price: 2850,
    purchase_cost: 2850,
    vendor: 'Softlogic Information Systems',
    invoice_no: 'INV-2023-4410',
    warranty_expiry: '2026-04-15',
    warranty_provider: 'AppleCare Enterprise Support',
    depreciation_method: 'Straight Line',
    useful_life_months: 36,
    residual_value: 450,
    status: 'Active / In Use'
  },
  {
    id: 'AST-TE-3002',
    categoryId: 'tools-and-equipments',
    categoryName: 'Tools and equipments',
    asset_name: 'Dell UltraSharp 32" 4K Video Conferencing Monitor',
    asset_number: 'AST-TE-3002',
    asset_id: 'AST-TE-3002',
    title: 'Product Engineering Workstation Display',
    asset_type: 'Monitor',
    category: 'Tools and equipments',
    brand: 'Dell',
    model: 'U3223QE 4K IPS Black',
    serial_no: 'CN-0P9214-8841',
    barcode: '88492019002',
    qr_code: 'QR-AST-TE-3002',
    description: '31.5-inch 4K UHD hub monitor with IPS Black technology, 90W USB-C power delivery',
    company: 'Orel Corp',
    department: 'Engineering',
    location: 'Level 3, Engineering Bay #12',
    purchase_date: '2023-06-20',
    purchase_price: 920,
    purchase_cost: 920,
    vendor: 'Dell Technologies Direct',
    invoice_no: 'INV-DELL-2023-559',
    warranty_expiry: '2026-06-20',
    warranty_provider: 'Dell 3-Year ProSupport',
    depreciation_method: 'Straight Line',
    useful_life_months: 48,
    residual_value: 120,
    status: 'Active / In Use'
  },
  {
    id: 'AST-TE-3003',
    categoryId: 'tools-and-equipments',
    categoryName: 'Tools and equipments',
    asset_name: 'High-Precision Digital Torque Analyzer & Calibration Kit',
    asset_number: 'AST-TE-3003',
    asset_id: 'AST-TE-3003',
    title: 'QA Workshop Calibration Station',
    asset_type: 'Testing & Calibration Kit',
    category: 'Tools and equipments',
    brand: 'Norbar',
    model: 'TruCheck 2 Plus',
    serial_no: 'DTA-900-54',
    barcode: '88492059001',
    qr_code: 'QR-AST-TE-3003',
    description: 'Precision digital torque tester with ISO 6789 calibration certificate and RS-232 interface',
    company: 'Orel Corp',
    department: 'Engineering',
    location: 'Quality Control Calibration Lab #2',
    purchase_date: '2023-06-18',
    purchase_price: 3850,
    purchase_cost: 3850,
    vendor: 'Norbar Industrial Instruments',
    invoice_no: 'INV-NOR-2023-77',
    warranty_expiry: '2025-06-18',
    warranty_provider: 'Norbar Calibration Support',
    depreciation_method: 'Straight Line',
    useful_life_months: 60,
    residual_value: 300,
    status: 'Active / In Use'
  },
  {
    id: 'AST-TE-3004',
    categoryId: 'tools-and-equipments',
    categoryName: 'Tools and equipments',
    asset_name: 'Heavy Duty Rotary Hammer Drill Set 18V',
    asset_number: 'AST-TE-3004',
    asset_id: 'AST-TE-3004',
    title: 'Facility Maintenance Crew Kit',
    asset_type: 'Power Tool',
    category: 'Tools and equipments',
    brand: 'Bosch Professional',
    model: 'GBH 18V-26 D-Handle',
    serial_no: 'BOSCH-GBH-18V-88',
    barcode: '88492059004',
    qr_code: 'QR-AST-TE-3004',
    description: '18V cordless brushless hammer drill with kickback control and dust extractor adapter',
    company: 'Orel Corp',
    department: 'Facilities & Maintenance',
    location: 'Maintenance Crew Mobile Van #2',
    purchase_date: '2024-02-14',
    purchase_price: 850,
    purchase_cost: 850,
    vendor: 'Brown & Company PLC',
    invoice_no: 'INV-BRN-2024-110',
    warranty_expiry: '2026-02-14',
    warranty_provider: 'Bosch 2-Year Warranty',
    depreciation_method: 'Straight Line',
    useful_life_months: 48,
    residual_value: 100,
    status: 'Active / In Use'
  },

  // 4. Furniture and fittings
  {
    id: 'AST-FF-4001',
    categoryId: 'furniture-and-fittings',
    categoryName: 'Furniture and fittings',
    asset_name: 'Steelcase Executive Boardroom Conference Table (16-Seater)',
    asset_number: 'AST-FF-4001',
    asset_id: 'AST-FF-4001',
    title: 'Executive Boardroom Central Suite',
    asset_type: 'Conference Table',
    category: 'Furniture and fittings',
    brand: 'Steelcase',
    model: 'Potrero415 Architectural Suite',
    serial_no: 'SC-CONF-8812',
    barcode: '88492049001',
    qr_code: 'QR-AST-FF-4001',
    description: 'Solid walnut finish with seamless integrated AV cable wells and dual AC power hubs',
    company: 'Orel Corp',
    department: 'Executive',
    location: 'Corporate HQ, Level 5, Boardroom Alpha',
    purchase_date: '2022-11-15',
    purchase_price: 12500,
    purchase_cost: 12500,
    vendor: 'Steelcase Architectural Furnishings',
    invoice_no: 'SC-SL-2022-887',
    warranty_expiry: '2032-11-15',
    warranty_provider: 'Steelcase 10-Year Global Warranty',
    depreciation_method: 'Straight Line',
    useful_life_months: 120,
    residual_value: 1500,
    status: 'Active / In Use'
  },
  {
    id: 'AST-FF-4002',
    categoryId: 'furniture-and-fittings',
    categoryName: 'Furniture and fittings',
    asset_name: 'Herman Miller Ergonomic Task Chairs (Batch of 25)',
    asset_number: 'AST-FF-4002',
    asset_id: 'AST-FF-4002',
    title: 'Open Office Ergonomic Seating Pack',
    asset_type: 'Ergonomic Chair',
    category: 'Furniture and fittings',
    brand: 'Herman Miller',
    model: 'Aeron Fully Loaded PostureFit SL',
    serial_no: 'HM-AERON-2023-B25',
    barcode: '88492049002',
    qr_code: 'QR-AST-FF-4002',
    description: 'Pellicle breathable mesh back with harmonic 2 tilt and fully adjustable arms',
    company: 'Orel Corp',
    department: 'Operations',
    location: 'Open Office Wing B, Level 3',
    purchase_date: '2023-04-10',
    purchase_price: 16250,
    purchase_cost: 16250,
    vendor: 'Herman Miller Commercial Interiors',
    invoice_no: 'INV-HM-2023-88',
    warranty_expiry: '2035-04-10',
    warranty_provider: 'Herman Miller 12-Yr Warranty',
    depreciation_method: 'Straight Line',
    useful_life_months: 120,
    residual_value: 2000,
    status: 'Active / In Use'
  }
];

export const Storage = {
  getCategories() {
    try {
      const data = localStorage.getItem(STORAGE_KEYS.CATEGORIES);
      if (data) {
        const parsed = JSON.parse(data);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Check if categories contain the exact 4 main categories
          const hasMain4 = MAIN_4_CATEGORIES.every(mName => 
            parsed.some(c => c.name.toLowerCase() === mName.toLowerCase())
          );
          if (hasMain4) {
            return parsed;
          }
        }
      }
    } catch (e) {
      console.error('Error loading categories:', e);
    }

    // Default categories initialized to the 4 Main Categories
    this.saveCategories(DEFAULT_CATEGORIES);
    return DEFAULT_CATEGORIES;
  },

  saveCategories(categories) {
    try {
      localStorage.setItem(STORAGE_KEYS.CATEGORIES, JSON.stringify(categories));
    } catch (e) {
      console.error('Error saving categories:', e);
    }
  },

  addCategory(category) {
    const categories = this.getCategories();
    const tabName = (category.name || '').trim();
    if (!tabName) throw new Error('Category name cannot be empty.');

    const id = category.id || tabName.toLowerCase().replace(/[^a-z0-9]/g, '-');
    
    // Check if category exists
    const existingIndex = categories.findIndex(c => c.id === id || c.name.toLowerCase() === tabName.toLowerCase());
    if (existingIndex >= 0) {
      return categories[existingIndex];
    }

    const newCategory = {
      id,
      name: tabName,
      tabName: tabName,
      description: category.description || `Asset collection for ${tabName}`,
      icon: category.icon || 'Folder',
      color: category.color || '#06b6d4',
      fields: category.fields && category.fields.length > 0 
        ? category.fields 
        : JSON.parse(JSON.stringify(STANDARD_TEMPLATE_FIELDS))
    };

    categories.push(newCategory);
    this.saveCategories(categories);
    return newCategory;
  },

  addFieldToCategory(categoryId, field) {
    const categories = this.getCategories();
    const cat = categories.find(c => c.id === categoryId);
    if (!cat) throw new Error(`Category not found: ${categoryId}`);

    const fieldId = field.id || field.label.toLowerCase().replace(/[^a-z0-9]/g, '_');
    if (cat.fields.some(f => f.id === fieldId)) {
      throw new Error(`Field "${field.label}" already exists in this category.`);
    }

    const newField = {
      id: fieldId,
      label: field.label.trim(),
      type: field.type || 'text',
      options: field.options || [],
      required: !!field.required,
      isDefault: false,
      group: 'Custom Fields'
    };

    cat.fields.push(newField);
    this.saveCategories(categories);
    return cat;
  },

  deleteFieldFromCategory(categoryId, fieldId) {
    const categories = this.getCategories();
    const cat = categories.find(c => c.id === categoryId);
    if (!cat) return;

    cat.fields = cat.fields.filter(f => f.id !== fieldId || f.isDefault);
    this.saveCategories(categories);
    return cat;
  },

  getAssets() {
    try {
      const data = localStorage.getItem(STORAGE_KEYS.ASSETS);
      if (data) {
        const parsed = JSON.parse(data);
        if (Array.isArray(parsed)) {
          // Remove only untouched built-in examples; preserve edited and user-created records.
          const assets = parsed.filter(asset => !INITIAL_ASSETS.some(example =>
            Object.keys(asset).length === Object.keys(example).length &&
            Object.keys(example).every(key => asset[key] === example[key])
          ));
          if (assets.length !== parsed.length) this.saveAssets(assets);
          return assets.map(a => ({
            ...a,
            asset_number: a.asset_number || a.asset_id || a.id,
            asset_id: a.asset_id || a.asset_number || a.id,
            purchase_price: a.purchase_price !== undefined ? a.purchase_price : a.purchase_cost,
            purchase_cost: a.purchase_cost !== undefined ? a.purchase_cost : a.purchase_price
          }));
        }
      }
    } catch (e) {
      console.error('Error loading assets:', e);
    }

    return [];
  },

  saveAssets(assets) {
    try {
      localStorage.setItem(STORAGE_KEYS.ASSETS, JSON.stringify(assets));
    } catch (e) {
      throw new Error('Unable to save assets. Browser storage may be full; try smaller photos.');
    }
  },

  addAsset(assetData) {
    const assets = this.getAssets();
    const categories = this.getCategories();
    const category = categories.find(c => c.id === assetData.categoryId || c.name.toLowerCase() === String(assetData.categoryName || '').toLowerCase());

    // Auto generate Asset Number if blank
    let assetNum = assetData.asset_number || assetData.asset_id || assetData.id;
    if (!assetNum) {
      const prefix = category ? category.name.substring(0, 2).toUpperCase() : 'AST';
      const randomNum = Math.floor(1000 + Math.random() * 9000);
      assetNum = `AST-${prefix}-${randomNum}`;
    }

    const price = assetData.purchase_price !== undefined && assetData.purchase_price !== ''
      ? parseFloat(assetData.purchase_price) || 0
      : (parseFloat(assetData.purchase_cost) || 0);

    const newAsset = {
      ...assetData,
      id: assetNum,
      asset_id: assetNum,
      asset_number: assetNum,
      purchase_price: price,
      purchase_cost: price,
      categoryId: category ? category.id : (assetData.categoryId || 'machinery'),
      categoryName: category ? category.name : (assetData.categoryName || assetData.Category || 'Machinery'),
      category: category ? category.name : (assetData.category || assetData.categoryName || 'Machinery'),
      company: assetData.company || assetData.Company || '',
      createdAt: assetData.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    // Do not replace an existing asset when adding a new one.
    const existingIdx = assets.findIndex(a => a.id === assetNum || a.asset_number === assetNum);
    if (existingIdx >= 0) {
      throw new Error(`Asset number "${assetNum}" already exists.`);
    }
    assets.unshift(newAsset);

    this.saveAssets(assets);
    return newAsset;
  },

  updateAsset(assetId, updatedData, expectedSnapshot = null) {
    const assets = this.getAssets();
    const index = assets.findIndex(a => a.id === assetId || a.asset_number === assetId);
    if (index === -1) throw new Error(`Asset not found: ${assetId}`);
    if (updatedData.asset_number && assets.some((a, i) => i !== index && (a.id === updatedData.asset_number || a.asset_number === updatedData.asset_number))) {
      throw new Error(`Asset number "${updatedData.asset_number}" already exists.`);
    }
    if (expectedSnapshot && JSON.stringify(assets[index]) !== expectedSnapshot) {
      throw new Error('This asset has changed since you opened it. Close the form and reopen it to review the latest version.');
    }

    const price = updatedData.purchase_price !== undefined && updatedData.purchase_price !== ''
      ? parseFloat(updatedData.purchase_price) || 0
      : (parseFloat(updatedData.purchase_cost) || 0);

    assets[index] = {
      ...assets[index],
      ...updatedData,
      id: updatedData.asset_number || assetId,
      asset_id: updatedData.asset_number || assetId,
      asset_number: updatedData.asset_number || assetId,
      purchase_price: price,
      purchase_cost: price,
      updatedAt: new Date().toISOString()
    };

    this.saveAssets(assets);
    return assets[index];
  },

  deleteAsset(assetId) {
    let assets = this.getAssets();
    assets = assets.filter(a => a.id !== assetId && a.asset_number !== assetId);
    this.saveAssets(assets);
  },

  getSettings() {
    try {
      const data = localStorage.getItem(STORAGE_KEYS.SETTINGS);
      if (data) return JSON.parse(data);
    } catch (e) {
      console.error('Error loading settings:', e);
    }
    return {
      googleAppsScriptUrl: '',
      autoSync: false,
      lastSyncTimestamp: null,
      currencySymbol: '$',
      companyName: 'Orel Corp'
    };
  },

  saveSettings(settings) {
    try {
      localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(settings));
    } catch (e) {
      console.error('Error saving settings:', e);
    }
  },

  resetToDefault() {
    localStorage.removeItem(STORAGE_KEYS.CATEGORIES);
    localStorage.removeItem(STORAGE_KEYS.ASSETS);
    return {
      categories: this.getCategories(),
      assets: this.getAssets()
    };
  }
};
