// main.js - Master Application Controller for OREL Asset Management
import './style.css';
import { createIcons, icons } from 'lucide';
import { Storage, DEFAULT_CATEGORIES, STANDARD_TEMPLATE_FIELDS } from './storage.js';
import { GoogleSheets, GOOGLE_APPS_SCRIPT_CODE } from './googleSheets.js';
import { ExcelHandler } from './excelHandler.js';

// Application State
const state = {
  categories: [],
  assets: [],
  settings: {},
  activeCategoryId: 'machinery',
  currentView: 'cards', // 'cards' | 'table'
  searchQuery: '',
  statusFilter: 'ALL',
  assetTypeFilter: 'ALL',
  sortBy: 'newest',
  selectedAssetForDetail: null,
  pendingExcelImport: null
};

let editingAssetSnapshot = null;
let assetPhotos = ['', ''];

function renderPhotoFields(container) {
  const section = document.createElement('section');
  section.className = 'form-section-card';
  section.innerHTML = `<div class="form-section-header"><span class="form-section-title">Asset photos</span></div>
    <p class="photo-help">Choose or take up to two photos. Photos are resized to save space.</p><div class="photo-grid"></div>`;
  const grid = section.querySelector('.photo-grid');
  assetPhotos.forEach((photo, index) => {
    const slot = document.createElement('div');
    slot.className = 'photo-slot';
    const label = document.createElement('label');
    label.textContent = `Photo ${index + 1}`;
    const preview = document.createElement('img');
    preview.className = 'photo-preview';
    preview.alt = `Asset photo ${index + 1} preview`;
    preview.hidden = !photo;
    if (photo) preview.src = photo;
    const input = document.createElement('input');
    let readVersion = 0;
    input.type = 'file';
    input.accept = 'image/jpeg,image/png,image/webp';
    input.className = 'form-control photo-input';
    label.appendChild(input);
    const cameraLabel = document.createElement('label');
    cameraLabel.className = 'btn btn-secondary photo-camera';
    cameraLabel.textContent = 'Take photo';
    const camera = document.createElement('input');
    camera.type = 'file';
    camera.accept = 'image/*';
    camera.setAttribute('capture', 'environment');
    camera.className = 'photo-camera-input';
    cameraLabel.appendChild(camera);
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'btn btn-secondary photo-remove';
    remove.textContent = 'Remove photo';
    remove.hidden = !photo;
    remove.addEventListener('click', () => {
      readVersion++;
      assetPhotos[index] = '';
      input.value = '';
      camera.value = '';
      input.setCustomValidity('');
      camera.setCustomValidity('');
      preview.hidden = true;
      remove.hidden = true;
    });
    const loadPhoto = (source) => {
      const version = ++readVersion;
      input.setCustomValidity('');
      camera.setCustomValidity('');
      const file = source.files[0];
      if (!file) return;
      if (!file.type.startsWith('image/')) {
        source.value = '';
        showToast('Please choose an image.', 'error');
        return;
      }
      source.setCustomValidity('Please wait for the photo to finish loading.');
      const url = URL.createObjectURL(file);
      const image = new Image();
      image.onload = () => {
        URL.revokeObjectURL(url);
        if (version !== readVersion || !source.isConnected) return;
        try {
          const scale = Math.min(1, 1200 / Math.max(image.width, image.height));
          const canvas = document.createElement('canvas');
          canvas.width = Math.max(1, Math.round(image.width * scale));
          canvas.height = Math.max(1, Math.round(image.height * scale));
          canvas.getContext('2d').drawImage(image, 0, 0, canvas.width, canvas.height);
          const data = canvas.toDataURL('image/jpeg', 0.7);
          assetPhotos[index] = data;
          preview.src = data;
          preview.hidden = false;
          remove.hidden = false;
        } catch (error) {
          showToast('Could not process the photo. Please try another image.', 'error');
        }
        source.setCustomValidity('');
        source.value = '';
      };
      image.onerror = () => {
        URL.revokeObjectURL(url);
        if (version !== readVersion || !source.isConnected) return;
        source.setCustomValidity('');
        source.value = '';
        showToast('Could not read the photo. Please try another image.', 'error');
      };
      image.src = url;
    };
    input.addEventListener('change', () => loadPhoto(input));
    camera.addEventListener('change', () => loadPhoto(camera));
    slot.append(label, cameraLabel, preview, remove);
    grid.appendChild(slot);
  });
  container.appendChild(section);
}

const categoryLabels = {
  'tools-and-equipments': 'Tools and Equipments',
  'land-and-building': 'Land and Buildings',
  'machinery': 'Machinery',
  'furniture-and-fittings': 'Furniture and Fittings'
};
const categoryLabel = category => categoryLabels[category.id] || category.name;

// Refresh Lucide icons anywhere in the DOM
function refreshIcons() {
  createIcons({ icons });
}

// Show Toast Notification
function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  
  let iconName = 'check-circle';
  if (type === 'error') iconName = 'alert-circle';
  if (type === 'info') iconName = 'info';

  toast.innerHTML = `
    <i data-lucide="${iconName}"></i>
    <span>${message}</span>
  `;

  container.appendChild(toast);
  refreshIcons();

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.25s ease';
    setTimeout(() => toast.remove(), 250);
  }, 3500);
}

// Format Currency
function formatCurrency(amount) {
  const num = parseFloat(amount) || 0;
  return new Intl.NumberFormat('en-LK', {
    style: 'currency',
    currency: 'LKR',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(num);
}

// Format Date
function formatDate(dateStr) {
  if (!dateStr) return 'N/A';
  try {
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? dateStr : d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
  } catch (e) {
    return dateStr;
  }
}

// Get active category object
function getActiveCategory() {
  return state.categories.find(c => c.id === state.activeCategoryId) || state.categories[0];
}

// Initialize Application
function initApp() {
  state.categories = Storage.getCategories();
  state.assets = Storage.getAssets();
  state.settings = Storage.getSettings();

  // Set active category to first if current not found
  if (!state.categories.some(c => c.id === state.activeCategoryId)) {
    state.activeCategoryId = state.categories[0]?.id || 'machinery';
  }

  // Bind DOM Event Listeners
  setupEventListeners();

  // Update initial UI
  updateSyncStatusDisplay();
  renderStatsBanner();
  renderCategoryChips();
  renderActiveCategoryToolbar();
  renderAssets();

  // Setup Script Preview Box
  const scriptPreviewEl = document.getElementById('scriptCodePreview');
  if (scriptPreviewEl) {
    scriptPreviewEl.textContent = GOOGLE_APPS_SCRIPT_CODE;
  }

  const scriptInputEl = document.getElementById('googleScriptUrlInput');
  if (scriptInputEl && state.settings.googleAppsScriptUrl) {
    scriptInputEl.value = state.settings.googleAppsScriptUrl;
  }

  refreshIcons();
}

// Update Top Google Sheets Sync Status indicator
function updateSyncStatusDisplay() {
  const syncDot = document.getElementById('syncDot');
  const syncStatusText = document.getElementById('syncStatusText');
  if (!syncDot || !syncStatusText) return;

  if (state.settings.googleAppsScriptUrl) {
    syncDot.className = 'sync-dot connected';
    syncStatusText.textContent = 'Sheets: Connected';
  } else {
    syncDot.className = 'sync-dot';
    syncStatusText.textContent = 'Sheets: Offline (Local)';
  }
}

// Render Top Dashboard Statistics Banner
function renderStatsBanner() {
  const totalAssetsEl = document.getElementById('statTotalAssets');
  const totalValueEl = document.getElementById('statTotalValue');
  const activeAssetsEl = document.getElementById('statActiveAssets');
  const tabsCountEl = document.getElementById('statTabsCount');
  const categoryCountEl = document.getElementById('statCategoryCount');

  const totalAssets = state.assets.length;
  const totalValuation = state.assets.reduce((sum, a) => sum + (parseFloat(a.purchase_price !== undefined ? a.purchase_price : a.purchase_cost) || 0), 0);
  
  const activeCount = state.assets.filter(a => {
    const s = String(a.status || '').toLowerCase();
    if (s === 'not currently in use') return false;
    return s.includes('active') || s.includes('operational') || s.includes('in use') || s.includes('available');
  }).length;

  if (totalAssetsEl) totalAssetsEl.textContent = totalAssets;
  if (totalValueEl) totalValueEl.textContent = formatCurrency(totalValuation);
  if (activeAssetsEl) activeAssetsEl.textContent = activeCount;
  if (tabsCountEl) tabsCountEl.textContent = `${state.categories.length} Tabs`;
  if (categoryCountEl) categoryCountEl.textContent = `Across ${state.categories.length} Sheet Tabs`;
}

// Render Category Chips
function renderCategoryChips() {
  const container = document.getElementById('categoryChipsContainer');
  if (!container) return;

  container.innerHTML = '';

  state.categories.forEach(cat => {
    const chip = document.createElement('button');
    chip.type = 'button';
    const isActive = cat.id === state.activeCategoryId;
    chip.className = `category-chip ${isActive ? 'active' : ''}`;
    
    // Calculate count of assets in this category
    const count = state.assets.filter(a => a.categoryId === cat.id || a.categoryName === cat.name).length;

    // Pick icon
    let iconName = 'folder';
    if (cat.icon) {
      iconName = cat.icon === 'Building2' ? 'building-2' : cat.icon.toLowerCase();
    } else if (cat.name.toLowerCase().includes('it') || cat.name.toLowerCase().includes('tech') || cat.name.toLowerCase().includes('laptop')) {
      iconName = 'laptop';
    } else if (cat.name.toLowerCase().includes('vehicle') || cat.name.toLowerCase().includes('car') || cat.name.toLowerCase().includes('fleet')) {
      iconName = 'truck';
    } else if (cat.name.toLowerCase().includes('machin') || cat.name.toLowerCase().includes('plant')) {
      iconName = 'cog';
    } else if (cat.name.toLowerCase().includes('tool') || cat.name.toLowerCase().includes('instrument')) {
      iconName = 'wrench';
    } else if (cat.name.toLowerCase().includes('furniture')) {
      iconName = 'armchair';
    } else if (cat.name.toLowerCase().includes('land') || cat.name.toLowerCase().includes('building')) {
      iconName = 'building-2';
    }

    chip.innerHTML = `
      <div class="category-chip-icon" style="color: ${cat.color || 'var(--accent-emerald-light)'}">
        <i data-lucide="${iconName}"></i>
      </div>
      <span class="category-card-content"><strong></strong><small></small><span class="category-card-link">View assets →</span></span>
      <span class="category-chip-badge">${count} assets</span>
    `;
    chip.querySelector('strong').textContent = categoryLabel(cat);
    chip.querySelector('small').textContent = cat.description || 'Manage assets in this category';
    chip.setAttribute('aria-pressed', String(isActive));

    chip.addEventListener('click', () => {
      state.activeCategoryId = cat.id;
      renderCategoryChips();
      renderActiveCategoryToolbar();
      renderAssets();
    });

    container.appendChild(chip);
  });

  refreshIcons();
}

// Render Active Category Toolbar Details
function renderActiveCategoryToolbar() {
  const activeCat = getActiveCategory();
  if (!activeCat) return;

  const activeTabNameDisplay = document.getElementById('activeTabNameDisplay');
  const activeCategoryFieldCount = document.getElementById('activeCategoryFieldCount');

  if (activeTabNameDisplay) {
    activeTabNameDisplay.textContent = categoryLabel(activeCat);
  }
  if (activeCategoryFieldCount) {
    activeCategoryFieldCount.textContent = `(${activeCat.fields.length} Fields configured)`;
  }
}

// Render Asset Cards or Table
function renderAssets() {
  const container = document.getElementById('assetsContainer');
  if (!container) return;

  const activeCat = getActiveCategory();
  if (!activeCat) return;

  // Filter assets for this category
  let filtered = state.assets.filter(a => a.categoryId === activeCat.id || a.categoryName === activeCat.name);

  // Apply Asset Type filter
  if (state.assetTypeFilter && state.assetTypeFilter !== 'ALL') {
    filtered = filtered.filter(a => {
      const t = String(a.asset_type || a['Asset Type'] || '').toLowerCase();
      return t === state.assetTypeFilter.toLowerCase();
    });
  }

  // Apply Status filter
  if (state.statusFilter !== 'ALL') {
    filtered = filtered.filter(a => {
      const s = String(a.status || '').toLowerCase();
      if (state.statusFilter === 'Active') return s !== 'not currently in use' && (s.includes('active') || s.includes('operational') || s.includes('in use') || s.includes('available'));
      if (state.statusFilter === 'Maintenance') return s.includes('maintenance') || s.includes('calibration') || s.includes('repair');
      if (state.statusFilter === 'Storage') return s.includes('storage') || s.includes('standby');
      if (state.statusFilter === 'In Stock') return s.includes('in stock');
      if (state.statusFilter === 'Not Currently in Use') return s === 'not currently in use';
      if (state.statusFilter === 'Disposed') return s.includes('disposed') || s.includes('damaged') || s.includes('breakdown');
      return true;
    });
  }

  // Apply Search query
  if (state.searchQuery.trim()) {
    const q = state.searchQuery.toLowerCase().trim();
    filtered = filtered.filter(a => {
      return Object.values(a).some(val => String(val).toLowerCase().includes(q));
    });
  }

  // Apply Sort
  filtered.sort((a, b) => {
    const priceA = parseFloat(a.purchase_price !== undefined ? a.purchase_price : a.purchase_cost) || 0;
    const priceB = parseFloat(b.purchase_price !== undefined ? b.purchase_price : b.purchase_cost) || 0;
    if (state.sortBy === 'cost-desc') return priceB - priceA;
    if (state.sortBy === 'cost-asc') return priceA - priceB;
    if (state.sortBy === 'name') return String(a.asset_name || '').localeCompare(String(b.asset_name || ''));
    return new Date(b.createdAt || b.purchase_date || 0) - new Date(a.createdAt || a.purchase_date || 0);
  });

  // Empty State Check
  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">
          <i data-lucide="package-open" style="width: 32px; height: 32px;"></i>
        </div>
        <h3>No Assets Found in "${activeCat.name}"</h3>
        <p>No matching assets found. Add an asset to get started.</p>
        <button class="btn btn-primary" id="btnEmptyAddAsset">
          <i data-lucide="plus"></i>
          <span>Add Asset in ${activeCat.name}</span>
        </button>
      </div>
    `;
    const btnEmpty = document.getElementById('btnEmptyAddAsset');
    if (btnEmpty) btnEmpty.addEventListener('click', () => openAssetModal());
    refreshIcons();
    return;
  }

  // Card View Mode
  if (state.currentView === 'cards') {
    let html = `<div class="asset-grid">`;
    filtered.forEach(asset => {
      const statusClass = getStatusClass(asset.status);
      const assetNum = asset.asset_number || asset.asset_id || asset.id;
      const brandModel = [asset.brand, asset.model].filter(Boolean).join(' • ');
      const price = asset.purchase_price !== undefined ? asset.purchase_price : asset.purchase_cost;

      html += `
        <div class="asset-card" data-asset-id="${assetNum}">
          <div>
            <div class="asset-card-top">
              <span class="asset-id-tag">${assetNum}</span>
              <span class="status-pill ${statusClass}">
                ${asset.status || 'Active / In Use'}
              </span>
            </div>

            <h4 class="asset-title">${asset.asset_name || 'Untitled Asset'}</h4>
            ${asset.title ? `<div style="font-size: 0.8rem; color: var(--accent-emerald-light); margin-top: 2px;">${asset.title}</div>` : ''}
            ${brandModel ? `<div class="asset-brand-model"><i data-lucide="tag" style="width: 13px; height: 13px;"></i><span>${brandModel}</span></div>` : ''}

            <div style="display: flex; flex-wrap: wrap; gap: 0.35rem; margin: 0.65rem 0 0.45rem;">
              <span class="asset-specs-pill">${asset.asset_type || activeCat.name}</span>
              ${asset.department ? `<span class="asset-specs-pill">${asset.department}</span>` : ''}
              ${asset.serial_no ? `<span class="asset-specs-pill">SN: ${asset.serial_no}</span>` : ''}
              ${asset.warranty_expiry ? `<span class="asset-warranty-tag"><i data-lucide="shield-check" style="width: 12px; height: 12px;"></i>Exp: ${formatDate(asset.warranty_expiry)}</span>` : ''}
            </div>

            <div class="asset-location">
              <i data-lucide="map-pin" style="width: 14px; height: 14px;"></i>
              <span>${asset.location || 'Location Unassigned'}</span>
            </div>
          </div>

          <div class="asset-card-footer">
            <div>
              <div style="font-size: 0.7rem; color: var(--text-muted); text-transform: uppercase;">Valuation / Price</div>
              <div class="asset-cost-badge">${formatCurrency(price)}</div>
            </div>

            <div class="asset-card-actions">
              <button class="btn-card-action btn-view-asset" data-id="${assetNum}" title="View Details & Tag">
                <i data-lucide="eye" style="width: 16px; height: 16px;"></i>
              </button>
              <button class="btn-card-action btn-edit-asset" data-id="${assetNum}" title="Edit Asset">
                <i data-lucide="edit" style="width: 16px; height: 16px;"></i>
              </button>
              <button class="btn-card-action delete btn-delete-asset" data-id="${assetNum}" title="Delete Asset">
                <i data-lucide="trash-2" style="width: 16px; height: 16px;"></i>
              </button>
            </div>
          </div>
        </div>
      `;
    });
    html += `</div>`;
    container.innerHTML = html;
  } else {
    // Table View Mode
    let html = `
      <div class="table-responsive-wrapper">
        <table class="asset-table">
          <thead>
            <tr>
              <th>Asset No</th>
              <th>Asset Name</th>
              <th>Type</th>
              <th>Brand / Model</th>
              <th>Serial No</th>
              <th>Department</th>
              <th>Location</th>
              <th>Price</th>
              <th>Warranty</th>
              <th>Status</th>
              <th style="text-align: right;">Actions</th>
            </tr>
          </thead>
          <tbody>
    `;

    filtered.forEach(asset => {
      const assetNum = asset.asset_number || asset.asset_id || asset.id;
      const brandModel = [asset.brand, asset.model].filter(Boolean).join(' ') || '-';
      const price = asset.purchase_price !== undefined ? asset.purchase_price : asset.purchase_cost;

      html += `
        <tr>
          <td><span class="asset-id-tag">${assetNum}</span></td>
          <td><strong>${asset.asset_name || '-'}</strong>${asset.title ? `<br><small style="color:var(--text-muted);">${asset.title}</small>` : ''}</td>
          <td>${asset.asset_type || '-'}</td>
          <td>${brandModel}</td>
          <td><code style="font-size:0.78rem;">${asset.serial_no || '-'}</code></td>
          <td>${asset.department || '-'}</td>
          <td>${asset.location || '-'}</td>
          <td><strong style="color: var(--accent-emerald-light);">${formatCurrency(price)}</strong></td>
          <td>${asset.warranty_expiry ? formatDate(asset.warranty_expiry) : '-'}</td>
          <td><span class="status-pill ${getStatusClass(asset.status)}">${asset.status || 'Active'}</span></td>
          <td style="text-align: right;">
            <div style="display: inline-flex; gap: 0.35rem;">
              <button class="btn-card-action btn-view-asset" data-id="${assetNum}" title="View">
                <i data-lucide="eye" style="width: 15px; height: 15px;"></i>
              </button>
              <button class="btn-card-action btn-edit-asset" data-id="${assetNum}" title="Edit">
                <i data-lucide="edit" style="width: 15px; height: 15px;"></i>
              </button>
              <button class="btn-card-action delete btn-delete-asset" data-id="${assetNum}" title="Delete">
                <i data-lucide="trash-2" style="width: 15px; height: 15px;"></i>
              </button>
            </div>
          </td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;
    container.innerHTML = html;
  }

  // Attach card/row action listeners
  container.querySelectorAll('.btn-view-asset').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openAssetDetailModal(btn.dataset.id);
    });
  });

  container.querySelectorAll('.btn-edit-asset').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openAssetModal(btn.dataset.id);
    });
  });

  container.querySelectorAll('.btn-delete-asset').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      deleteAssetPrompt(btn.dataset.id);
    });
  });

  refreshIcons();
}

// Determine status pill CSS class
function getStatusClass(status) {
  const s = String(status || '').toLowerCase();
  if (s === 'not currently in use' || s.includes('in stock') || s.includes('storage')) return 'status-maintenance';
  if (s.includes('active') || s.includes('operational') || s.includes('in use') || s.includes('available')) {
    return 'status-active';
  }
  if (s.includes('maintenance') || s.includes('checked out') || s.includes('calibration') || s.includes('repair')) {
    return 'status-maintenance';
  }
  return 'status-breakdown';
}

// Open Dynamic Asset Add / Edit Modal (Organized 23 Template Fields)
function openAssetModal(assetIdToEdit = null) {
  const activeCat = getActiveCategory();
  if (!activeCat) return;

  const modalTitle = document.getElementById('assetModalTitle');
  const modalTabIndicator = document.getElementById('modalTabNameIndicator');
  const catSelect = document.getElementById('assetCategorySelect');
  const editIdInput = document.getElementById('assetEditId');
  const dynamicContainer = document.getElementById('dynamicFieldsContainer');

  // Populate category dropdown
  catSelect.innerHTML = state.categories.map(c => `
    <option value="${c.id}" ${c.id === activeCat.id ? 'selected' : ''}>${categoryLabel(c)}</option>
  `).join('');

  catSelect.onchange = (e) => {
    state.activeCategoryId = e.target.value;
    renderActiveCategoryToolbar();
    openAssetModal(assetIdToEdit);
  };

  const asset = assetIdToEdit ? state.assets.find(a => a.id === assetIdToEdit || a.asset_number === assetIdToEdit) : null;
  editingAssetSnapshot = asset ? JSON.stringify(asset) : null;
  assetPhotos = [asset?.photos?.[0] || '', asset?.photos?.[1] || ''];

  modalTitle.textContent = asset 
    ? `Edit Asset: ${asset.asset_number || asset.asset_name || asset.id}` 
    : `Add Asset to "${categoryLabel(activeCat)}"`;
  modalTabIndicator.textContent = categoryLabel(activeCat);
  editIdInput.value = asset ? (asset.asset_number || asset.id) : '';

  // Render Category Fields Grouped by Visual Sections
  dynamicContainer.innerHTML = '';

  const sectionDefs = [
    {
      id: 'Identification',
      title: '1. Asset Identification',
      icon: 'tag',
      subtitle: 'Name, Type & Tracking codes'
    },
    {
      id: 'Specifications',
      title: '2. Hardware & Specifications',
      icon: 'cpu',
      subtitle: 'Brand, Model & Serials'
    },
    {
      id: 'Organization',
      title: '3. Organization & Location',
      icon: 'building-2',
      subtitle: 'Company, Dept & Facility'
    },
    {
      id: 'Financials',
      title: '4. Procurement & Financials',
      icon: 'wallet',
      subtitle: 'Pricing, Vendor & Invoicing'
    },
    {
      id: 'Warranty & Depreciation',
      title: '5. Warranty & Depreciation',
      icon: 'shield-check',
      subtitle: 'Protection & Accounting Life'
    },
    {
      id: 'Custom Fields',
      title: '6. Custom Tab Fields',
      icon: 'sliders',
      subtitle: 'Additional Category Fields'
    }
  ];

  const fieldsByGroup = new Map();
  sectionDefs.forEach(s => fieldsByGroup.set(s.id, []));

  activeCat.fields.forEach(field => {
    const group = field.group || 'Custom Fields';
    if (!fieldsByGroup.has(group)) {
      fieldsByGroup.set(group, []);
    }
    fieldsByGroup.get(group).push(field);
  });

  sectionDefs.forEach(section => {
    const fields = fieldsByGroup.get(section.id) || [];
    if (fields.length === 0) return;

    const sectionEl = document.createElement('div');
    sectionEl.className = 'form-section-card';

    sectionEl.innerHTML = `
      <div class="form-section-header">
        <div class="form-section-icon"><i data-lucide="${section.icon}"></i></div>
        <span class="form-section-title">${section.title}</span>
        <span class="form-section-subtitle">${section.subtitle}</span>
      </div>
      <div class="form-grid-2 form-section-grid"></div>
    `;

    const gridEl = sectionEl.querySelector('.form-section-grid');

    fields.forEach(field => {
      let val = '';
      if (asset) {
        if (asset[field.id] !== undefined && asset[field.id] !== null) val = asset[field.id];
        else if (asset[field.label] !== undefined && asset[field.label] !== null) val = asset[field.label];
        else if (field.id === 'asset_number' && asset.asset_id) val = asset.asset_id;
        else if (field.id === 'purchase_price' && asset.purchase_cost) val = asset.purchase_cost;
      }

      let defaultValue = val;
      if (!asset) {
        if (field.id === 'category') defaultValue = activeCat.name;
        if (field.id === 'status') defaultValue = 'Active / In Use';
        if (field.id === 'depreciation_method') defaultValue = 'Straight Line';
        if (field.id === 'useful_life_months') defaultValue = '36';
        if (field.id === 'asset_number' || field.id === 'asset_id') {
          const prefix = activeCat.name.substring(0, 2).toUpperCase();
          defaultValue = `AST-${prefix}-${crypto.randomUUID().slice(0, 8).toUpperCase()}`;
        }
      }

      const fieldGroup = document.createElement('div');
      fieldGroup.className = 'form-group';
      if (field.type === 'textarea') {
        fieldGroup.style.gridColumn = '1 / -1';
      }

      const labelWrapper = document.createElement('div');
      labelWrapper.style.display = 'flex';
      labelWrapper.style.alignItems = 'center';
      labelWrapper.style.justifyContent = 'space-between';

      const label = document.createElement('label');
      label.className = 'form-label';
      label.textContent = field.label;
      labelWrapper.appendChild(label);

      if (field.id === 'asset_number' || field.id === 'asset_id') {
        const autoBtn = document.createElement('button');
        autoBtn.type = 'button';
        autoBtn.className = 'btn-autogen';
        autoBtn.textContent = '⚡ Auto ID';
        autoBtn.title = 'Generate next unique asset number';
        autoBtn.onclick = () => {
          const prefix = activeCat.name.substring(0, 2).toUpperCase();
          const rand = crypto.randomUUID().slice(0, 8).toUpperCase();
          const input = fieldGroup.querySelector('input');
          if (input) input.value = `AST-${prefix}-${rand}`;
        };
        labelWrapper.appendChild(autoBtn);
      }

      fieldGroup.appendChild(labelWrapper);

      let inputEl;

      if (field.type === 'select' && !['asset_type', 'company', 'department'].includes(field.id)) {
        inputEl = document.createElement('select');
        inputEl.className = 'form-control';
        inputEl.name = field.id;

        const opts = [...(field.options && field.options.length ? field.options : ['Active / In Use', 'Available / In Stock', 'Under Maintenance', 'In Storage'])];
        if (field.id === 'status') {
          ['In Stock', 'Not Currently in Use'].forEach(option => {
            if (!opts.includes(option)) opts.push(option);
          });
        }
        inputEl.innerHTML = opts.map(o => `
          <option value="${o}" ${String(defaultValue) === String(o) ? 'selected' : ''}>${o}</option>
        `).join('');
      } else if (field.type === 'textarea') {
        inputEl = document.createElement('textarea');
        inputEl.className = 'form-control';
        inputEl.name = field.id;
        inputEl.rows = 2;
        inputEl.value = defaultValue || '';
        inputEl.placeholder = field.placeholder || '';
      } else {
        inputEl = document.createElement('input');
        inputEl.className = 'form-control';
        inputEl.name = field.id;
        inputEl.value = defaultValue || '';
        inputEl.placeholder = field.placeholder || '';

        if (field.type === 'currency' || field.type === 'number') {
          inputEl.type = 'number';
          inputEl.step = field.type === 'currency' ? '0.01' : '1';
        } else if (field.type === 'date') {
          inputEl.type = 'date';
        } else {
          inputEl.type = 'text';
        }
      }

      fieldGroup.appendChild(inputEl);
      gridEl.appendChild(fieldGroup);
    });

    dynamicContainer.appendChild(sectionEl);
  });

  renderPhotoFields(dynamicContainer);
  openModal('assetModal');
  refreshIcons();
}

// Save Asset (Submit handler)
async function handleAssetFormSubmit(e) {
  e.preventDefault();
  const activeCat = getActiveCategory();
  if (!activeCat) return;

  const form = document.getElementById('assetForm');
  const editId = document.getElementById('assetEditId').value;
  const formData = new FormData(form);

  const assetData = {
    categoryId: activeCat.id,
    categoryName: activeCat.name,
    photos: [...assetPhotos]
  };

  formData.forEach((value, key) => {
    assetData[key] = value;
  });

  // Ensure asset_number exists
  if (!assetData.asset_number) {
    if (assetData.asset_id) {
      assetData.asset_number = assetData.asset_id;
    } else {
      const prefix = activeCat.name.substring(0, 2).toUpperCase();
      assetData.asset_number = `AST-${prefix}-${crypto.randomUUID().slice(0, 8).toUpperCase()}`;
    }
  }
  assetData.asset_id = assetData.asset_number;

  // Sync purchase_price and purchase_cost
  if (assetData.purchase_price !== undefined) {
    assetData.purchase_cost = assetData.purchase_price;
  }

  try {
    const currentAssets = Storage.getAssets();
    const existing = editId ? currentAssets.find(a => a.id === editId || a.asset_number === editId) : null;
    if (editId && (!existing || JSON.stringify(existing) !== editingAssetSnapshot)) {
      throw new Error('This asset has changed. Close and reopen the form to review it.');
    }
    if (currentAssets.some(a => a !== existing && (a.id === assetData.asset_number || a.asset_number === assetData.asset_number))) {
      throw new Error('This asset number already exists.');
    }
    const saveButton = document.getElementById('btnSaveAsset');
    saveButton.disabled = true;
    try {
    if (state.settings.googleAppsScriptUrl) {
      const response = await GoogleSheets.saveSingleAsset(state.settings.googleAppsScriptUrl, activeCat, { ...existing, ...assetData }, editId);
      assetData._revision = response.revision;
    }
    let savedAsset;
    if (editId) {
      savedAsset = Storage.updateAsset(editId, assetData, editingAssetSnapshot);
      showToast(`Asset "${savedAsset.asset_name}" updated successfully!`);
    } else {
      savedAsset = Storage.addAsset(assetData);
      showToast(`Asset "${savedAsset.asset_name}" saved to "${activeCat.name}" tab!`);
    }

    // Refresh Local UI
    state.assets = Storage.getAssets();
    renderStatsBanner();
    renderCategoryChips();
    renderAssets();
    closeModal('assetModal');

    if (state.settings.googleAppsScriptUrl) showToast('Saved to Google Sheets.', 'info');
    } finally {
      saveButton.disabled = false;
    }
  } catch (error) {
    showToast(`Error saving asset: ${error.message}`, 'error');
  }
}

// Delete Asset Prompt
async function deleteAssetPrompt(assetId) {
  const asset = state.assets.find(a => a.id === assetId || a.asset_number === assetId);
  if (!asset) return;

  if (confirm(`Are you sure you want to delete "${asset.asset_name || asset.asset_id}"?`)) {
    if (state.settings.googleAppsScriptUrl) {
      try {
        await GoogleSheets.deleteSingleAsset(state.settings.googleAppsScriptUrl, asset.categoryName, asset.asset_number || asset.asset_id || asset.id, asset._revision);
      } catch (error) {
        showToast(`Could not delete from Google Sheets: ${error.message}`, 'error');
        return;
      }
    }
    Storage.deleteAsset(assetId);
    state.assets = Storage.getAssets();
    showToast(`Asset deleted.`);
    renderStatsBanner();
    renderCategoryChips();
    renderAssets();

  }
}

// Open Asset Detail & Tag Modal
function openAssetDetailModal(assetId) {
  const asset = state.assets.find(a => a.id === assetId || a.asset_number === assetId || a.asset_id === assetId);
  if (!asset) return;
  state.selectedAssetForDetail = asset;

  const cat = state.categories.find(c => c.id === asset.categoryId) || getActiveCategory();
  const assetNum = asset.asset_number || asset.asset_id || asset.id;
  const price = asset.purchase_price !== undefined ? asset.purchase_price : asset.purchase_cost;

  const titleEl = document.getElementById('detailAssetTitle');
  const badgeEl = document.getElementById('detailCategoryBadge');
  const bodyEl = document.getElementById('detailModalBody');

  if (titleEl) titleEl.textContent = asset.asset_name || assetNum;
  if (badgeEl) badgeEl.textContent = `${cat.name} (${assetNum})`;

  // Generate printable asset tag barcode simulation + property list
  let html = `
    <div style="background: #0f172a; border: 2px dashed rgba(16, 185, 129, 0.4); border-radius: var(--radius-md); padding: 1.25rem; text-align: center; margin-bottom: 1.25rem;">
      <div style="font-size: 0.75rem; color: var(--accent-emerald-light); font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase;">OREL ENTERPRISE ASSET TAG</div>
      <div style="font-size: 1.25rem; font-weight: 800; font-family: var(--font-heading); margin: 0.25rem 0;">${asset.asset_name}</div>
      ${asset.brand || asset.model ? `<div style="font-size: 0.82rem; color: var(--text-secondary); margin-bottom: 0.35rem;">${[asset.brand, asset.model].filter(Boolean).join(' • ')}</div>` : ''}
      <div style="font-family: monospace; font-size: 1.25rem; letter-spacing: 0.18em; color: var(--text-primary); margin-bottom: 0.25rem;">||| | ||||| || |||| ||||| |||</div>
      <div style="font-family: monospace; font-size: 0.85rem; color: var(--accent-emerald-light); letter-spacing: 0.08em; margin-bottom: 0.5rem;">${asset.barcode || asset.serial_no || assetNum}</div>
      <div style="display: flex; justify-content: center; flex-wrap: wrap; gap: 1rem; font-size: 0.8rem; color: var(--text-secondary); border-top: 1px solid rgba(255,255,255,0.08); padding-top: 0.5rem;">
        <span>Number: <strong>${assetNum}</strong></span>
        <span>Company: <strong>${asset.company || 'Orel Corp'}</strong></span>
        <span>Location: <strong>${asset.location || 'HQ'}</strong></span>
        <span>Price: <strong>${formatCurrency(price)}</strong></span>
      </div>
    </div>

    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.75rem;">
  `;

  cat.fields.forEach(f => {
    let val = asset[f.id] !== undefined ? asset[f.id] : (asset[f.label] !== undefined ? asset[f.label] : '-');
    if (f.type === 'currency' && val !== '-') val = formatCurrency(val);
    else if (f.type === 'date' && val !== '-') val = formatDate(val);

    const isFullWidth = f.type === 'textarea' || f.id === 'description';

    html += `
      <div style="background: var(--bg-surface-elevated); padding: 0.65rem 0.85rem; border-radius: var(--radius-sm); border: 1px solid var(--border-subtle); ${isFullWidth ? 'grid-column: 1 / -1;' : ''}">
        <div style="font-size: 0.72rem; color: var(--text-muted); text-transform: uppercase; font-weight: 600;">${f.label}</div>
        <div style="font-size: 0.88rem; font-weight: 600; color: #fff; margin-top: 0.15rem; word-break: break-word;">${val || '-'}</div>
      </div>
    `;
  });

  html += `</div>`;
  bodyEl.innerHTML = html;
  const photos = (asset.photos || []).filter(photo => typeof photo === 'string' && /^data:image\/(jpeg|png|webp);base64,/i.test(photo));
  if (photos.length) {
    const gallery = document.createElement('div');
    gallery.className = 'photo-grid detail-photos';
    photos.forEach((photo, index) => {
      const image = document.createElement('img');
      image.className = 'photo-preview';
      image.alt = `Asset photo ${index + 1}`;
      image.src = photo;
      gallery.appendChild(image);
    });
    bodyEl.prepend(gallery);
  }

  openModal('assetDetailModal');
  refreshIcons();
}

// Open Custom Field Modal ("can be able to add-field as i want")
function openAddCustomFieldModal() {
  const activeCat = getActiveCategory();
  if (!activeCat) return;

  const modalCatName = document.getElementById('fieldModalCategoryName');
  if (modalCatName) modalCatName.textContent = `"${activeCat.name}" (Google Sheet Tab)`;

  document.getElementById('fieldNameInput').value = '';
  document.getElementById('fieldTypeSelect').value = 'text';
  document.getElementById('fieldOptionsInput').value = '';
  document.getElementById('fieldRequiredCheck').checked = false;
  document.getElementById('fieldOptionsGroup').style.display = 'none';

  openModal('fieldModal');
}

// Handle Add Custom Field Form Submit
function handleFieldFormSubmit(e) {
  e.preventDefault();
  const activeCat = getActiveCategory();
  if (!activeCat) return;

  const label = document.getElementById('fieldNameInput').value.trim();
  const type = document.getElementById('fieldTypeSelect').value;
  const rawOptions = document.getElementById('fieldOptionsInput').value;
  const required = document.getElementById('fieldRequiredCheck').checked;

  if (!label) return;

  const options = type === 'select'
    ? rawOptions.split(',').map(s => s.trim()).filter(s => s.length > 0)
    : [];

  try {
    Storage.addFieldToCategory(activeCat.id, {
      label: label,
      type: type,
      options: options,
      required: required
    });

    state.categories = Storage.getCategories();
    renderActiveCategoryToolbar();
    renderAssets();
    closeModal('fieldModal');

    showToast(`Column "${label}" added to tab "${activeCat.name}"!`);
  } catch (err) {
    showToast(err.message, 'error');
  }
}

// Open Category Fields Manager Modal
function openCategoryFieldsModal() {
  const activeCat = getActiveCategory();
  if (!activeCat) return;

  const catNameEl = document.getElementById('catFieldsModalName');
  const listEl = document.getElementById('modalFieldsList');

  if (catNameEl) catNameEl.textContent = `"${activeCat.name}"`;

  listEl.innerHTML = activeCat.fields.map(field => `
    <div class="field-item-row">
      <div class="field-info">
        <i data-lucide="${field.isDefault ? 'shield-check' : 'tag'}" style="width: 18px; height: 18px; color: ${field.isDefault ? 'var(--accent-emerald)' : 'var(--text-secondary)'}"></i>
        <div>
          <div style="font-weight: 600; font-size: 0.9rem; color: #fff;">${field.label}</div>
          <span class="field-type-badge">${field.type.toUpperCase()}${field.required ? ' • Required' : ''}</span>
        </div>
      </div>
      <div>
        ${!field.isDefault ? `
          <button class="btn-card-action delete btn-remove-field" data-field-id="${field.id}" title="Remove Field">
            <i data-lucide="trash-2" style="width: 15px; height: 15px;"></i>
          </button>
        ` : `<span style="font-size: 0.72rem; color: var(--text-muted);">Core Tab Column</span>`}
      </div>
    </div>
  `).join('');

  listEl.querySelectorAll('.btn-remove-field').forEach(btn => {
    btn.addEventListener('click', () => {
      if (confirm(`Remove field column?`)) {
        Storage.deleteFieldFromCategory(activeCat.id, btn.dataset.fieldId);
        state.categories = Storage.getCategories();
        openCategoryFieldsModal();
        renderActiveCategoryToolbar();
        renderAssets();
        showToast('Field column removed.');
      }
    });
  });

  refreshIcons();
  openModal('categoryFieldsModal');
}

// Handle Add New Category (New Tab)
function handleCategoryFormSubmit(e) {
  e.preventDefault();
  const name = document.getElementById('categoryNameInput').value.trim();
  const desc = document.getElementById('categoryDescInput').value.trim();
  const color = document.getElementById('categoryColorInput').value;

  if (!name) return;

  try {
    const newCat = Storage.addCategory({
      name: name,
      description: desc,
      color: color,
      icon: 'Folder'
    });

    state.categories = Storage.getCategories();
    state.activeCategoryId = newCat.id;

    renderStatsBanner();
    renderCategoryChips();
    renderActiveCategoryToolbar();
    renderAssets();
    closeModal('categoryModal');

    showToast(`New Category Tab "${newCat.name}" created! Ready in Google Sheet.`);
  } catch (err) {
    showToast(err.message, 'error');
  }
}

// Handle Excel Template File Selection
async function handleExcelFileSelected(file) {
  if (!file) return;

  try {
    showToast('Parsing Excel template tabs...', 'info');
    const parsedData = await ExcelHandler.parseTemplateFile(file);
    state.pendingExcelImport = parsedData;

    // Show Preview in Modal
    const previewArea = document.getElementById('excelPreviewArea');
    const previewSummary = document.getElementById('previewTabsSummary');
    const previewList = document.getElementById('previewTabsList');

    if (parsedData.isStandardBulkTemplate) {
      previewSummary.textContent = `OREL Bulk Template: ${parsedData.totalRows} Assets in ${parsedData.tabs.length} Categories`;
    } else {
      previewSummary.textContent = `${parsedData.tabs.length} Tabs Detected in "${parsedData.fileName}"`;
    }

    let listHtml = '';
    if (parsedData.isStandardBulkTemplate) {
      listHtml += `
        <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: var(--radius-sm); padding: 0.65rem 0.85rem; margin-bottom: 0.5rem;">
          <div style="font-size: 0.8rem; font-weight: 700; color: var(--accent-emerald-light);">✓ Official OREL 23-Column Schema Verified</div>
          <div style="font-size: 0.72rem; color: var(--text-secondary); margin-top: 2px;">Asset Name, Asset Number, Title, Asset Type, Category, Brand, Model, Serial No, Barcode, QR Code, Company, Department, Location, Purchase Price, Warranty, Depreciation...</div>
        </div>
      `;
    }

    listHtml += parsedData.tabs.map(tab => `
      <div style="background: var(--bg-surface); padding: 0.65rem 0.85rem; border-radius: var(--radius-sm); border: 1px solid var(--border-subtle); display: flex; align-items: center; justify-content: space-between;">
        <div>
          <strong style="color: #fff; font-size: 0.88rem;">Category Tab: "${tab.tabName}"</strong>
          <div style="font-size: 0.75rem; color: var(--text-secondary); margin-top: 2px;">
            ${tab.headers.length} Columns: ${tab.headers.slice(0, 5).join(', ')}${tab.headers.length > 5 ? '...' : ''}
          </div>
        </div>
        <span class="sync-badge" style="font-size: 0.75rem; font-weight: 700; color: var(--accent-emerald-light);">${tab.rows.length} Assets</span>
      </div>
    `).join('');

    previewList.innerHTML = listHtml;
    previewArea.style.display = 'block';
    refreshIcons();
  } catch (err) {
    showToast(err.message, 'error');
  }
}

// Confirm and Apply Excel Template Import
function confirmExcelTemplateImport() {
  if (!state.pendingExcelImport) return;

  const { tabs, isStandardBulkTemplate } = state.pendingExcelImport;
  let importedTabsCount = 0;
  let importedAssetsCount = 0;

  tabs.forEach(tab => {
    // Check if category exists
    let existingCat = state.categories.find(c => c.name.toLowerCase() === tab.tabName.toLowerCase());

    if (!existingCat) {
      existingCat = Storage.addCategory({
        name: tab.tabName,
        fields: JSON.parse(JSON.stringify(STANDARD_TEMPLATE_FIELDS))
      });
      importedTabsCount++;
    } else {
      // Merge new custom fields if any
      tab.headers.forEach(h => {
        const fId = h.toLowerCase().replace(/[^a-z0-9]/g, '_');
        if (!existingCat.fields.some(f => f.id === fId || f.label.toLowerCase() === h.toLowerCase())) {
          existingCat.fields.push({
            id: fId,
            label: h,
            type: 'text',
            required: false,
            isDefault: false,
            group: 'Custom Fields'
          });
        }
      });
      Storage.saveCategories(state.categories);
    }

    // Import asset rows
    tab.rows.forEach(row => {
      const assetData = {
        ...row,
        categoryId: existingCat.id,
        categoryName: existingCat.name
      };
      Storage.addAsset(assetData);
      importedAssetsCount++;
    });
  });

  state.categories = Storage.getCategories();
  state.assets = Storage.getAssets();

  document.getElementById('excelPreviewArea').style.display = 'none';
  state.pendingExcelImport = null;
  closeModal('templateModal');

  renderStatsBanner();
  renderCategoryChips();
  renderActiveCategoryToolbar();
  renderAssets();

  showToast(`Excel Imported! ${importedAssetsCount} assets loaded across ${tabs.length} categories.`);
}

// Google Sheets Sync Functions
async function testGoogleSheetsConnection() {
  const urlInput = document.getElementById('googleScriptUrlInput');
  const feedback = document.getElementById('connectionTestFeedback');
  const url = urlInput.value.trim();

  if (!url) {
    feedback.textContent = 'Please paste your Google Apps Script Web App URL first.';
    feedback.style.color = 'var(--accent-rose)';
    return;
  }

  feedback.textContent = 'Testing connection to Google Sheet...';
  feedback.style.color = 'var(--accent-amber)';

  try {
    const res = await GoogleSheets.testConnection(url);
    feedback.textContent = `Connected! ${res.message}`;
    feedback.style.color = 'var(--accent-emerald-light)';

    state.settings.googleAppsScriptUrl = url;
    Storage.saveSettings(state.settings);
    updateSyncStatusDisplay();
    showToast('Google Sheets Backend verified & saved!');
  } catch (err) {
    feedback.textContent = err.message;
    feedback.style.color = 'var(--accent-rose)';
  }
}

async function pushAllDataToGoogleSheets() {
  const url = state.settings.googleAppsScriptUrl || document.getElementById('googleScriptUrlInput').value.trim();
  if (!url) {
    showToast('Please enter your Google Apps Script Web App URL in settings first!', 'error');
    openModal('googleSheetsModal');
    return;
  }

  const syncDot = document.getElementById('syncDot');
  if (syncDot) syncDot.className = 'sync-dot syncing';
  showToast('Pushing all categories to Google Sheet tabs...', 'info');

  try {
    const res = await GoogleSheets.pushAllToGoogleSheet(url, state.categories, state.assets);
    showToast(`Google Sheets Updated! ${res.message || 'All tabs synced.'}`, 'success');
  } catch (err) {
    showToast(`Push failed: ${err.message}`, 'error');
  } finally {
    updateSyncStatusDisplay();
  }
}

async function pullAllDataFromGoogleSheets() {
  const url = state.settings.googleAppsScriptUrl || document.getElementById('googleScriptUrlInput').value.trim();
  if (!url) {
    showToast('Please enter your Google Apps Script Web App URL first!', 'error');
    openModal('googleSheetsModal');
    return;
  }

  showToast('Pulling data from Google Sheet tabs...', 'info');

  try {
    const res = await GoogleSheets.fetchRemoteData(url);
    if (res.status === 'success' && res.tabs) {
      // Process tabs into categories
      res.tabs.forEach(tab => {
        if (tab.tabName === 'Asset Upload') return;
        let cat = state.categories.find(c => c.name.toLowerCase() === tab.tabName.toLowerCase());
        const headers = tab.headers || [];
        const fields = headers.map(h => ({
          id: h.toLowerCase().replace(/[^a-z0-9]/g, '_'),
          label: h,
          type: 'text',
          required: false,
          isDefault: true
        }));

        if (!cat && headers.length > 0) {
          Storage.addCategory({
            name: tab.tabName,
            fields: fields
          });
        }
      });

      // Update assets
      if (Array.isArray(res.assets)) {
        state.assets = res.assets.map(a => {
          const cat = state.categories.find(c => c.name.toLowerCase() === (a.categoryName || '').toLowerCase());
          const fields = cat?.fields || STANDARD_TEMPLATE_FIELDS;
          const normalized = {};
          fields.forEach(field => { if (a[field.label] !== undefined) normalized[field.id] = a[field.label]; });
          const number = a['Asset Number'] || a['Asset ID'] || a.asset_id || a.id;
          return {
            ...a,
            ...normalized,
            id: number,
            asset_number: number,
            asset_id: number,
            _revision: a.Revision || '',
            categoryId: cat ? cat.id : 'general',
            categoryName: a.categoryName || (cat ? cat.name : 'General')
          };
        });
        Storage.saveAssets(state.assets);
      }

      state.categories = Storage.getCategories();
      renderStatsBanner();
      renderCategoryChips();
      renderActiveCategoryToolbar();
      renderAssets();

      showToast(`Pulled ${res.assets ? res.assets.length : 0} assets from Google Sheet!`);
    }
  } catch (err) {
    showToast(`Pull failed: ${err.message}`, 'error');
  }
}

// Modal open/close helpers
function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.add('open');
    refreshIcons();
  }
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.remove('open');
  }
}

// Setup Event Listeners
function setupEventListeners() {
  window.addEventListener('storage', (event) => {
    if (event.key !== 'orel_asset_items_v4') return;
    state.assets = Storage.getAssets();
    renderStatsBanner();
    renderCategoryChips();
    renderAssets();
  });
  // Modal close triggers
  document.querySelectorAll('[data-close-modal]').forEach(btn => {
    btn.addEventListener('click', () => {
      closeModal(btn.dataset.closeModal);
    });
  });

  // Modal backdrop click to close
  document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
    backdrop.addEventListener('click', (e) => {
      if (e.target === backdrop) {
        backdrop.classList.remove('open');
      }
    });
  });

  // Search input
  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      renderAssets();
    });
  }

  // Asset Type Filter
  const assetTypeFilter = document.getElementById('assetTypeFilter');
  if (assetTypeFilter) {
    assetTypeFilter.addEventListener('change', (e) => {
      state.assetTypeFilter = e.target.value;
      renderAssets();
    });
  }

  // Status Filter
  const statusFilter = document.getElementById('statusFilter');
  if (statusFilter) {
    statusFilter.addEventListener('change', (e) => {
      state.statusFilter = e.target.value;
      renderAssets();
    });
  }

  // Sort By
  const sortBySelect = document.getElementById('sortBySelect');
  if (sortBySelect) {
    sortBySelect.addEventListener('change', (e) => {
      state.sortBy = e.target.value;
      renderAssets();
    });
  }

  // View Mode Toggles
  const btnCards = document.getElementById('btnViewCards');
  const btnTable = document.getElementById('btnViewTable');

  if (btnCards && btnTable) {
    btnCards.addEventListener('click', () => {
      state.currentView = 'cards';
      btnCards.classList.add('active');
      btnTable.classList.remove('active');
      renderAssets();
    });
    btnTable.addEventListener('click', () => {
      state.currentView = 'table';
      btnTable.classList.add('active');
      btnCards.classList.remove('active');
      renderAssets();
    });
  }

  // Top Action Buttons
  const btnOpenAddAsset = document.getElementById('btnOpenAddAsset');
  const fabAddAsset = document.getElementById('fabAddAsset');
  if (btnOpenAddAsset) btnOpenAddAsset.addEventListener('click', () => openAssetModal());
  if (fabAddAsset) fabAddAsset.addEventListener('click', () => openAssetModal());

  const btnOpenSync = document.getElementById('btnOpenSync');
  if (btnOpenSync) btnOpenSync.addEventListener('click', () => openModal('googleSheetsModal'));

  const btnQuickSync = document.getElementById('btnQuickSync');
  if (btnQuickSync) btnQuickSync.addEventListener('click', () => pushAllDataToGoogleSheets());

  const btnOpenTemplate = document.getElementById('btnOpenTemplate');
  if (btnOpenTemplate) btnOpenTemplate.addEventListener('click', () => openModal('templateModal'));

  const btnDownloadTemplateQuick = document.getElementById('btnDownloadTemplateQuick');
  if (btnDownloadTemplateQuick) {
    btnDownloadTemplateQuick.addEventListener('click', () => {
      ExcelHandler.exportStandardTemplateWorkbook(state.assets, false);
      showToast('Official OREL Bulk Upload Template downloaded!');
    });
  }

  const btnOpenUploadModalQuick = document.getElementById('btnOpenUploadModalQuick');
  if (btnOpenUploadModalQuick) {
    btnOpenUploadModalQuick.addEventListener('click', () => openModal('templateModal'));
  }

  // Category Actions
  const btnAddCustomField = document.getElementById('btnAddCustomField');
  if (btnAddCustomField) btnAddCustomField.addEventListener('click', () => openAddCustomFieldModal());

  const btnManageCategoryFields = document.getElementById('btnManageCategoryFields');
  if (btnManageCategoryFields) btnManageCategoryFields.addEventListener('click', () => openCategoryFieldsModal());

  const btnOpenAddFieldFromManager = document.getElementById('btnOpenAddFieldFromManager');
  if (btnOpenAddFieldFromManager) {
    btnOpenAddFieldFromManager.addEventListener('click', () => {
      closeModal('categoryFieldsModal');
      openAddCustomFieldModal();
    });
  }

  // Form Submissions
  const assetForm = document.getElementById('assetForm');
  if (assetForm) assetForm.addEventListener('submit', handleAssetFormSubmit);

  const fieldForm = document.getElementById('fieldForm');
  if (fieldForm) fieldForm.addEventListener('submit', handleFieldFormSubmit);

  const fieldTypeSelect = document.getElementById('fieldTypeSelect');
  if (fieldTypeSelect) {
    fieldTypeSelect.addEventListener('change', (e) => {
      const optsGroup = document.getElementById('fieldOptionsGroup');
      if (optsGroup) optsGroup.style.display = e.target.value === 'select' ? 'block' : 'none';
    });
  }

  const categoryForm = document.getElementById('categoryForm');
  if (categoryForm) categoryForm.addEventListener('submit', handleCategoryFormSubmit);

  const colorInput = document.getElementById('categoryColorInput');
  const colorText = document.getElementById('categoryColorText');
  if (colorInput && colorText) {
    colorInput.addEventListener('input', (e) => {
      colorText.value = e.target.value;
    });
  }

  // Excel Dropzone & File Input
  const dropzone = document.getElementById('excelDropzone');
  const fileInput = document.getElementById('excelFileInput');

  if (dropzone && fileInput) {
    dropzone.addEventListener('click', () => fileInput.click());

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleExcelFileSelected(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleExcelFileSelected(e.target.files[0]);
      }
    });
  }

  const btnConfirmExcelImport = document.getElementById('btnConfirmExcelImport');
  if (btnConfirmExcelImport) btnConfirmExcelImport.addEventListener('click', confirmExcelTemplateImport);

  const btnCancelExcelImport = document.getElementById('btnCancelExcelImport');
  if (btnCancelExcelImport) {
    btnCancelExcelImport.addEventListener('click', () => {
      document.getElementById('excelPreviewArea').style.display = 'none';
      state.pendingExcelImport = null;
    });
  }

  const btnDownloadStandardTemplate = document.getElementById('btnDownloadStandardTemplate');
  if (btnDownloadStandardTemplate) {
    btnDownloadStandardTemplate.addEventListener('click', () => {
      ExcelHandler.exportStandardTemplateWorkbook(state.assets, false);
      showToast('Official OREL Bulk Upload Template downloaded!');
    });
  }

  const btnDownloadCleanTemplate = document.getElementById('btnDownloadCleanTemplate');
  if (btnDownloadCleanTemplate) {
    btnDownloadCleanTemplate.addEventListener('click', () => {
      ExcelHandler.exportTemplateWorkbook(state.categories, [], false);
      showToast('Multi-tab template downloaded!');
    });
  }

  const btnExportFullAssets = document.getElementById('btnExportFullAssets');
  if (btnExportFullAssets) {
    btnExportFullAssets.addEventListener('click', () => {
      ExcelHandler.exportStandardTemplateWorkbook(state.assets, true);
      showToast('Full asset registry exported to Official Template!');
    });
  }

  // Google Sheets Modal Buttons
  const btnTestConnection = document.getElementById('btnTestGoogleConnection');
  if (btnTestConnection) btnTestConnection.addEventListener('click', testGoogleSheetsConnection);

  const btnPushAll = document.getElementById('btnPushAllToSheets');
  if (btnPushAll) btnPushAll.addEventListener('click', pushAllDataToGoogleSheets);

  const btnPullAll = document.getElementById('btnPullAllFromSheets');
  if (btnPullAll) btnPullAll.addEventListener('click', pullAllDataFromGoogleSheets);

  const btnCopyScript = document.getElementById('btnCopyScriptCode');
  if (btnCopyScript) {
    btnCopyScript.addEventListener('click', () => {
      navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_CODE);
      showToast('Apps Script code copied to clipboard!');
    });
  }

  const btnSaveSettings = document.getElementById('btnSaveGoogleSettings');
  if (btnSaveSettings) {
    btnSaveSettings.addEventListener('click', () => {
      const url = document.getElementById('googleScriptUrlInput').value.trim();
      state.settings.googleAppsScriptUrl = url;
      Storage.saveSettings(state.settings);
      updateSyncStatusDisplay();
      closeModal('googleSheetsModal');
      showToast('Settings saved!');
    });
  }

  // Asset Tag Print Button
  const btnPrintTag = document.getElementById('btnPrintAssetTag');
  if (btnPrintTag) {
    btnPrintTag.addEventListener('click', () => {
      window.print();
    });
  }

  const btnEditFromDetail = document.getElementById('btnEditFromDetail');
  if (btnEditFromDetail) {
    btnEditFromDetail.addEventListener('click', () => {
      if (state.selectedAssetForDetail) {
        closeModal('assetDetailModal');
        openAssetModal(state.selectedAssetForDetail.id);
      }
    });
  }

  // Mobile Bottom Navigation Bar Buttons
  document.querySelectorAll('.bottom-nav .nav-item').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.bottom-nav .nav-item').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const navType = btn.dataset.nav;
      if (navType === 'categories') {
        const catSection = document.querySelector('.category-section');
        if (catSection) catSection.scrollIntoView({ behavior: 'smooth' });
      } else if (navType === 'template') {
        openModal('templateModal');
      } else if (navType === 'sync') {
        openModal('googleSheetsModal');
      } else {
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }
    });
  });
}

// Boot application when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}
