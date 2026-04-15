/**
 * ============================================================
 *  15B to BRZ Inventory — Main Application Logic
 * ============================================================
 * Connects to Google Sheets via Apps Script Web App API.
 * Handles CRUD operations, search, and UI interactions.
 */

// ======================== CONFIG ========================
const CONFIG = {
    SCRIPT_URL: localStorage.getItem('scriptUrl') || '',
    DEMO_MODE: false,
    REFRESH_INTERVAL: null // set to ms for auto-refresh, or null
};

// Sheet metadata
const SHEET_META = {
    all_inventory: {
        name: 'ALL INVENTORY',
        headers: ['CODE', 'NAME OF ITEM', 'PLACE OF ITEM', 'DESCRIPTION', 'UPDATE', 'PICTURES'],
        tableId: 'tableAllInventory',
        emptyId: 'emptyAllInventory',
        icon: 'inventory_2',
        formFields: [
            { key: 'CODE', label: 'Item Code', type: 'text', placeholder: 'e.g. A-1, B-5', required: true },
            { key: 'NAME OF ITEM', label: 'Name of Item', type: 'text', placeholder: 'Enter item name', required: true },
            { key: 'PLACE OF ITEM', label: 'Place of Item', type: 'text', placeholder: 'Where is the item stored?' },
            { key: 'DESCRIPTION', label: 'Description', type: 'textarea', placeholder: 'Describe the item...' },
            { key: 'UPDATE', label: 'Status', type: 'select', options: ['HINDI PA NAKA ALIS', 'NAKA ALIS NA'] },
            { key: 'PICTURES', label: 'Picture', type: 'image' }
        ]
    },
    figurine_list: {
        name: 'FIGURINE LIST',
        headers: ['FIGURINE', 'FIGURINE PLACE', 'FIGURINE UPDATE', 'PICTURES', 'DATE DEPARTURE'],
        tableId: 'tableFigurineList',
        emptyId: 'emptyFigurineList',
        icon: 'emoji_objects',
        formFields: [
            { key: 'FIGURINE', label: 'Figurine Code', type: 'text', placeholder: 'e.g. F-1, F-2', required: true },
            { key: 'FIGURINE PLACE', label: 'Figurine Place', type: 'text', placeholder: 'Where is the figurine?' },
            { key: 'FIGURINE UPDATE', label: 'Status', type: 'select', options: ['HINDI PA NAKA ALIS', 'NAKA ALIS NA'] },
            { key: 'PICTURES', label: 'Picture', type: 'image' },
            { key: 'DATE DEPARTURE', label: 'Date Departure', type: 'date' }
        ]
    },
    item_from_warehouse: {
        name: 'ITEM FROM WAREHOUSE',
        headers: ['# NUMBER OF ITEM', 'ITEM NAME', 'COLOR OF ITEM', 'DEPARTURE DATE', 'DESTINATION PLACE', 'PICTURES'],
        tableId: 'tableItemFromWarehouse',
        emptyId: 'emptyItemFromWarehouse',
        icon: 'local_shipping',
        formFields: [
            { key: '# NUMBER OF ITEM', label: 'Item Number', type: 'text', placeholder: 'e.g. 1, 2, 3', required: true },
            { key: 'ITEM NAME', label: 'Item Name', type: 'text', placeholder: 'Enter item name', required: true },
            { key: 'COLOR OF ITEM', label: 'Status', type: 'select', options: ['GOOD', 'DISPOSED', 'JUNKSHOP', 'DONATION/CHARITY', 'PENDING'] },
            { key: 'DEPARTURE DATE', label: 'Departure Date', type: 'date' },
            { key: 'DESTINATION PLACE', label: 'Destination Place', type: 'select', options: ['Breezenia', 'White Lilly', 'Corinthians', 'Junkshop'] },
            { key: 'PICTURES', label: 'Picture', type: 'image' }
        ]
    },
    all_box: {
        name: 'DATA INFORMATION (BOX)',
        headers: ['BOX NUMBER', 'BOX NAME', 'BOX PLACE', 'BOX DESCRIPTION', 'BOX UPDATE', 'PICTURES', 'DATE DEPARTURE', 'COMPANY'],
        tableId: 'tableAllBox',
        emptyId: 'emptyAllBox',
        icon: 'inbox',
        formFields: [
            { key: 'BOX NUMBER', label: 'Box Number', type: 'text', placeholder: 'e.g. B-1, B-2', required: true },
            { key: 'BOX NAME', label: 'Box Name', type: 'text', placeholder: 'Enter box name', required: true },
            { key: 'BOX PLACE', label: 'Box Place', type: 'text', placeholder: 'Where is the box?' },
            { key: 'BOX DESCRIPTION', label: 'Box Description', type: 'text', placeholder: 'e.g. BLUE TOP BOX' },
            { key: 'BOX UPDATE', label: 'Status', type: 'select', options: ['HINDI PA NAKA ALIS', 'NAKA ALIS NA'] },
            { key: 'PICTURES', label: 'Picture', type: 'image' },
            { key: 'DATE DEPARTURE', label: 'Date Departure', type: 'date' },
            { key: 'COMPANY', label: 'Company', type: 'text', placeholder: 'Company name' }
        ]
    },
    lipat_bahay: {
        name: 'LIPAT BAHAY',
        headers: ['ITEM/BOX NAME', 'ORIGIN ROOM', 'DESTINATION ROOM', 'STATUS', 'NOTES', 'PICTURES'],
        tableId: 'tableLipatBahay',
        emptyId: 'emptyLipatBahay',
        icon: 'rv_hookup',
        formFields: [
            { key: 'ITEM/BOX NAME', label: 'Item/Box Name', type: 'text', placeholder: 'e.g. Living Room TV Box', required: true },
            { key: 'ORIGIN ROOM', label: 'Origin Room', type: 'text', placeholder: 'Current location' },
            { key: 'DESTINATION ROOM', label: 'Destination Room', type: 'text', placeholder: 'Where it should go' },
            { key: 'STATUS', label: 'Status', type: 'select', options: ['Packed', 'In Transit', 'Delivered'] },
            { key: 'NOTES', label: 'Notes', type: 'textarea', placeholder: 'Special instructions...' },
            { key: 'PICTURES', label: 'Picture', type: 'image' }
        ]
    }
};

// ======================== DEMO DATA ========================
const DEMO_DATA = {
    all_inventory: [
        { _rowIndex: 2, 'CODE': 'A-1', 'NAME OF ITEM': 'BUSINESS PERMIT 2024', 'PLACE OF ITEM': 'CABINET A - 2ND FLOOR', 'DESCRIPTION': 'Original copy of business permit', 'UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '' },
        { _rowIndex: 3, 'CODE': 'A-2', 'NAME OF ITEM': 'LEASE CONTRACT', 'PLACE OF ITEM': 'CABINET A - 2ND FLOOR', 'DESCRIPTION': 'Signed lease agreement', 'UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '' },
        { _rowIndex: 4, 'CODE': 'A-3', 'NAME OF ITEM': 'BIR DOCUMENTS', 'PLACE OF ITEM': 'ROOM 1 - 1ST FLOOR', 'DESCRIPTION': 'Tax filing documents', 'UPDATE': 'NAKA ALIS NA', 'PICTURES': '' },
        { _rowIndex: 5, 'CODE': 'A-4', 'NAME OF ITEM': 'SEC REGISTRATION', 'PLACE OF ITEM': 'CABINET B', 'DESCRIPTION': 'Corporate registration papers', 'UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '' },
        { _rowIndex: 6, 'CODE': 'A-5', 'NAME OF ITEM': 'OLD RECEIPTS (2023)', 'PLACE OF ITEM': 'STORAGE ROOM', 'DESCRIPTION': 'Bundle of old receipts', 'UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '' },
    ],
    figurine_list: [
        { _rowIndex: 2, 'FIGURINE': 'F-1', 'FIGURINE PLACE': '', 'FIGURINE UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '' },
        { _rowIndex: 3, 'FIGURINE': 'F-2', 'FIGURINE PLACE': '', 'FIGURINE UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '' },
        { _rowIndex: 4, 'FIGURINE': 'F-3', 'FIGURINE PLACE': '', 'FIGURINE UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '' },
        { _rowIndex: 5, 'FIGURINE': 'F-4', 'FIGURINE PLACE': '', 'FIGURINE UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '' },
    ],
    item_from_warehouse: [
        { _rowIndex: 2, '# NUMBER OF ITEM': '1', 'ITEM NAME': 'ANTIQUE VASE', 'COLOR OF ITEM': 'GOOD', 'DEPARTURE DATE': '', 'DESTINATION PLACE': '', 'PICTURES': '' },
        { _rowIndex: 3, '# NUMBER OF ITEM': '2', 'ITEM NAME': 'WOODEN SHELF', 'COLOR OF ITEM': 'DISPOSED', 'DEPARTURE DATE': 'APRIL 10, 2026', 'DESTINATION PLACE': 'Junkshop', 'PICTURES': '' },
        { _rowIndex: 4, '# NUMBER OF ITEM': '3', 'ITEM NAME': 'CERAMIC PLATE SET', 'COLOR OF ITEM': 'DONATION/CHARITY', 'DEPARTURE DATE': 'APRIL 05, 2026', 'DESTINATION PLACE': 'White Lilly', 'PICTURES': '' },
    ],
    all_box: [
        { _rowIndex: 2, 'BOX NUMBER': 'B-1', 'BOX NAME': 'GLD-TO-DO', 'BOX PLACE': '', 'BOX DESCRIPTION': 'BLUE TOP BOX', 'BOX UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '', 'COMPANY': '' },
        { _rowIndex: 3, 'BOX NUMBER': 'B-2', 'BOX NAME': 'W.P. (B.I.R)', 'BOX PLACE': '', 'BOX DESCRIPTION': 'LIGHT BLUE TOP BOX', 'BOX UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '', 'COMPANY': '' },
        { _rowIndex: 4, 'BOX NUMBER': 'B-3', 'BOX NAME': 'ZENAIDA SERRANO OLD FILES', 'BOX PLACE': '', 'BOX DESCRIPTION': 'BLUE TOP BOX', 'BOX UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '', 'COMPANY': '' },
        { _rowIndex: 5, 'BOX NUMBER': 'B-4', 'BOX NAME': 'BUILDING INSURANCE (TZT/WP)', 'BOX PLACE': '', 'BOX DESCRIPTION': 'GREEN TOP BOX', 'BOX UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '', 'COMPANY': '' },
        { _rowIndex: 6, 'BOX NUMBER': 'B-5', 'BOX NAME': 'SEC REGISTRATION / LEASE CONTRACT', 'BOX PLACE': '', 'BOX DESCRIPTION': 'BROWN BOX', 'BOX UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '', 'COMPANY': '' },
        { _rowIndex: 7, 'BOX NUMBER': 'B-6', 'BOX NAME': 'PLANO 1', 'BOX PLACE': '', 'BOX DESCRIPTION': 'BLACK BOX', 'BOX UPDATE': 'HINDI PA NAKA ALIS', 'PICTURES': '', 'DATE DEPARTURE': '', 'COMPANY': '' },
    ],
    lipat_bahay: [
        { _rowIndex: 2, 'ITEM/BOX NAME': 'Kitchen Equipment Phase 1', 'ORIGIN ROOM': 'Kitchen DB', 'DESTINATION ROOM': 'Kitchen Setup', 'STATUS': 'Packed', 'NOTES': 'Fragile plates inside', 'PICTURES': '' },
        { _rowIndex: 3, 'ITEM/BOX NAME': 'Executive Desks', 'ORIGIN ROOM': 'CEO Office', 'DESTINATION ROOM': 'Storage Unit 5', 'STATUS': 'In Transit', 'NOTES': 'Requires 3 men to lift', 'PICTURES': '' },
    ]
};

// ======================== STATE ========================
let currentTab = 'dashboard';
let cachedData = {};
let deleteTarget = null;
let pendingImageUploads = {}; // key -> base64 data for images to upload

// ======================== INIT & AUTH ========================
document.addEventListener('DOMContentLoaded', () => {
    const savedPin = sessionStorage.getItem('authPin');
    if (!savedPin) {
        // No PIN saved, show login modal
        document.getElementById('app').style.display = 'none';
        setupLogin();
    } else {
        // Assume valid, load app. Backend will reject if changed.
        document.getElementById('loginOverlay').style.display = 'none';
        document.getElementById('app').style.display = 'block';
        initApp();
    }
});

function setupLogin() {
    const btnSubmit = document.getElementById('btnLoginSubmit');
    const inputPin = document.getElementById('authPinInput');
    const errorMsg = document.getElementById('loginErrorMsg');

    const handleLogin = async () => {
        const pin = inputPin.value.trim();
        if (!pin) return;
        
        btnSubmit.disabled = true;
        btnSubmit.innerHTML = '<div class="spinner" style="width:16px;height:16px;border-width:2px;display:inline-block;"></div>';
        errorMsg.style.display = 'none';

        try {
            if (CONFIG.DEMO_MODE) {
                // In demo mode, bypass actual backend
                sessionStorage.setItem('authPin', pin);
                document.getElementById('loginOverlay').style.display = 'none';
                document.getElementById('app').style.display = 'block';
                initApp();
            } else {
                // Not in demo mode, authenticate with backend
                if (!CONFIG.SCRIPT_URL) {
                    // Force them to put generic app mode first, or just accept if URL missing?
                    // Let's just bypass to SetupModal if no URL is set at all.
                    sessionStorage.setItem('authPin', pin);
                    document.getElementById('loginOverlay').style.display = 'none';
                    document.getElementById('app').style.display = 'block';
                    initApp();
                    return;
                }

                const url = `${CONFIG.SCRIPT_URL}?action=auth_check&auth=${encodeURIComponent(pin)}`;
                const response = await fetch(url);
                const result = await response.json();

                if (result.success) {
                    sessionStorage.setItem('authPin', pin);
                    document.getElementById('loginOverlay').style.display = 'none';
                    document.getElementById('app').style.display = 'block';
                    initApp();
                } else {
                    errorMsg.style.display = 'block';
                    btnSubmit.disabled = false;
                    btnSubmit.innerHTML = '<span class="material-icons-round">login</span>';
                }
            }
        } catch (e) {
            errorMsg.textContent = "Network Error. Check connection.";
            errorMsg.style.display = 'block';
            btnSubmit.disabled = false;
            btnSubmit.innerHTML = '<span class="material-icons-round">login</span>';
        }
    };

    btnSubmit.addEventListener('click', handleLogin);
    inputPin.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleLogin();
    });
}

function initApp() {
    setupThemeToggle();
    setupDate();
    setupNavigation();
    setupSearch();
    setupModals();
    setupQuickActions();
    setupSidebarToggle();
    setupMobileNav();
    checkConnection();
    setupDatabaseManager();

    // Brand logo refresh
    const brandRefresh = document.getElementById('brandRefresh');
    if (brandRefresh) {
        brandRefresh.addEventListener('click', () => {
            if (CONFIG.SCRIPT_URL && !CONFIG.DEMO_MODE) {
                showToast('Refreshing data...', 'info');
                loadAllData();
            } else {
                showToast('Cannot refresh in Demo Mode or without Connection', 'warning');
            }
        });
    }
}

// ======================== THEME TOGGLE ========================
function setupThemeToggle() {
    const toggle = document.getElementById('themeToggle');
    const icon = document.getElementById('themeIcon');
    const savedTheme = localStorage.getItem('theme') || 'dark';

    // Apply saved theme on load
    applyTheme(savedTheme);

    toggle.addEventListener('click', () => {
        const current = document.documentElement.getAttribute('data-theme') || 'dark';
        const next = current === 'dark' ? 'light' : 'dark';
        applyTheme(next);
        localStorage.setItem('theme', next);
    });

    function applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        if (theme === 'light') {
            icon.textContent = 'dark_mode';
            toggle.title = 'Switch to Dark Mode';
        } else {
            icon.textContent = 'light_mode';
            toggle.title = 'Switch to Light Mode';
        }
    }
}

// ======================== CONNECTION ========================
function checkConnection() {
    if (CONFIG.SCRIPT_URL) {
        CONFIG.DEMO_MODE = false;
        setConnectionStatus(true);
        loadAllData();
    } else {
        // Show setup modal
        document.getElementById('setupModal').style.display = 'flex';
        document.getElementById('loadingOverlay').classList.add('hidden');

        document.getElementById('btnSaveUrl').addEventListener('click', () => {
            const url = document.getElementById('scriptUrlInput').value.trim();
            if (url && url.startsWith('https://script.google.com')) {
                localStorage.setItem('scriptUrl', url);
                CONFIG.SCRIPT_URL = url;
                CONFIG.DEMO_MODE = false;
                document.getElementById('setupModal').style.display = 'none';
                setConnectionStatus(true);
                showLoading();
                loadAllData();
            } else {
                showToast('Please enter a valid Google Apps Script URL', 'error');
            }
        });

        document.getElementById('btnDemoMode').addEventListener('click', () => {
            CONFIG.DEMO_MODE = true;
            document.getElementById('setupModal').style.display = 'none';
            setConnectionStatus(false);
            loadDemoData();
        });
    }
}

function setConnectionStatus(connected) {
    const el = document.getElementById('connectionStatus');
    const text = el.querySelector('.status-text');
    if (connected) {
        el.classList.add('connected');
        text.textContent = 'Connected';
    } else {
        el.classList.remove('connected');
        text.textContent = 'Demo';
    }
}

// ======================== DATE ========================
function setupDate() {
    const dateEl = document.getElementById('currentDate');
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    dateEl.textContent = now.toLocaleDateString('en-US', options);
}

// ======================== NAVIGATION ========================
function setupNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const tab = item.dataset.tab;
            switchTab(tab);
        });
    });
}

function switchTab(tab) {
    currentTab = tab;

    // Update nav
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    const activeNav = document.querySelector(`.nav-item[data-tab="${tab}"]`);
    if (activeNav) activeNav.classList.add('active');

    // Update content
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));

    const tabMap = {
        dashboard: 'tabDashboard',
        all_inventory: 'tabAllInventory',
        figurine_list: 'tabFigurineList',
        item_from_warehouse: 'tabItemFromWarehouse',
        all_box: 'tabAllBox',
        lipat_bahay: 'tabLipatBahay'
    };

    const targetTab = document.getElementById(tabMap[tab]);
    if (targetTab) targetTab.classList.add('active');

    // If switching to All Inventory, trigger the master view render
    if (tab === 'all_inventory') {
        renderTable('all_inventory', []); // Pass empty array as data is merged inside renderTable
    }

    // Close sidebar on mobile
    document.getElementById('sidebar').classList.remove('open');
    
    // Update mobile nav UI
    updateMobileNavStatus(tab);
}

function setupMobileNav() {
    const mobileNavItems = document.querySelectorAll('.mobile-nav-item');
    mobileNavItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const tab = item.dataset.tab;
            switchTab(tab);
        });
    });

    // Mobile Add Button (FAB)
    const btnMobileAdd = document.getElementById('btnMobileAdd');
    if (btnMobileAdd) {
        btnMobileAdd.addEventListener('click', () => {
            const sheet = currentTab === 'dashboard' ? 'all_inventory' : currentTab;
            openAddModal(sheet);
        });
    }
}

function updateMobileNavStatus(tab) {
    document.querySelectorAll('.mobile-nav-item').forEach(item => {
        if (item.dataset.tab === tab) {
            item.classList.add('active');
        } else {
            item.classList.remove('active');
        }
    });
}

function renderMasterRows(tbody, masterData) {
    tbody.innerHTML = '';
    masterData.forEach(row => {
        const tr = document.createElement('tr');
        
        // Code
        const tdCode = document.createElement('td');
        tdCode.dataset.label = 'CODE / ID';
        tdCode.textContent = row._displayCode;
        tr.appendChild(tdCode);

        // Name
        const tdName = document.createElement('td');
        tdName.dataset.label = 'NAME';
        tdName.textContent = row._displayName;
        tr.appendChild(tdName);

        // Place
        const tdPlace = document.createElement('td');
        tdPlace.dataset.label = 'LOCATION';
        tdPlace.textContent = row._displayPlace;
        tr.appendChild(tdPlace);

        // Status
        const tdStatus = document.createElement('td');
        tdStatus.dataset.label = 'STATUS';
        const statusVal = row._displayStatus;
        if (statusVal.toUpperCase().includes('HINDI') || statusVal.toUpperCase().includes('NAKA ALIS')) {
             tdStatus.innerHTML = renderStatusBadge(statusVal);
        } else {
             tdStatus.innerHTML = renderItemColor(statusVal);
        }
        tdStatus.style.cursor = 'pointer';
        tdStatus.onclick = () => openEditModal(row._sourceSheet, row._rowIndex);
        tr.appendChild(tdStatus);

        // Category
        const tdCategory = document.createElement('td');
        tdCategory.dataset.label = 'CATEGORY';
        tdCategory.innerHTML = `<span class="category-badge">${row._sourceName}</span>`;
        tr.appendChild(tdCategory);

        // Picture
        const tdPic = document.createElement('td');
        tdPic.dataset.label = 'PICTURES';
        tdPic.innerHTML = renderPicture(row['PICTURES'] || '', row, SHEET_META[row._sourceSheet]);
        tr.appendChild(tdPic);

        // Actions
        const tdActions = document.createElement('td');
        tdActions.dataset.label = 'ACTIONS';
        tdActions.innerHTML = `
            <div class="action-btns">
                <button class="btn-icon edit" title="Edit" onclick="event.stopPropagation(); openEditModal('${row._sourceSheet}', ${row._rowIndex})">
                    <span class="material-icons-round">edit</span>
                </button>
            </div>
        `;
        tr.appendChild(tdActions);
        
        tbody.appendChild(tr);
    });
}

// ======================== SIDEBAR TOGGLE ========================
function setupSidebarToggle() {
    document.getElementById('sidebarToggle').addEventListener('click', () => {
        if (window.innerWidth <= 768) {
            document.getElementById('sidebar').classList.toggle('open');
        } else {
            document.body.classList.toggle('sidebar-collapsed');
        }
    });

    // Settings button — reopen setup
    document.getElementById('btnSettings').addEventListener('click', () => {
        document.getElementById('setupModal').style.display = 'flex';
        document.getElementById('scriptUrlInput').value = CONFIG.SCRIPT_URL;
    });

    // Refresh button
    document.getElementById('btnRefresh').addEventListener('click', () => {
        if (CONFIG.DEMO_MODE) {
            showToast('Demo mode — connect to Google Sheets to refresh live data', 'info');
            return;
        }
        showToast('Refreshing data from Google Sheets...', 'info');
        loadAllData();
    });
}

// ======================== DATA LOADING ========================
async function loadAllData() {
    showLoading();
    try {
        const sheets = ['all_inventory', 'figurine_list', 'item_from_warehouse', 'all_box', 'lipat_bahay'];
        const promises = sheets.map(sheet => fetchSheetData(sheet));
        const results = await Promise.allSettled(promises);

        results.forEach((result, i) => {
            if (result.status === 'fulfilled') {
                cachedData[sheets[i]] = result.value;
                renderTable(sheets[i], result.value);
            } else {
                console.error(`Error loading ${sheets[i]}:`, result.reason);
            }
        });

        updateStats();
        hideLoading();
        showToast('Inventory data loaded successfully', 'success');
    } catch (err) {
        console.error('Load error:', err);
        hideLoading();
        showToast('Error loading data: ' + err.message, 'error');
    }
}

async function fetchSheetData(sheetKey) {
    if (CONFIG.DEMO_MODE) {
        return DEMO_DATA[sheetKey] || [];
    }

    const pin = sessionStorage.getItem('authPin') || '';
    const url = `${CONFIG.SCRIPT_URL}?action=fetch&sheet=${sheetKey}&auth=${encodeURIComponent(pin)}`;
    const response = await fetch(url);
    const json = await response.json();

    if (json.error) throw new Error(json.error);
    return json.data || [];
}

function loadDemoData() {
    Object.keys(DEMO_DATA).forEach(key => {
        cachedData[key] = DEMO_DATA[key];
        renderTable(key, DEMO_DATA[key]);
    });
    updateStats();
    hideLoading();
    showToast('Demo data loaded — connect Google Sheets to use live data', 'info');
}

// ======================== TABLE RENDERING ========================
function renderTable(sheetKey, data) {
    const meta = SHEET_META[sheetKey];
    if (!meta) return;

    const table = document.getElementById(meta.tableId);
    const tbody = table.querySelector('tbody');
    const emptyEl = document.getElementById(meta.emptyId);

    tbody.innerHTML = '';

    // Specialized Master List Rendering for All Inventory
    if (sheetKey === 'all_inventory' && currentTab === 'all_inventory') {
        const masterData = [];
        Object.keys(SHEET_META).forEach(key => {
            const sheetItems = cachedData[key] || [];
            sheetItems.forEach(item => {
                masterData.push({
                    ...item,
                    _sourceSheet: key,
                    _sourceName: SHEET_META[key].name,
                    _displayCode: item['CODE'] || item['FIGURINE'] || item['# NUMBER OF ITEM'] || item['BOX NUMBER'] || item['ITEM/BOX NAME'] || '—',
                    _displayName: item['NAME OF ITEM'] || item['ITEM NAME'] || item['BOX NAME'] || item['ITEM/BOX NAME'] || '—',
                    _displayPlace: item['PLACE OF ITEM'] || item['FIGURINE PLACE'] || item['DESTINATION PLACE'] || item['BOX PLACE'] || item['ORIGIN ROOM'] || '—',
                    _displayStatus: item['UPDATE'] || item['FIGURINE UPDATE'] || item['COLOR OF ITEM'] || item['BOX UPDATE'] || item['STATUS'] || '—'
                });
            });
        });
        renderMasterRows(tbody, masterData);
        return;
    }

    if (!data || data.length === 0) {
        table.style.display = 'none';
        emptyEl.style.display = 'block';
        return;
    }

    table.style.display = '';
    emptyEl.style.display = 'none';

    data.forEach(row => {
        const tr = document.createElement('tr');
        meta.headers.forEach(header => {
            const td = document.createElement('td');
            td.dataset.label = header;
            const value = row[header] || '';

            // Handle special columns
            if (header.includes('UPDATE') && value) {
                td.innerHTML = renderStatusBadge(value);
                td.style.cursor = 'pointer';
                td.title = "Click to edit status";
                td.onclick = () => openEditModal(sheetKey, row._rowIndex);
            } else if (header === 'PICTURES' && value) {
                td.innerHTML = renderPicture(value, row, meta);
            } else if (header === 'BOX DESCRIPTION' && value) {
                td.innerHTML = renderBoxColor(value);
            } else if (header === 'COLOR OF ITEM' && value) {
                td.innerHTML = renderItemColor(value);
            } else {
                td.textContent = value;
                td.title = value; // tooltip for truncated text
            }

            tr.appendChild(td);
        });

        // Action buttons
        const actionTd = document.createElement('td');
        actionTd.dataset.label = 'ACTIONS';
        actionTd.innerHTML = `
            <div class="action-btns">
                <button class="btn-icon edit" title="Edit" data-sheet="${sheetKey}" data-row="${row._rowIndex}">
                    <span class="material-icons-round">edit</span>
                </button>
                <button class="btn-icon delete" title="Delete" data-sheet="${sheetKey}" data-row="${row._rowIndex}" data-name="${row[meta.headers[0]] || ''}">
                    <span class="material-icons-round">delete</span>
                </button>
            </div>
        `;
        tr.appendChild(actionTd);
        tbody.appendChild(tr);
    });

    // Attach image viewer handlers
    tbody.querySelectorAll('.table-thumb').forEach(img => {
        img.addEventListener('click', () => {
            openImageViewer(img.src, img.dataset.name, img.dataset.status);
        });
    });

    // Attach action handlers
    tbody.querySelectorAll('.btn-icon.edit').forEach(btn => {
        btn.addEventListener('click', () => {
            const sheet = btn.dataset.sheet;
            const rowIndex = parseInt(btn.dataset.row);
            openEditModal(sheet, rowIndex);
        });
    });

    tbody.querySelectorAll('.btn-icon.delete').forEach(btn => {
        btn.addEventListener('click', () => {
            const sheet = btn.dataset.sheet;
            const rowIndex = parseInt(btn.dataset.row);
            const name = btn.dataset.name;
            openDeleteModal(sheet, rowIndex, name);
        });
    });
}

function renderStatusBadge(status) {
    const isPending = status.toUpperCase().includes('HINDI');
    const cls = isPending ? 'pending' : 'done';
    const icon = isPending ? 'schedule' : 'check_circle';
    const label = isPending ? 'Hindi Pa Naka Alis' : 'Naka Alis Na';
    return `<span class="status-badge ${cls}"><span class="material-icons-round">${icon}</span>${label}</span>`;
}

function renderPicture(url, row, meta) {
    if (!url || url.trim() === '') return '—';
    
    // Find name and status from row using headers
    const nameKey = meta.headers.find(h => h.includes('NAME') || h === 'FIGURINE') || meta.headers[0];
    const statusKey = meta.headers.find(h => h.includes('UPDATE') || h.includes('STATUS') || h === 'COLOR OF ITEM') || '';
    
    const name = row[nameKey] || 'Unknown Item';
    const status = statusKey ? (row[statusKey] || 'No Status') : '';

    // Create safe data attribute strings by escaping quotes
    const safeName = name.toString().replace(/"/g, '&quot;');
    const safeStatus = status.toString().replace(/"/g, '&quot;');

    return `<img src="${url}" alt="Item photo" class="table-thumb" onerror="this.style.display='none'" data-name="${safeName}" data-status="${safeStatus}" />`;
}

function renderBoxColor(description) {
    const colors = {
        'BLUE': '#3b82f6',
        'LIGHT BLUE': '#67b3f4',
        'GREEN': '#22c55e',
        'BLACK': '#374151',
        'BROWN': '#a16207',
        'RED': '#ef4444',
        'WHITE': '#e5e7eb',
        'TRANSPARENT': '#94a3b8',
        'YELLOW': '#eab308'
    };

    let dotColor = '#94a3b8';
    const upperDesc = description.toUpperCase();
    for (const [name, hex] of Object.entries(colors)) {
        if (upperDesc.includes(name)) {
            dotColor = hex;
            break;
        }
    }

    return `<span class="box-color-badge"><span class="box-color-dot" style="background:${dotColor}"></span>${description}</span>`;
}

function renderItemColor(status) {
    if (!status) return '—';
    const statusMap = {
        'GOOD': { color: '#10b981', bg: 'rgba(16, 185, 129, 0.15)', icon: 'check_circle' },
        'DISPOSED': { color: '#f43f5e', bg: 'rgba(244, 63, 94, 0.15)', icon: 'cancel' },
        'JUNKSHOP': { color: '#f59e0b', bg: 'rgba(245, 158, 11, 0.15)', icon: 'recycling' },
        'DONATION/CHARITY': { color: '#3b82f6', bg: 'rgba(59, 130, 246, 0.15)', icon: 'volunteer_activism' },
        'PENDING': { color: '#94a3b8', bg: 'rgba(148, 163, 184, 0.15)', icon: 'hourglass_empty' }
    };
    const upper = status.toUpperCase();
    const info = statusMap[upper] || { color: '#94a3b8', bg: 'rgba(148,163,184,0.15)', icon: 'label' };
    return `<span class="item-status-badge" style="background:${info.bg}; color:${info.color}">
        <span class="material-icons-round">${info.icon}</span>${status}
    </span>`;
}

// ======================== STATS ========================
let inventoryChartInstance = null;

function updateStats() {
    let totals = {
        items: 0,
        figurines: 0,
        warehouse: 0,
        boxes: 0,
        lipat_bahay: 0
    };

    if (cachedData['all_inventory']) totals.items = cachedData['all_inventory'].length;
    if (cachedData['figurine_list']) totals.figurines = cachedData['figurine_list'].length;
    if (cachedData['item_from_warehouse']) totals.warehouse = cachedData['item_from_warehouse'].length;
    if (cachedData['all_box']) totals.boxes = cachedData['all_box'].length;
    if (cachedData['lipat_bahay']) totals.lipat_bahay = cachedData['lipat_bahay'].length;

    // Dashboard stat cards
    animateNumber(document.querySelector('#statAllInventory .stat-number'), totals.items);
    animateNumber(document.querySelector('#statFigurines .stat-number'), totals.figurines);
    animateNumber(document.querySelector('#statWarehouse .stat-number'), totals.warehouse);
    animateNumber(document.querySelector('#statBoxes .stat-number'), totals.boxes);

    // Sidebar badges
    const updateBadge = (id, count) => {
        const el = document.getElementById(id);
        if (el) el.textContent = count;
    };
    updateBadge('badgeAllInventory', totals.items);
    updateBadge('badgeFigurineList', totals.figurines);
    updateBadge('badgeWarehouse', totals.warehouse);
    updateBadge('badgeAllBox', totals.boxes);
    updateBadge('badgeLipatBahay', totals.lipat_bahay);

    // Make stat cards clickable
    document.getElementById('statAllInventory').onclick = () => switchTab('all_inventory');
    document.getElementById('statFigurines').onclick = () => switchTab('figurine_list');
    document.getElementById('statWarehouse').onclick = () => switchTab('item_from_warehouse');
    document.getElementById('statBoxes').onclick = () => switchTab('all_box');

    // Update Chart.js Graphic
    const ctx = document.getElementById('inventoryChart');
    if (ctx) {
        const dataValues = [totals.items, totals.figurines, totals.warehouse, totals.boxes, totals.lipat_bahay];
        const labels = ['All Inventory', 'Figurines', 'Warehouse', 'Boxes', 'Lipat Bahay'];
        const bgColors = [
            'rgba(99, 102, 241, 0.8)',
            'rgba(6, 182, 212, 0.8)',
            'rgba(245, 158, 11, 0.8)',
            'rgba(16, 185, 129, 0.8)',
            'rgba(236, 72, 153, 0.8)'
        ];

        if (inventoryChartInstance) {
            inventoryChartInstance.data.datasets[0].data = dataValues;
            inventoryChartInstance.update();
        } else {
            inventoryChartInstance = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: labels,
                    datasets: [{
                        data: dataValues,
                        backgroundColor: bgColors,
                        borderWidth: 0,
                        hoverOffset: 10
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'right',
                            labels: {
                                color: getComputedStyle(document.body).getPropertyValue('--text-secondary').trim() || '#94a3b8',
                                font: {
                                    family: "'Inter', sans-serif",
                                    size: 13
                                }
                            }
                        }
                    },
                    cutout: '70%'
                }
            });
        }
    }
}

function animateNumber(el, target) {
    if (!el) return;
    const current = parseInt(el.textContent) || 0;
    if (current === target) { el.textContent = target; return; }

    const duration = 600;
    const start = performance.now();

    function update(now) {
        const elapsed = now - start;
        const progress = Math.min(elapsed / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
        const value = Math.round(current + (target - current) * eased);
        el.textContent = value;
        if (progress < 1) requestAnimationFrame(update);
    }

    requestAnimationFrame(update);
}

// ======================== SEARCH ========================
function setupSearch() {
    // Global search
    document.getElementById('globalSearch').addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (currentTab === 'dashboard') return;
        filterCurrentTable(query);
    });

    // Tab-specific search
    document.querySelectorAll('.tab-search').forEach(input => {
        input.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            const sheet = input.dataset.sheet;
            filterTable(sheet, query);
        });
    });
}

function filterCurrentTable(query) {
    filterTable(currentTab, query);
}

function filterTable(sheetKey, query) {
    const tbody = document.querySelector(`#${SHEET_META[sheetKey]?.tableId || 'tableAllInventory'} tbody`);

    if (sheetKey === 'all_inventory' && currentTab === 'all_inventory') {
        const allItems = [];
        Object.keys(SHEET_META).forEach(key => {
            const data = cachedData[key] || [];
            data.forEach(item => {
                allItems.push({
                    ...item,
                    _sourceSheet: key,
                    _sourceName: SHEET_META[key].name,
                    _displayCode: item['CODE'] || item['FIGURINE'] || item['# NUMBER OF ITEM'] || item['BOX NUMBER'] || item['ITEM/BOX NAME'] || '—',
                    _displayName: item['NAME OF ITEM'] || item['ITEM NAME'] || item['BOX NAME'] || item['ITEM/BOX NAME'] || '—',
                    _displayPlace: item['PLACE OF ITEM'] || item['FIGURINE PLACE'] || item['DESTINATION PLACE'] || item['BOX PLACE'] || item['ORIGIN ROOM'] || '—',
                    _displayStatus: item['UPDATE'] || item['FIGURINE UPDATE'] || item['COLOR OF ITEM'] || item['BOX UPDATE'] || item['STATUS'] || '—'
                });
            });
        });

        if (!query) {
            renderMasterRows(tbody, allItems);
            return;
        }

        const filtered = allItems.filter(row => {
            return [row._displayCode, row._displayName, row._displayPlace, row._displayStatus, row._sourceName].some(val => 
                val.toString().toLowerCase().includes(query)
            );
        });

        renderMasterRows(tbody, filtered);
        return;
    }

    const meta = SHEET_META[sheetKey];
    if (!meta) return;

    const data = cachedData[sheetKey] || [];
    if (!query) {
        renderTable(sheetKey, data);
        return;
    }

    const filtered = data.filter(row => {
        return meta.headers.some(header => {
            const val = (row[header] || '').toString().toLowerCase();
            return val.includes(query);
        });
    });

    renderTable(sheetKey, filtered);
}

// ======================== MODALS ========================
function setupModals() {
    // Add New button
    document.getElementById('btnAddNew').addEventListener('click', () => {
        if (currentTab === 'dashboard') {
            showToast('Select a category first, or use Quick Actions below', 'info');
            return;
        }
        openAddModal(currentTab);
    });

    // Add buttons in each tab
    document.querySelectorAll('[data-action="add"]').forEach(btn => {
        btn.addEventListener('click', () => {
            openAddModal(btn.dataset.sheet);
        });
    });

    // Close modal
    document.getElementById('btnCloseModal').addEventListener('click', closeItemModal);
    document.getElementById('btnCancelForm').addEventListener('click', closeItemModal);

    // Submit form
    document.getElementById('btnSubmitForm').addEventListener('click', submitForm);

    // Image viewer modal
    document.getElementById('btnCloseViewer').addEventListener('click', () => {
        document.getElementById('imageViewerModal').style.display = 'none';
    });

    // Delete modal
    document.getElementById('btnCancelDelete').addEventListener('click', () => {
        document.getElementById('deleteModal').style.display = 'none';
    });
    document.getElementById('btnConfirmDelete').addEventListener('click', confirmDelete);

    // Close modals on overlay click
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                overlay.style.display = 'none';
            }
        });
    });
}

function getNextSequence(sheetKey) {
    const data = cachedData[sheetKey] || [];
    let fieldKey = '';
    let prefix = '';

    if (sheetKey === 'all_inventory') { fieldKey = 'CODE'; prefix = 'A-'; }
    else if (sheetKey === 'figurine_list') { fieldKey = 'FIGURINE'; prefix = 'F-'; }
    else if (sheetKey === 'item_from_warehouse') { fieldKey = '# NUMBER OF ITEM'; prefix = ''; }
    else if (sheetKey === 'all_box') { fieldKey = 'BOX NUMBER'; prefix = 'B-'; }
    else { return ''; }

    let maxNum = 0;
    data.forEach(row => {
        const val = row[fieldKey];
        if (val) {
            const strVal = val.toString();
            // Match any trailing numbers (e.g., 'A-15' -> '15', 'F-2' -> '2', '24' -> '24')
            const match = strVal.match(/(\d+)$/);
            if (match) {
                const num = parseInt(match[1], 10);
                if (num > maxNum) maxNum = num;
            }
        }
    });

    return prefix + (maxNum + 1);
}

function openAddModal(sheetKey) {
    const meta = SHEET_META[sheetKey];
    if (!meta) return;

    // Switch to that tab if coming from dashboard
    if (currentTab === 'dashboard') {
        switchTab(sheetKey);
    }

    document.getElementById('modalTitle').textContent = `Add ${meta.name}`;
    document.getElementById('modalIcon').textContent = meta.icon;
    document.getElementById('formSheet').value = sheetKey;
    document.getElementById('formRowIndex').value = '';

    // Automatically determine the next logical number sequence
    const nextSeq = getNextSequence(sheetKey);
    let initialData = {};
    if (nextSeq) {
        if (sheetKey === 'all_inventory') initialData['CODE'] = nextSeq;
        if (sheetKey === 'figurine_list') initialData['FIGURINE'] = nextSeq;
        if (sheetKey === 'item_from_warehouse') initialData['# NUMBER OF ITEM'] = nextSeq;
        if (sheetKey === 'all_box') initialData['BOX NUMBER'] = nextSeq;
    }

    buildFormFields(meta.formFields, initialData);
    document.getElementById('itemModal').style.display = 'flex';
}

function openEditModal(sheetKey, rowIndex) {
    const meta = SHEET_META[sheetKey];
    const data = cachedData[sheetKey] || [];
    const row = data.find(r => r._rowIndex === rowIndex);

    if (!row) {
        showToast('Item not found', 'error');
        return;
    }

    document.getElementById('modalTitle').textContent = `Edit ${meta.name}`;
    document.getElementById('modalIcon').textContent = 'edit';
    document.getElementById('formSheet').value = sheetKey;
    document.getElementById('formRowIndex').value = rowIndex;

    buildFormFields(meta.formFields, row);
    document.getElementById('itemModal').style.display = 'flex';
}

function buildFormFields(fields, data) {
    const container = document.getElementById('formFields');
    container.innerHTML = '';

    fields.forEach(field => {
        const group = document.createElement('div');
        group.className = 'form-group';

        const label = document.createElement('label');
        label.textContent = field.label;
        label.setAttribute('for', `field_${field.key}`);
        group.appendChild(label);

        const value = data[field.key] || '';

        if (field.type === 'select') {
            const select = document.createElement('select');
            select.id = `field_${field.key}`;
            select.dataset.key = field.key;

            // Add blank option
            const blankOpt = document.createElement('option');
            blankOpt.value = '';
            blankOpt.textContent = '— Select —';
            select.appendChild(blankOpt);

            field.options.forEach(opt => {
                const option = document.createElement('option');
                option.value = opt;
                option.textContent = opt;
                if (value.toUpperCase() === opt.toUpperCase()) option.selected = true;
                select.appendChild(option);
            });

            group.appendChild(select);
        } else if (field.type === 'textarea') {
            const textarea = document.createElement('textarea');
            textarea.id = `field_${field.key}`;
            textarea.dataset.key = field.key;
            textarea.placeholder = field.placeholder || '';
            textarea.value = value;
            group.appendChild(textarea);
        } else if (field.type === 'image') {
            // Image upload zone
            const uploadZone = buildImageUploadZone(field.key, value);
            group.appendChild(uploadZone);
        } else {
            const input = document.createElement('input');
            input.type = field.type || 'text';
            input.id = `field_${field.key}`;
            input.dataset.key = field.key;
            input.placeholder = field.placeholder || '';
            input.value = value;
            if (field.required) input.required = true;
            
            if (['CODE', 'FIGURINE', '# NUMBER OF ITEM', 'BOX NUMBER'].includes(field.key)) {
                input.readOnly = true;
                input.style.backgroundColor = 'var(--bg-body)';
                input.style.color = 'var(--text-secondary)';
                input.style.cursor = 'not-allowed';
                input.style.border = '1px solid var(--border-color)';
                input.title = "This number is automatically generated";
            }

            group.appendChild(input);
        }

        container.appendChild(group);
    });
}

function closeItemModal() {
    document.getElementById('itemModal').style.display = 'none';
    pendingImageUploads = {};
}

// ======================== IMAGE UPLOAD ========================
function buildImageUploadZone(fieldKey, existingUrl) {
    const wrapper = document.createElement('div');
    wrapper.className = 'image-upload-wrapper';

    // Hidden input to store the URL value
    const hiddenInput = document.createElement('input');
    hiddenInput.type = 'hidden';
    hiddenInput.id = `field_${fieldKey}`;
    hiddenInput.dataset.key = fieldKey;
    hiddenInput.value = existingUrl || '';
    wrapper.appendChild(hiddenInput);

    // Preview area
    const preview = document.createElement('div');
    preview.className = 'image-preview';
    preview.id = `preview_${fieldKey}`;
    if (existingUrl) {
        preview.innerHTML = `<img src="${existingUrl}" alt="Current image" /><button type="button" class="btn-icon remove-image" title="Remove image"><span class="material-icons-round">close</span></button>`;
        preview.classList.add('has-image');
        preview.querySelector('.remove-image').addEventListener('click', () => {
            hiddenInput.value = '';
            fileInput.value = ''; // Reset file input
            delete pendingImageUploads[fieldKey];
            preview.innerHTML = '';
            preview.classList.remove('has-image');
            dropZone.style.display = '';
        });
    }
    wrapper.appendChild(preview);

    // Drop zone
    const dropZone = document.createElement('div');
    dropZone.className = 'image-drop-zone';
    dropZone.id = `dropzone_${fieldKey}`;
    if (existingUrl) dropZone.style.display = 'none';
    dropZone.innerHTML = `
        <span class="material-icons-round">add_a_photo</span>
        <p>Take a photo, drop image, or click to upload</p>
        <span class="drop-zone-hint">Supports Camera, JPG, PNG, WebP</span>
    `;
    wrapper.appendChild(dropZone);

    // File input (hidden)
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/*';
    fileInput.setAttribute('capture', 'environment'); // Suggests camera on mobile
    fileInput.style.display = 'none';
    fileInput.id = `file_${fieldKey}`;
    wrapper.appendChild(fileInput);

    // URL input fallback
    const urlRow = document.createElement('div');
    urlRow.className = 'image-url-row';
    urlRow.innerHTML = `
        <span class="url-divider">or enter URL</span>
        <input type="url" class="image-url-input" id="url_${fieldKey}" placeholder="https://..." value="${existingUrl || ''}" />
    `;
    wrapper.appendChild(urlRow);

    // --- Event Handlers ---

    // Click to upload
    dropZone.addEventListener('click', () => fileInput.click());

    // File selected
    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
            handleImageFile(e.target.files[0], fieldKey, hiddenInput, preview, dropZone);
        }
    });

    // Drag & Drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            handleImageFile(e.dataTransfer.files[0], fieldKey, hiddenInput, preview, dropZone);
        }
    });

    // Paste from clipboard (on the entire wrapper)
    wrapper.addEventListener('paste', (e) => {
        const items = e.clipboardData.items;
        for (let i = 0; i < items.length; i++) {
            if (items[i].type.startsWith('image/')) {
                e.preventDefault();
                const file = items[i].getAsFile();
                handleImageFile(file, fieldKey, hiddenInput, preview, dropZone);
                break;
            }
        }
    });
    // Make wrapper focusable for paste events
    wrapper.setAttribute('tabindex', '0');

    // URL input change
    const urlInput = urlRow.querySelector('.image-url-input');
    urlInput.addEventListener('input', () => {
        const url = urlInput.value.trim();
        hiddenInput.value = url;
        delete pendingImageUploads[fieldKey];
        if (url) {
            preview.innerHTML = `<img src="${url}" alt="Preview" onerror="this.style.display='none'" /><button type="button" class="btn-icon remove-image" title="Remove"><span class="material-icons-round">close</span></button>`;
            preview.classList.add('has-image');
            dropZone.style.display = 'none';
            preview.querySelector('.remove-image').addEventListener('click', () => {
                hiddenInput.value = '';
                fileInput.value = ''; // Reset file input
                urlInput.value = '';
                preview.innerHTML = '';
                preview.classList.remove('has-image');
                dropZone.style.display = '';
            });
        } else {
            preview.innerHTML = '';
            preview.classList.remove('has-image');
            dropZone.style.display = '';
        }
    });

    return wrapper;
}

function handleImageFile(file, fieldKey, hiddenInput, preview, dropZone) {
    if (!file.type.startsWith('image/')) {
        showToast('Please select an image file', 'error');
        return;
    }

    // Max 5MB
    if (file.size > 5 * 1024 * 1024) {
        showToast('Image must be under 5MB', 'error');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const base64 = e.target.result;

        // Store for upload on submit
        pendingImageUploads[fieldKey] = {
            data: base64,
            fileName: file.name || `image_${Date.now()}.png`
        };

        // Show preview
        preview.innerHTML = `<img src="${base64}" alt="Preview" /><button type="button" class="btn-icon remove-image" title="Remove image"><span class="material-icons-round">close</span></button>`;
        preview.classList.add('has-image');
        dropZone.style.display = 'none';

        // Clear URL input and value
        const urlInput = document.getElementById(`url_${fieldKey}`);
        if (urlInput) urlInput.value = '';
        
        // Reset file input value after processing so the same file can be selected again
        // fileInput.value = ''; // Wait, if I reset it here, it might lose reference? 
        // No, reader is already finished or started. 
        // Actually, let's reset it here so the 'change' event can trigger again for the same file.
        const fileInput = document.getElementById(`file_${fieldKey}`);
        if (fileInput) fileInput.value = '';

        // Remove button handler
        preview.querySelector('.remove-image').addEventListener('click', () => {
            hiddenInput.value = '';
            fileInput.value = ''; // Reset file input
            delete pendingImageUploads[fieldKey];
            preview.innerHTML = '';
            preview.classList.remove('has-image');
            dropZone.style.display = '';
        });
    };
    reader.readAsDataURL(file);
}

async function uploadPendingImages() {
    const keys = Object.keys(pendingImageUploads);
    if (keys.length === 0) return;

    for (const key of keys) {
        const { data, fileName } = pendingImageUploads[key];

        if (CONFIG.DEMO_MODE) {
            // In demo mode, just use the base64 data directly
            const hiddenInput = document.getElementById(`field_${key}`);
            if (hiddenInput) hiddenInput.value = data;
            continue;
        }

        // Upload to Google Drive via Apps Script
        try {
            const response = await fetch(CONFIG.SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    auth: sessionStorage.getItem('authPin') || '',
                    action: 'upload_image',
                    sheet: 'upload', // dummy, won't be used
                    imageData: data,
                    fileName: fileName
                }),
            });
            const result = await response.json();
            if (result.url) {
                const hiddenInput = document.getElementById(`field_${key}`);
                if (hiddenInput) hiddenInput.value = result.url;
            } else {
                throw new Error(result.error || 'Upload failed');
            }
        } catch (err) {
            showToast(`Image upload failed: ${err.message}`, 'error');
            throw err;
        }
    }

    pendingImageUploads = {};
}

async function submitForm() {
    const sheetKey = document.getElementById('formSheet').value;
    const rowIndex = document.getElementById('formRowIndex').value;
    const meta = SHEET_META[sheetKey];

    // Gather form data
    const formData = {};
    meta.formFields.forEach(field => {
        const el = document.getElementById(`field_${field.key}`);
        if (el) formData[field.key] = el.value;
    });

    // Validate required
    const missing = meta.formFields.filter(f => f.required && !formData[f.key]);
    if (missing.length > 0) {
        showToast(`Please fill in: ${missing.map(f => f.label).join(', ')}`, 'error');
        return;
    }

    const isEdit = !!rowIndex;

    // Upload any pending images first
    const submitBtn = document.getElementById('btnSubmitForm');
    if (Object.keys(pendingImageUploads).length > 0) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<span class="material-icons-round">cloud_upload</span> Uploading image...';
        try {
            await uploadPendingImages();
        } catch (err) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<span class="material-icons-round">save</span> Save to Spreadsheet';
            return;
        }
        // Re-gather form data after upload (hidden inputs now have URLs)
        meta.formFields.forEach(field => {
            const el = document.getElementById(`field_${field.key}`);
            if (el) formData[field.key] = el.value;
        });
    }

    if (CONFIG.DEMO_MODE) {
        // In demo mode, update local cache
        if (isEdit) {
            const idx = cachedData[sheetKey].findIndex(r => r._rowIndex === parseInt(rowIndex));
            if (idx !== -1) {
                Object.assign(cachedData[sheetKey][idx], formData);
            }
        } else {
            const newRow = { _rowIndex: Date.now(), ...formData };
            cachedData[sheetKey].push(newRow);
        }
        renderTable(sheetKey, cachedData[sheetKey]);
        updateStats();
        closeItemModal();
        showToast(isEdit ? 'Item updated (demo mode)' : 'Item added (demo mode)', 'success');
        return;
    }

    // Live mode — send to Google Sheets
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<span class="material-icons-round">hourglass_top</span> Saving...';

    try {
        const payload = {
            auth: sessionStorage.getItem('authPin') || '',
            action: isEdit ? 'update' : 'add',
            sheet: sheetKey,
            data: formData,
        };
        if (isEdit) payload.row = parseInt(rowIndex);

        const response = await fetch(CONFIG.SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload),
        });

        const result = await response.json();

        if (result.error) {
            throw new Error(result.error);
        }

        // Reload this sheet's data
        const freshData = await fetchSheetData(sheetKey);
        cachedData[sheetKey] = freshData;
        renderTable(sheetKey, freshData);
        updateStats();

        closeItemModal();
        showToast(isEdit ? 'Item updated in Google Sheets!' : 'Item added to Google Sheets!', 'success');
    } catch (err) {
        console.error('Submit error:', err);
        showToast('Error: ' + err.message, 'error');
    } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<span class="material-icons-round">save</span> Save to Spreadsheet';
    }
}

// ======================== DELETE ========================
function openDeleteModal(sheetKey, rowIndex, name) {
    deleteTarget = { sheetKey, rowIndex };
    document.getElementById('deleteItemName').textContent = name || `Row ${rowIndex}`;
    document.getElementById('deleteModal').style.display = 'flex';
}

async function confirmDelete() {
    if (!deleteTarget) return;

    const { sheetKey, rowIndex } = deleteTarget;

    if (CONFIG.DEMO_MODE) {
        cachedData[sheetKey] = cachedData[sheetKey].filter(r => r._rowIndex !== rowIndex);
        renderTable(sheetKey, cachedData[sheetKey]);
        updateStats();
        document.getElementById('deleteModal').style.display = 'none';
        showToast('Item deleted (demo mode)', 'success');
        deleteTarget = null;
        return;
    }

    const confirmBtn = document.getElementById('btnConfirmDelete');
    confirmBtn.disabled = true;
    confirmBtn.innerHTML = '<span class="material-icons-round">hourglass_top</span> Deleting...';

    try {
        const payload = {
            auth: sessionStorage.getItem('authPin') || '',
            action: 'delete',
            sheet: sheetKey,
            row: rowIndex,
        };

        const response = await fetch(CONFIG.SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload),
        });

        const result = await response.json();
        if (result.error) throw new Error(result.error);

        // Reload
        const freshData = await fetchSheetData(sheetKey);
        cachedData[sheetKey] = freshData;
        renderTable(sheetKey, freshData);
        updateStats();

        document.getElementById('deleteModal').style.display = 'none';
        showToast('Item deleted from Google Sheets!', 'success');
    } catch (err) {
        console.error('Delete error:', err);
        showToast('Error deleting: ' + err.message, 'error');
    } finally {
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = '<span class="material-icons-round">delete</span> Delete';
        deleteTarget = null;
    }
}

// ======================== IMAGE VIEWER ========================
function openImageViewer(url, name, status) {
    document.getElementById('viewerImage').src = url;
    document.getElementById('viewerItemName').textContent = name || 'Item Picture';
    
    const statusEl = document.getElementById('viewerItemStatus');
    if (status && status !== 'No Status') {
        statusEl.textContent = status;
        statusEl.style.display = 'inline-block';
    } else {
        statusEl.style.display = 'none';
    }

    document.getElementById('imageViewerModal').style.display = 'flex';
}

// ======================== QUICK ACTIONS ========================
function setupQuickActions() {
    document.querySelectorAll('.quick-action-card').forEach(card => {
        card.addEventListener('click', () => {
            const action = card.dataset.action;
            const sheet = card.dataset.sheet;
            if (action === 'add') {
                switchTab(sheet);
                setTimeout(() => openAddModal(sheet), 200);
            }
        });
    });
}

// ======================== TOAST ========================
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    const icons = { success: 'check_circle', error: 'error', info: 'info' };
    toast.innerHTML = `
        <span class="material-icons-round">${icons[type] || 'info'}</span>
        <span class="toast-message">${message}</span>
    `;

    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('toast-out');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// ======================== LOADING ========================
function showLoading() {
    document.getElementById('loadingOverlay').classList.remove('hidden');
}

function hideLoading() {
    setTimeout(() => {
        document.getElementById('loadingOverlay').classList.add('hidden');
    }, 500);
}

function setupDatabaseManager() {
    // One-Click Init
    const btnInit = document.getElementById('btnInitDatabase');
    if (btnInit) {
        btnInit.addEventListener('click', async () => {
            if (!CONFIG.SCRIPT_URL) {
                showToast('Connect to Google Sheets first!', 'error');
                return;
            }
            if (!confirm('This will create all missing standard tabs in your Google Sheet. Continue?')) return;

            showLoading();
            try {
                const response = await fetch(CONFIG.SCRIPT_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify({
                        auth: sessionStorage.getItem('authPin') || '',
                        action: 'init_database'
                    })
                });
                const result = await response.json();
                if (result.success) {
                    showToast('Database initialized successfully!', 'success');
                    loadAllData();
                } else {
                    throw new Error(result.error);
                }
            } catch (err) {
                showToast('Init failed: ' + err.message, 'error');
            } finally {
                hideLoading();
            }
        });
    }

    // Custom Creator Modal
    const dbModal = document.getElementById('dbCreatorModal');
    const btnOpen = document.getElementById('btnOpenDbCreator');
    if (btnOpen && dbModal) {
        btnOpen.addEventListener('click', () => {
            dbModal.style.display = 'flex';
        });
        document.getElementById('btnCloseDbModal').addEventListener('click', () => dbModal.style.display = 'none');
        document.getElementById('btnCancelDbForm').addEventListener('click', () => dbModal.style.display = 'none');
    }

    // Create Custom Sheet
    const btnCreate = document.getElementById('btnCreateCustomSheet');
    if (btnCreate) {
        btnCreate.addEventListener('click', async () => {
            const name = document.getElementById('customSheetName').value.trim();
            const headersRaw = document.getElementById('customHeaders').value.trim();
            
            if (!name) {
                showToast('Please enter a table name', 'error');
                return;
            }

            const headers = headersRaw.split(',').map(h => h.trim()).filter(h => h !== '');
            
            showLoading();
            try {
                const response = await fetch(CONFIG.SCRIPT_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify({
                        auth: sessionStorage.getItem('authPin') || '',
                        action: 'create_sheet',
                        sheetName: name,
                        headers: headers
                    })
                });
                const result = await response.json();
                if (result.success) {
                    showToast(`Sheet "${name}" created!`, 'success');
                    dbModal.style.display = 'none';
                    document.getElementById('customSheetName').value = '';
                    document.getElementById('customHeaders').value = '';
                } else {
                    throw new Error(result.error);
                }
            } catch (err) {
                showToast('Error: ' + err.message, 'error');
            } finally {
                hideLoading();
            }
        });
    }
}

