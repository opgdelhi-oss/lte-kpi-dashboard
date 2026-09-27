// app.js - Main Application Logic for LTE & VoLTE KPI Dashboard
// Maintains Site Name and LNCEL completely separate (no concatenation).

(function () {
    'use strict';

    // State Management
    const state = {
        rawRecords: [],
        filteredRecords: [],
        hourlyRecords: [],
        daywiseRecords: [],
        dataSourceName: 'Final_24 hrs_Hourly 4G KPI Report 24-09-2026.xlsx',
        isUserReport: true,
        detectedColumns: {},
        activeTheme: localStorage.getItem('lte_dashboard_theme') || 'dark',
        viewMode: 'hourly',
        filters: {
            operator: 'ALL',
            date: 'ALL',
            band: 'ALL',
            hour: 'ALL',
            site: 'ALL',
            lncel: 'ALL',
            busyHourOnly: false,
            degradedOnly: false,
            alertDateFilterTouched: false,
            lowThroughputLowUtilization: false
        },
        pagination: {
            currentPage: 1,
            pageSize: 15,
            sortColumn: 'Total_Traffic_GB',
            sortDirection: 'desc'
        },
        charts: {}
    };

    // KPI Schema mapping
    const KPI_DEFINITIONS = {
        'Date': { label: 'Date', aliases: ['date', 'day', 'measurement date', 'period start time', 'timestamp'] },
        'Operator': { label: 'Operator', aliases: ['operator', 'op', 'carrier', 'telecom operator', 'plmn'] },
        'Band': { label: 'Band / Layer', aliases: ['band', 'duplex', 'frequency', 'freq', 'carrier'] },
        'Timestamp': { label: 'Time / Date', aliases: ['period start time', 'time', 'timestamp', 'begintime'] },
        'Site_Name': { label: 'Site / eNodeB Name', aliases: ['mrbts/sbts name', 'mrbts_name', 'sbts name', 'sitename', 'site_name', 'site name', 'enodeb_name', 'bts_name'] },
        'Site_ID': { label: 'Site / MRBTS ID', aliases: ['mrbts', 'mrbts id', 'site_id', 'site id', 'enodeb_id'] },
        'LNCEL': { label: 'LNCEL (Sector ID)', aliases: ['lncel', 'sector', 'cell_id', 'localcellid', 'eci'] },
        'Cell_Availability': { label: 'Cell Availability (%)', target: 99.5, higherIsBetter: true, unit: '%', aliases: ['cell availability lte-4g_kpi_report', 'cell availability', 'availability', 'radio network availability'] },
        'VoLTE_CSSR': { label: 'VoLTE Call Setup SR (%)', target: 99.0, higherIsBetter: true, unit: '%', aliases: ['nokia_lte_volte call setup success rate', 'volte call setup success rate', 'volte_cssr', 'volte cssr'] },
        'E2E_CSSR': { label: 'E2E Call Setup SR (%)', target: 98.5, higherIsBetter: true, unit: '%', aliases: ['e2e call setup success rate - all bearer new', 'e2e call setup success rate', 'cssr'] },
        'RRC_Setup_SR': { label: 'RRC Setup SR (%)', target: 98.5, higherIsBetter: true, unit: '%', aliases: ['rrc setup success rate', 'rrc_setup_sr', 'rrc_succ_rate', 'pmrrcconnestabsucc'] },
        'ERAB_Setup_SR': { label: 'E-RAB Setup SR (%)', target: 98.5, higherIsBetter: true, unit: '%', aliases: ['erab setup success rate', 'erab_setup_sr', 'erab_succ_rate', 'pmerabestabsucc'] },
        'Total_Traffic_GB': { label: 'Total Payload (GB)', unit: 'GB', aliases: ['nsn_data volume - total_gb', 'data volume - total_gb', 'total_traffic_gb', 'payload_gb', 'data volume gb'] },
        'VoLTE_Traffic_Erl': { label: 'VoLTE Traffic (Erl)', unit: 'Erl', aliases: ['nokia_lte_volte traffic erl', 'volte traffic erl', 'voice_traffic_erl'] },
        'VoLTE_Drop_Rate': { label: 'VoLTE Drop Rate (%)', target: 1.0, higherIsBetter: false, unit: '%', aliases: ['nsn_lte_volte_drop_call_rate', 'volte_drop_call_rate', 'volte drop rate'] },
        'Drop_Rate': { label: 'Service Drop Rate (%)', target: 0.8, higherIsBetter: false, unit: '%', aliases: ['service dcr active e-rab', 'service drop rate', 'drop_rate', 'erab_drop_rate'] },
        'DL_User_Throughput_Mbps': { label: 'DL User Thrpt (Mbps)', target: 3.0, higherIsBetter: true, unit: 'Mbps', aliases: ['nsn_lte_dl throughput per ue (na)', 'dl throughput per ue', 'dl_user_throughput', 'dl_throughput'] },
        'UL_User_Throughput_Mbps': { label: 'UL User Thrpt (Mbps)', target: 3.5, higherIsBetter: true, unit: 'Mbps', aliases: ['nsn_lte_ul throughput per ue (na)', 'ul throughput per ue', 'ul_user_throughput'] },
        'DL_PRB_Util': { label: 'DL PRB Util (%)', target: 80.0, higherIsBetter: false, unit: '%', aliases: ['pdsch utilization rate prb dl', 'dl prb utilization', 'dl_prb_util'] },
        'UL_PRB_Util': { label: 'UL PRB Util (%)', target: 80.0, higherIsBetter: false, unit: '%', aliases: ['pdsch utilization rate prb ul', 'ul prb utilization', 'ul_prb_util'] },
        'CQI': { label: 'Avg CQI', target: 8.5, higherIsBetter: true, aliases: ['e-utran average cqi', 'average cqi', 'cqi'] },
        'SINR': { label: 'SINR (dB)', target: 10.0, higherIsBetter: true, unit: 'dB', aliases: ['sinr', 'average sinr'] },
        'RSSI': { label: 'UL RSSI (dBm)', target: -105.0, higherIsBetter: false, unit: 'dBm', aliases: ['ul pusch rssi', 'pusch rssi', 'rssi'] },
        'Active_Users_Avg': { label: 'Avg Active UEs', aliases: ['avg rrc conn ue', 'active users', 'active_users_avg'] },
        'Max_Users': { label: 'Max Active UEs', aliases: ['max rrc connected users', 'max active users'] },
        'Handover_SR': { label: 'Intra-Freq HO SR (%)', target: 95.0, higherIsBetter: true, unit: '%', aliases: ['intra frequency handover sucess rate new', 'handover sr', 'hosr'] }
    };

    const LNCEL_BAND_MAP = {
        '1': 'FDD', '2': 'FDD', '3': 'FDD', '4': 'FDD', '5': 'FDD',
        '6': 'FDD', '7': 'FDD', '8': 'FDD',
        '21': 'FDD', '22': 'FDD', '23': 'FDD', '24': 'FDD',
        '25': 'FDD', '26': 'FDD', '27': 'FDD', '28': 'FDD', '29': 'FDD',
        '51': 'TDD', '52': 'TDD', '53': 'TDD', '54': 'TDD', '55': 'TDD',
        '56': 'TDD', '57': 'TDD', '58': 'TDD', '59': 'TDD', '60': 'TDD', '61': 'TDD',
        '101': 'TDD', '102': 'TDD', '103': 'TDD', '104': 'TDD', '105': 'TDD',
        '106': 'TDD', '107': 'TDD', '108': 'TDD', '109': 'TDD', '110': 'TDD',
        '111': 'TDD', '112': 'TDD', '113': 'TDD', '114': 'TDD', '115': 'TDD',
        '116': 'TDD', '117': 'TDD', '118': 'TDD', '119': 'TDD'
    };

    function resolveRecordBand(record) {
        const lncel = String(record.LNCEL ?? '').trim().replace(/\.0+$/, '');
        return LNCEL_BAND_MAP[lncel] || String(record.Band || '').trim() || 'Unknown';
    }

    // UI Element References
    const elements = {
        themeToggleBtn: document.getElementById('themeToggleBtn'),
        fileInput: document.getElementById('fileInput'),
        dropzone: document.getElementById('dropzone'),
        loadUserReportBtn: document.getElementById('loadUserReportBtn'),
        loadSampleBtn: document.getElementById('loadSampleBtn'),
        dataSourceTitle: document.getElementById('dataSourceTitle'),
        dataSourceSubtitle: document.getElementById('dataSourceSubtitle'),
        alertBanner: document.getElementById('alertBanner'),
        alertText: document.getElementById('alertText'),
        filterOperator: document.getElementById('filterOperator'),
        filterDate: document.getElementById('filterDate'),
        filterBand: document.getElementById('filterBand'),
        filterHour: document.getElementById('filterHour'),
        filterSite: document.getElementById('filterSite'),
        filterLncel: document.getElementById('filterLncel'),
        filterBusyHour: document.getElementById('filterBusyHour'),
        filterDegraded: document.getElementById('filterDegraded'),
        filterLowThroughputLowUtilization: document.getElementById('filterLowThroughputLowUtilization'),
        resetFiltersBtn: document.getElementById('resetFiltersBtn'),
        exportCsvBtn: document.getElementById('exportCsvBtn'),
        tableSearch: document.getElementById('tableSearch'),
        cellTableBody: document.getElementById('cellTableBody'),
        prevPageBtn: document.getElementById('prevPageBtn'),
        nextPageBtn: document.getElementById('nextPageBtn'),
        paginationInfo: document.getElementById('paginationInfo'),
        operatorPillContainer: document.getElementById('operatorPillContainer'),
        viewModeContainer: document.getElementById('viewModeContainer'),
        worstCellsChartSubtitle: document.getElementById('worstCellsChartSubtitle'),
        lowThroughputChartSubtitle: document.getElementById('lowThroughputChartSubtitle'),
        trendChartTitle: document.getElementById('trendChartTitle'),
        trendChartSubtitle: document.getElementById('trendChartSubtitle'),
        hourlyProfileChartTitle: document.getElementById('hourlyProfileChartTitle'),
        hourlyProfileChartSubtitle: document.getElementById('hourlyProfileChartSubtitle'),
        volteQualityChartTitle: document.getElementById('volteQualityChartTitle'),
        volteQualityChartSubtitle: document.getElementById('volteQualityChartSubtitle'),
        mappingModal: document.getElementById('mappingModal'),
        closeModalBtn: document.getElementById('closeModalBtn'),
        cancelModalBtn: document.getElementById('cancelModalBtn'),
        saveMappingBtn: document.getElementById('saveMappingBtn'),
        mappingGrid: document.getElementById('mappingGrid')
    };

    // Initialize Dashboard
    function init() {
        applyTheme(state.activeTheme);
        setupEventListeners();
        setupDataSyncButton();
        setupInventorySiteFilter();
        loadDefaultReportData();
        startDataRefreshMonitor();
    }

    function setupDataSyncButton() {
        const button = document.getElementById('syncKpiDataBtn');
        const status = document.getElementById('syncKpiDataStatus');
        if (!button || !status) return;
        const buttonLabel = button.querySelector('span:last-child');

        button.addEventListener('click', async () => {
            button.disabled = true;
            buttonLabel.textContent = 'SYNCING...';
            status.textContent = 'Checking the OneDrive KPI folder…';
            try {
                const response = await fetch('/api/sync-data', { method: 'POST', cache: 'no-store' });
                const result = await response.json();
                if (!response.ok && response.status !== 202) {
                    throw new Error(result.message || `Sync failed with HTTP ${response.status}.`);
                }
                if (result.pending) {
                    status.textContent = result.message;
                    return;
                }
                if (result.changed) {
                    status.textContent = `Updated through ${result.latestDate || 'the latest report'}. Reloading…`;
                    window.setTimeout(() => window.location.reload(), 1200);
                    return;
                }
                status.textContent = result.message || `Already up to date through ${result.latestDate || 'the latest report'}.`;
            } catch (error) {
                console.error('Could not sync KPI data:', error);
                status.textContent = error instanceof TypeError
                    ? 'Unable to reach sync service. Start the dashboard with python serve.py and try again.'
                    : error.message || 'Sync failed.';
            } finally {
                button.disabled = false;
                buttonLabel.textContent = 'SYNC';
            }
        });
    }

    function renderInventorySummaryCards(totals, siteCount) {
        const container = document.getElementById('inventoryTotalCards');
        const table = document.querySelector('.inventory-table');
        const headerCells = table?.tHead?.rows[0]?.cells;
        if (!container || !headerCells) return;

        container.innerHTML = totals.map((total, index) => {
            const name = headerCells[index + 2]?.textContent.trim();
            if (!name) return '';
            return `
                <article class="kpi-card inventory-total-card">
                    <div class="kpi-card-header">
                        <div class="kpi-title">${name}</div>
                        <span class="status-badge good">TOTAL</span>
                    </div>
                    <div class="kpi-value-row">
                        <span class="kpi-value">${total}</span>
                    </div>
                    <div class="kpi-footer">
                        <div class="kpi-target">Across ${siteCount} ${siteCount === 1 ? 'site' : 'sites'}</div>
                    </div>
                </article>
            `;
        }).join('');
    }

    function setupInventorySiteFilter() {
        const siteSearch = document.getElementById('inventorySiteFilter');
        const siteOptions = document.getElementById('inventorySiteOptions');
        const clearFilterButton = document.getElementById('clearInventorySiteFilterBtn');
        const status = document.getElementById('inventoryFilterStatus');
        const tbody = document.getElementById('inventoryTableBody');
        const ipTbody = document.getElementById('inventoryIpTableBody');
        const table = document.querySelector('.inventory-table');
        const totalsRow = table?.tFoot?.rows[0];
        const headerCells = table?.tHead?.rows[0]?.cells;
        if (!siteSearch || !siteOptions || !clearFilterButton || !status || !tbody || !ipTbody || !totalsRow || !headerCells) return;

        const inventoryStorageKey = 'lte_dashboard_inventory_sites';
        const pinStorageKey = 'lte_dashboard_inventory_pin';
        const assetFieldIds = [
            'inventoryEditPlatform', 'inventoryEditAcoc', 'inventoryEditAsoe',
            'inventoryEditAsib', 'inventoryEditAbia', 'inventoryEditAbio',
            'inventoryEditAzha', 'inventoryEditAheb', 'inventoryEditAzna'
        ];
        const pinDialog = document.getElementById('inventoryPinDialog');
        const pinForm = document.getElementById('inventoryPinForm');
        const pinInput = document.getElementById('inventoryPinInput');
        const pinConfirm = document.getElementById('inventoryPinConfirm');
        const pinConfirmField = document.getElementById('inventoryPinConfirmField');
        const pinTitle = document.getElementById('inventoryPinTitle');
        const pinHelp = document.getElementById('inventoryPinHelp');
        const pinLabel = document.getElementById('inventoryPinLabel');
        const pinSubmit = document.getElementById('inventoryPinSubmit');
        const pinError = document.getElementById('inventoryPinError');
        const editorDialog = document.getElementById('inventoryEditorDialog');
        const editorForm = document.getElementById('inventoryEditorForm');
        const editorTitle = document.getElementById('inventoryEditorTitle');
        const editorError = document.getElementById('inventoryEditorError');
        const deleteDialog = document.getElementById('inventoryDeleteDialog');
        const deleteForm = document.getElementById('inventoryDeleteForm');
        const deleteMessage = document.getElementById('inventoryDeleteMessage');
        const deleteError = document.getElementById('inventoryDeleteError');
        const addButton = document.getElementById('addInventorySiteBtn');
        const editButton = document.getElementById('editInventorySiteBtn');
        const deleteButton = document.getElementById('deleteInventorySiteBtn');
        if (!pinDialog || !pinForm || !pinInput || !pinConfirm || !pinConfirmField ||
            !pinTitle || !pinHelp || !pinLabel || !pinSubmit || !pinError ||
            !editorDialog || !editorForm || !editorTitle || !editorError ||
            !deleteDialog || !deleteForm || !deleteMessage || !deleteError || !addButton || !editButton || !deleteButton) return;

        const readIpAddress = (cell) => cell?.querySelector('a')?.textContent.trim() || cell?.textContent.trim() || '';
        let records = Array.from(tbody.rows).map(row => {
            const site = row.cells[1]?.textContent.trim() || '';
            const ipRow = Array.from(ipTbody.rows).find(ip => ip.cells[0]?.textContent.trim() === site);
            return {
                site,
                assets: Array.from(row.cells).slice(2).map(cell => Number(cell.textContent.trim()) || 0),
                mrbtsIp: readIpAddress(ipRow?.cells[1]),
                switchIp: readIpAddress(ipRow?.cells[2])
            };
        });
        const isValidIp = value => {
            const parts = value.split('.');
            return parts.length === 4 && parts.every(part => /^\d{1,3}$/.test(part) && Number(part) <= 255);
        };
        const getSelectedSite = () => records.find(record =>
            record.site.toLocaleLowerCase() === siteSearch.value.trim().toLocaleLowerCase()
        )?.site || '';

        try {
            const savedRecords = JSON.parse(localStorage.getItem(inventoryStorageKey) || '[]');
            if (!Array.isArray(savedRecords)) throw new Error('Saved site inventory must be a list.');
            const savedSites = new Set();
            savedRecords.forEach(record => {
                if (!record || typeof record.site !== 'string' || !record.site.trim() ||
                    !Array.isArray(record.assets) || record.assets.length !== assetFieldIds.length ||
                    !record.assets.every(value => Number.isInteger(value) && value >= 0) ||
                    !isValidIp(record.mrbtsIp) || !isValidIp(record.switchIp)) {
                    throw new Error('Saved site inventory contains an invalid entry.');
                }
                if (savedSites.has(record.site.toLocaleLowerCase())) {
                    throw new Error('Saved site inventory contains duplicate site names.');
                }
                savedSites.add(record.site.toLocaleLowerCase());
            });
            if (savedRecords.length) records = savedRecords;
        } catch (error) {
            console.error('Could not load saved site inventory:', error);
            status.textContent = 'Saved site changes could not be loaded. Check this browser’s stored inventory data.';
        }

        const makeCell = (value) => {
            const cell = document.createElement('td');
            cell.textContent = String(value);
            return cell;
        };
        const makeIpLink = (ip) => {
            const link = document.createElement('a');
            link.href = `https://${ip}/`;
            link.target = '_blank';
            link.rel = 'noopener noreferrer';
            link.textContent = ip;
            return link;
        };

        const render = () => {
            const query = siteSearch.value.trim();
            const normalizedQuery = query.toLocaleLowerCase();
            const selectedSite = getSelectedSite();
            if (selectedSite) {
                document.body.classList.add('inventory-show-assets', 'inventory-show-ip');
            }

            tbody.replaceChildren();
            ipTbody.replaceChildren();
            const visibleRecords = [];
            records.forEach((record, index) => {
                const matches = !normalizedQuery || record.site.toLocaleLowerCase().includes(normalizedQuery);
                const assetRow = document.createElement('tr');
                assetRow.hidden = !matches;
                assetRow.appendChild(makeCell(index + 1));
                assetRow.appendChild(makeCell(record.site));
                record.assets.forEach(value => assetRow.appendChild(makeCell(value)));
                tbody.appendChild(assetRow);

                const ipRow = document.createElement('tr');
                ipRow.hidden = !matches;
                ipRow.appendChild(makeCell(record.site));
                const mrbtsCell = document.createElement('td');
                mrbtsCell.appendChild(makeIpLink(record.mrbtsIp));
                ipRow.appendChild(mrbtsCell);
                const switchCell = document.createElement('td');
                switchCell.appendChild(makeIpLink(record.switchIp));
                ipRow.appendChild(switchCell);
                ipTbody.appendChild(ipRow);

                if (matches) visibleRecords.push(record);
            });

            const currentSiteOptions = Array.from(siteOptions.options, option => option.value);
            if (currentSiteOptions.length !== records.length ||
                records.some((record, index) => currentSiteOptions[index] !== record.site)) {
                siteOptions.replaceChildren();
                records.forEach(record => {
                    const option = document.createElement('option');
                    option.value = record.site;
                    siteOptions.appendChild(option);
                });
            }
            clearFilterButton.disabled = !query;
            editButton.disabled = !selectedSite;
            deleteButton.disabled = !selectedSite;

            const totals = Array.from({ length: headerCells.length - 2 }, (_, index) =>
                records.reduce((sum, record) => sum + Number(record.assets[index] || 0), 0)
            );
            totalsRow.cells[0].textContent = 'Grand Total';
            totals.forEach((total, index) => {
                if (totalsRow.cells[index + 1]) totalsRow.cells[index + 1].textContent = String(total);
            });
            renderInventorySummaryCards(totals, records.length);
            status.textContent = !query
                ? `Showing all ${records.length} sites.`
                : selectedSite
                    ? `Showing ${selectedSite} only (1 of ${records.length} sites).`
                    : visibleRecords.length
                        ? `Showing ${visibleRecords.length} of ${records.length} sites matching "${query}".`
                        : `No sites match "${query}".`;
        };

        let pendingMode = 'add';
        let editingSite = '';
        const openEditor = (mode) => {
            pendingMode = mode;
            editingSite = mode === 'edit' ? getSelectedSite() : '';
            const record = records.find(item => item.site === editingSite);
            editorTitle.textContent = mode === 'edit' ? 'Edit Site' : 'Add Site';
            editorError.textContent = '';
            document.getElementById('inventoryEditSite').value = record?.site || '';
            assetFieldIds.forEach((id, index) => {
                document.getElementById(id).value = String(record?.assets[index] ?? 0);
            });
            document.getElementById('inventoryEditMrbtsIp').value = record?.mrbtsIp || '';
            document.getElementById('inventoryEditSwitchIp').value = record?.switchIp || '';
            editorDialog.showModal();
            document.getElementById('inventoryEditSite').focus();
        };

        const digestPin = async (pin, salt) => {
            if (!window.crypto?.subtle) throw new Error('PIN protection requires a secure browser context such as localhost or HTTPS.');
            const bytes = new TextEncoder().encode(`${salt}:${pin}`);
            const digest = await window.crypto.subtle.digest('SHA-256', bytes);
            return Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
        };
        const readPinConfiguration = () => {
            const savedPin = localStorage.getItem(pinStorageKey);
            if (!savedPin) return null;
            const configuration = JSON.parse(savedPin);
            if (!configuration.salt || !configuration.hash) throw new Error('The saved inventory PIN configuration is invalid.');
            return configuration;
        };

        const openPinPrompt = (mode) => {
            pendingMode = mode;
            pinError.textContent = '';
            pinInput.value = '';
            pinConfirm.value = '';
            try {
                const configured = readPinConfiguration();
                pinConfirmField.hidden = Boolean(configured);
                pinConfirm.required = !configured;
                pinInput.autocomplete = configured ? 'current-password' : 'new-password';
                pinTitle.textContent = configured ? 'Verify Inventory PIN' : 'Create Inventory PIN';
                pinLabel.textContent = configured ? 'PIN' : 'Create a 4–12 digit PIN';
                pinHelp.textContent = configured
                    ? 'Enter your PIN to manage site inventory.'
                    : 'Set a PIN before adding or editing site inventory. This PIN is stored only in this browser.';
                pinSubmit.textContent = configured ? 'Continue' : 'Set PIN & Continue';
                pinDialog.showModal();
                pinInput.focus();
            } catch (error) {
                console.error('Could not read inventory PIN:', error);
                status.textContent = error.message;
            }
        };

        addButton.addEventListener('click', () => openPinPrompt('add'));
        editButton.addEventListener('click', () => {
            if (getSelectedSite()) openPinPrompt('edit');
        });
        deleteButton.addEventListener('click', () => {
            if (getSelectedSite()) openPinPrompt('delete');
        });
        siteSearch.addEventListener('input', render);
        siteSearch.addEventListener('change', render);
        clearFilterButton.addEventListener('click', () => {
            siteSearch.value = '';
            render();
            siteSearch.focus();
        });
        pinForm.addEventListener('submit', async event => {
            event.preventDefault();
            pinError.textContent = '';
            const pin = pinInput.value;
            if (!/^\d{4,12}$/.test(pin)) {
                pinError.textContent = 'Enter a PIN containing 4 to 12 digits.';
                pinInput.focus();
                return;
            }
            try {
                const configuration = readPinConfiguration();
                if (!configuration) {
                    if (pin !== pinConfirm.value) {
                        pinError.textContent = 'The PIN entries do not match.';
                        pinConfirm.focus();
                        return;
                    }
                    const saltBytes = window.crypto.getRandomValues(new Uint8Array(16));
                    const salt = Array.from(saltBytes, byte => byte.toString(16).padStart(2, '0')).join('');
                    const hash = await digestPin(pin, salt);
                    localStorage.setItem(pinStorageKey, JSON.stringify({ salt, hash }));
                } else if (await digestPin(pin, configuration.salt) !== configuration.hash) {
                    pinError.textContent = 'Incorrect PIN. Please try again.';
                    pinInput.value = '';
                    pinInput.focus();
                    return;
                }
                pinDialog.close();
                if (pendingMode === 'delete') {
                    deleteError.textContent = '';
                    deleteMessage.textContent = `Delete "${getSelectedSite()}" from both the Site Asset and IP Address tables? This cannot be undone.`;
                    deleteDialog.showModal();
                } else {
                    openEditor(pendingMode);
                }
            } catch (error) {
                console.error('Could not verify inventory PIN:', error);
                pinError.textContent = error.message || 'The PIN could not be verified.';
            }
        });

        editorForm.addEventListener('submit', event => {
            event.preventDefault();
            editorError.textContent = '';
            const site = document.getElementById('inventoryEditSite').value.trim();
            const assetInputs = assetFieldIds.map(id => document.getElementById(id));
            const assets = assetInputs.map(input => Number(input.value));
            const mrbtsIp = document.getElementById('inventoryEditMrbtsIp').value.trim();
            const switchIp = document.getElementById('inventoryEditSwitchIp').value.trim();
            if (!site || assetInputs.some(input => !input.value || !input.checkValidity()) ||
                !isValidIp(mrbtsIp) || !isValidIp(switchIp)) {
                editorError.textContent = 'Enter a site name, non-negative whole-number hardware counts, and valid IPv4 addresses.';
                return;
            }
            if (records.some(record => record.site.toLocaleLowerCase() === site.toLocaleLowerCase() && record.site !== editingSite)) {
                editorError.textContent = 'A site with this name already exists.';
                return;
            }

            const updatedRecord = { site, assets, mrbtsIp, switchIp };
            const updatedRecords = pendingMode === 'edit'
                ? records.map(record => record.site === editingSite ? updatedRecord : record)
                : [...records, updatedRecord];
            try {
                localStorage.setItem(inventoryStorageKey, JSON.stringify(updatedRecords));
            } catch (error) {
                console.error('Could not save site inventory:', error);
                editorError.textContent = 'Site changes could not be saved in this browser. Check available storage and try again.';
                return;
            }
            records = updatedRecords;
            siteSearch.value = site;
            document.body.classList.add('inventory-show-assets', 'inventory-show-ip');
            render();
            editorDialog.close();
        });
        deleteForm.addEventListener('submit', event => {
            event.preventDefault();
            deleteError.textContent = '';
            const siteToDelete = getSelectedSite();
            if (!siteToDelete || !records.some(record => record.site === siteToDelete)) {
                deleteError.textContent = 'Select a site to delete, then try again.';
                return;
            }
            const updatedRecords = records.filter(record => record.site !== siteToDelete);
            try {
                localStorage.setItem(inventoryStorageKey, JSON.stringify(updatedRecords));
            } catch (error) {
                console.error('Could not save site deletion:', error);
                deleteError.textContent = 'The site could not be deleted from this browser. Check available storage and try again.';
                return;
            }

            records = updatedRecords;
            siteSearch.value = '';
            render();
            deleteDialog.close();
        });
        document.querySelectorAll('[data-close-dialog]').forEach(button => {
            button.addEventListener('click', () => document.getElementById(button.dataset.closeDialog)?.close());
        });
        [pinDialog, editorDialog, deleteDialog].forEach(dialog => {
            dialog.addEventListener('click', event => {
                if (event.target === dialog) dialog.close();
            });
        });

        render();
    }

    function startDataRefreshMonitor() {
        if (!['localhost', '127.0.0.1'].includes(window.location.hostname)) return;

        let currentVersion = window.KPIDataSources ? window.KPIDataSources.updatedAt : null;
        const checkForUpdates = async () => {
            try {
                const response = await fetch('/api/data-version', { cache: 'no-store' });
                if (!response.ok) throw new Error(`Update check failed with HTTP ${response.status}`);
                const version = await response.json();
                if (currentVersion === null) {
                    currentVersion = version.updatedAt;
                } else if (version.updatedAt && currentVersion !== version.updatedAt) {
                    window.location.reload();
                }
            } catch (error) {
                console.error('Could not check for newer KPI reports:', error);
            }
        };

        checkForUpdates();
        window.setInterval(checkForUpdates, 30000);
    }

    function reportSourceName(mode, fallback) {
        const sources = window.KPIDataSources;
        const sourceName = sources && (mode === 'hourly' ? sources.hourlyFile : sources.daywiseFile);
        return sourceName || fallback;
    }

    function getPeriodKey(date, mode) {
        if (mode === 'weekly') {
            const [year, month, day] = String(date).split('-').map(Number);
            const start = new Date(Date.UTC(year, month - 1, day));
            start.setUTCDate(start.getUTCDate() - (start.getUTCDay() + 6) % 7);
            return `WEEK:${start.toISOString().slice(0, 10)}`;
        }
        if (mode === 'monthly') return `MONTH:${String(date).slice(0, 7)}`;
        return String(date);
    }

    function getPeriodLabel(periodKey, mode) {
        if (mode === 'weekly') {
            const start = periodKey.slice(5);
            const [year, month, day] = start.split('-').map(Number);
            const end = new Date(Date.UTC(year, month - 1, day + 6));
            return `Week of ${formatChartDate(start)}-${String(year).slice(-2)} (${formatChartDate(end.toISOString().slice(0, 10))})`;
        }
        if (mode === 'monthly') {
            const [year, month] = periodKey.slice(6).split('-').map(Number);
            return new Date(Date.UTC(year, month - 1, 1)).toLocaleDateString('en-US', {
                month: 'long',
                year: 'numeric',
                timeZone: 'UTC'
            });
        }
        return formatChartDate(periodKey);
    }

    function getPeriodKeys(dates, mode) {
        return [...new Set(dates.filter(Boolean).map(date => getPeriodKey(date, mode)))].sort();
    }

    function matchesSelectedPeriod(date, selectedPeriod, mode) {
        if (selectedPeriod === 'ALL') return true;
        if (mode === 'weekly' || mode === 'monthly') {
            return getPeriodKey(date, mode) === selectedPeriod;
        }
        return date === selectedPeriod;
    }

    // Theme Management
    function applyTheme(theme) {
        state.activeTheme = theme;
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('lte_dashboard_theme', theme);
        if (elements.themeToggleBtn) {
            elements.themeToggleBtn.innerHTML = theme === 'light' ? '🌙 Dark Mode' : '☀️ Light Mode';
        }
        updateChartTheme();
    }

    function toggleTheme() {
        applyTheme(state.activeTheme === 'dark' ? 'light' : 'dark');
    }

    function setViewMode(mode) {
        state.viewMode = mode;
        const isPeriodView = mode !== 'hourly';
        const daywiseSummarySection = document.getElementById('daywiseSummarySection');
        if (daywiseSummarySection) {
            daywiseSummarySection.style.display = isPeriodView ? '' : 'none';
        }
        const periodSummaryTitle = document.getElementById('periodSummaryTitle');
        if (periodSummaryTitle) {
            periodSummaryTitle.textContent = mode === 'weekly' ? 'Weekly KPI Summary'
                : mode === 'monthly' ? 'Monthly KPI Summary'
                : '8-Day KPI Summary';
        }
        if (elements.viewModeContainer) {
            elements.viewModeContainer.querySelectorAll('.view-mode-btn').forEach(button => {
                button.classList.toggle('active', button.dataset.viewMode === mode);
            });
        }

        if (isPeriodView) {
            if (!state.daywiseRecords.length && window.DaywiseLTEData) {
                state.daywiseRecords = unpackCompactData(window.DaywiseLTEData);
            }
            if (state.daywiseRecords.length) {
                const dates = state.daywiseRecords.map(record => record.Date).filter(Boolean).sort();
                state.filters.date = mode === 'daily' ? 'ALL'
                    : dates.length ? getPeriodKey(dates[dates.length - 1], mode) : 'ALL';
                state.filters.hour = 'ALL';
                state.filters.busyHourOnly = false;
                if (elements.filterBusyHour) elements.filterBusyHour.checked = false;
                loadProcessedRecords(
                    state.daywiseRecords,
                    reportSourceName(mode, 'Daywise KPI report'),
                    true,
                    mode
                );
            }
        } else if (state.hourlyRecords.length) {
            state.filters.date = 'ALL';
            loadProcessedRecords(
                state.hourlyRecords,
                reportSourceName('hourly', '24-Hour KPI report'),
                true,
                'hourly'
            );
        }
    }

    // Load Default Data
    function loadDefaultReportData() {
        if (window.UserLTEDataCompact) {
            console.log('Unpacking user report dataset (separate Site Name & LNCEL)...');
            const records = unpackCompactData(window.UserLTEDataCompact);
            state.hourlyRecords = records;
            loadProcessedRecords(
                records,
                reportSourceName('hourly', '24-Hour KPI report'),
                true,
                'hourly'
            );
        } else if (window.SampleLTEData) {
            const data = window.SampleLTEData.generateDataset(14).map(r => ({
                ...r,
                LNCEL: r.Cell_ID,
                Operator: 'Sample_RAN'
            }));
            loadProcessedRecords(data, 'Preloaded 14-Day Simulation Sample', false);
        }
    }

    function unpackCompactData(data) {
        const cols = data.columns;
        return data.rows.map(row => {
            const obj = {};
            cols.forEach((col, idx) => {
                obj[col] = row[idx];
            });
            return obj;
        });
    }

    // Load records into state
    function loadProcessedRecords(records, sourceName, isUserReport = true, mode = 'active') {
        const normalizedRecords = records.map(record => ({
            ...record,
            Band: resolveRecordBand(record)
        }));
        state.rawRecords = normalizedRecords;
        if (mode === 'hourly') state.hourlyRecords = normalizedRecords;
        if (mode !== 'hourly') state.daywiseRecords = normalizedRecords;
        state.dataSourceName = sourceName;
        state.isUserReport = isUserReport;
        state.pagination.currentPage = 1;

        populateFilterDropdowns();
        updateDataSourceBanner();
        applyFilters();
    }

    function updateDataSourceBanner() {
        if (elements.dataSourceTitle) {
            elements.dataSourceTitle.innerText = state.dataSourceName;
        }
        if (elements.dataSourceSubtitle) {
            const rowCount = state.filteredRecords.length.toLocaleString();
            const totalCount = state.rawRecords.length.toLocaleString();
            const opLabel = state.filters.operator === 'ALL' ? 'All Operators (AIRTEL & VIL)' : `Operator: ${state.filters.operator}`;
            const dateLabel = state.filters.date === 'ALL'
                ? state.viewMode === 'weekly' || state.viewMode === 'monthly' ? 'All Periods' : 'All Dates'
                : `${state.viewMode === 'weekly' || state.viewMode === 'monthly' ? 'Period' : 'Date'}: ${getPeriodLabel(state.filters.date, state.viewMode)}`;
            const siteCount = new Set(state.filteredRecords.map(r => r.Site_Name || r.Site_ID).filter(Boolean)).size;
            const lncelCount = new Set(state.filteredRecords.map(r => `${r.Site_Name}_${r.LNCEL}`).filter(Boolean)).size;
            elements.dataSourceSubtitle.innerText = `Filtered: ${rowCount} of ${totalCount} records | ${opLabel} | ${dateLabel} | ${siteCount} Sites | ${lncelCount} Sectors`;
        }
    }

    // Populate Dynamic Filter Selects
    function populateFilterDropdowns() {
        // Operators
        const operators = [...new Set(state.rawRecords.map(r => r.Operator).filter(Boolean))].sort();
        populateSelect(elements.filterOperator, operators, 'All Operators');

        // Render Quick-Toggle Operator Pills
        if (elements.operatorPillContainer) {
            elements.operatorPillContainer.innerHTML = '';
            const allBtn = document.createElement('button');
            allBtn.className = `operator-btn ${state.filters.operator === 'ALL' ? 'active' : ''}`;
            allBtn.innerText = 'All Operators';
            allBtn.addEventListener('click', () => setOperatorFilter('ALL'));
            elements.operatorPillContainer.appendChild(allBtn);

            operators.forEach(op => {
                const btn = document.createElement('button');
                btn.className = `operator-btn ${state.filters.operator === op ? 'active' : ''}`;
                btn.setAttribute('data-op', op);
                btn.innerText = op;
                btn.addEventListener('click', () => setOperatorFilter(op));
                elements.operatorPillContainer.appendChild(btn);
            });
        }

        // Dates and reporting periods
        const dates = [...new Set(state.rawRecords.map(r => r.Date).filter(Boolean))].sort();
        if (state.viewMode === 'weekly' || state.viewMode === 'monthly') {
            const periods = getPeriodKeys(dates, state.viewMode);
            const periodOptions = periods.map(period => ({
                value: period,
                label: getPeriodLabel(period, state.viewMode)
            }));
            if (elements.filterDate) {
                elements.filterDate.innerHTML = '<option value="ALL">All Periods</option>';
                periodOptions.forEach(period => {
                    const option = document.createElement('option');
                    option.value = period.value;
                    option.textContent = period.label;
                    elements.filterDate.appendChild(option);
                });
            }
            if (!periods.includes(state.filters.date)) {
                state.filters.date = periods[periods.length - 1] || 'ALL';
            }
            if (elements.filterDate) elements.filterDate.value = state.filters.date;
        } else {
            populateSelect(elements.filterDate, dates, 'All Dates');
            if (state.filters.date !== 'ALL' && !dates.includes(state.filters.date)) {
                state.filters.date = 'ALL';
            }
            if (state.filters.date === 'ALL' && dates.length > 0 && state.viewMode === 'hourly') {
                state.filters.date = dates[dates.length - 1];
                if (elements.filterDate) elements.filterDate.value = state.filters.date;
            }
        }
        if (elements.filterDate) {
            elements.filterDate.options[0].textContent =
                state.viewMode === 'weekly' || state.viewMode === 'monthly' ? 'All Periods' : 'All Dates';
        }
        if (elements.filterHour) {
            const isPeriodView = state.viewMode === 'weekly' || state.viewMode === 'monthly';
            elements.filterHour.disabled = isPeriodView;
            if (isPeriodView) elements.filterHour.title = 'Measurement hour is not applicable to period summaries';
            else elements.filterHour.removeAttribute('title');
        }
        if (elements.filterBusyHour) {
            const busyHourGroup = elements.filterBusyHour.closest('.toggle-filter');
            const isPeriodView = state.viewMode === 'weekly' || state.viewMode === 'monthly';
            elements.filterBusyHour.disabled = isPeriodView;
            if (busyHourGroup) busyHourGroup.title = isPeriodView
                ? 'Busy Hour filtering is not applicable to period summaries'
                : 'Filter to Busy Hour only — 18:00 HRS (Peak Network Hour)';
        }

        // Bands
        const bands = [...new Set(state.rawRecords.map(r => r.Band).filter(Boolean))].sort();
        populateSelect(elements.filterBand, bands, 'All Bands (FDD/TDD)');

        // Sites & LNCELs
        updateSiteAndLncelDropdowns();

        // Hours
        const hours = [...new Set(state.rawRecords.map(r => r.Hour).filter(h => h !== undefined && h !== null))].sort((a, b) => a - b);
        if (elements.filterHour) {
            elements.filterHour.innerHTML = '<option value="ALL">All 24 Hours</option>';
            hours.forEach(h => {
                const opt = document.createElement('option');
                opt.value = h;
                const isBh = (parseInt(h, 10) === 18);
                if (isBh) {
                    opt.textContent = `18:00 HRS ⭐ BUSY HOUR`;
                    opt.classList.add('bh-option');
                    opt.setAttribute('data-bh', 'true');
                } else {
                    opt.textContent = `${String(h).padStart(2, '0')}:00 HRS`;
                }
                elements.filterHour.appendChild(opt);
            });
            if (state.filters.hour) {
                elements.filterHour.value = state.filters.hour;
            }
        }
    }

    function setOperatorFilter(op) {
        state.filters.operator = op;
        if (elements.filterOperator) elements.filterOperator.value = op;
        updateOperatorPillActive();
        updateSiteAndLncelDropdowns();
        applyFilters();
    }

    function updateOperatorPillActive() {
        if (!elements.operatorPillContainer) return;
        const btns = elements.operatorPillContainer.querySelectorAll('.operator-btn');
        btns.forEach(btn => {
            const op = btn.getAttribute('data-op') || 'ALL';
            if (op === state.filters.operator) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });
    }

    // Updates Site and LNCEL dropdowns without concatenating them
    function updateSiteAndLncelDropdowns() {
        let pool = state.rawRecords;
        if (state.filters.operator !== 'ALL') {
            pool = pool.filter(r => r.Operator === state.filters.operator);
        }

        const sites = [...new Set(pool.map(r => r.Site_Name || r.Site_ID).filter(Boolean))].sort();
        const lncels = [...new Set(pool.map(r => r.LNCEL).filter(Boolean))].sort((a, b) => {
            const numA = parseInt(a, 10);
            const numB = parseInt(b, 10);
            return (!isNaN(numA) && !isNaN(numB)) ? numA - numB : String(a).localeCompare(String(b));
        });

        populateSelect(elements.filterSite, sites, `All Sites (${sites.length})`);
        
        if (elements.filterLncel) {
            elements.filterLncel.innerHTML = `<option value="ALL">All LNCELs (${lncels.length})</option>`;
            lncels.forEach(l => {
                const opt = document.createElement('option');
                opt.value = l;
                opt.textContent = `LNCEL: ${l}`;
                elements.filterLncel.appendChild(opt);
            });
            if (lncels.includes(state.filters.lncel)) {
                elements.filterLncel.value = state.filters.lncel;
            }
        }
    }

    function populateSelect(selectEl, items, allLabel) {
        if (!selectEl) return;
        const currentVal = selectEl.value;
        selectEl.innerHTML = `<option value="ALL">${allLabel}</option>`;
        items.forEach(item => {
            const opt = document.createElement('option');
            opt.value = item;
            opt.textContent = item.length > 40 ? item.substring(0, 38) + '...' : item;
            selectEl.appendChild(opt);
        });
        if (items.includes(currentVal)) {
            selectEl.value = currentVal;
        }
    }

    // Filter Processing
    function applyFilters() {
        const {
            operator,
            date,
            band,
            hour,
            site,
            lncel,
            busyHourOnly,
            degradedOnly,
            lowThroughputLowUtilization
        } = state.filters;
        const searchVal = (elements.tableSearch ? elements.tableSearch.value.trim().toLowerCase() : '');

        state.filteredRecords = state.rawRecords.filter(row => {
            // Operator filter
            if (operator !== 'ALL' && row.Operator !== operator) return false;

            // Date filter
            if (!matchesSelectedPeriod(row.Date, date, state.viewMode)) return false;

            // Band filter
            if (band !== 'ALL' && row.Band !== band) return false;

            // Hour filter
            if (hour !== 'ALL') {
                if (parseInt(row.Hour, 10) !== parseInt(hour, 10)) return false;
            }

            // Site filter
            if (site !== 'ALL') {
                const siteVal = row.Site_Name || row.Site_ID;
                if (siteVal !== site) return false;
            }

            // LNCEL filter (separate, clean check)
            if (lncel !== 'ALL') {
                if (String(row.LNCEL) !== String(lncel)) return false;
            }

            // Busy Hour filter — exactly 18:00 hrs
            if (busyHourOnly) {
                const h = parseInt(row.Hour, 10);
                if (isNaN(h) || h !== 18) return false;
            }

            // Degraded filter
            if (degradedOnly) {
                const isDegraded = (row.VoLTE_Drop_Rate > 0.8) || (row.Drop_Rate > 0.8) || (row.DL_User_Throughput_Mbps > 0 && row.DL_User_Throughput_Mbps < 6.0) || (row.DL_PRB_Util > 75.0);
                if (!isDegraded) return false;
            }

            // Low throughput and low utilization filter
            if (lowThroughputLowUtilization) {
                const throughput = getDlThroughputMbps(row);
                const utilization = parseFloat(row.DL_PRB_Util);
                if (!Number.isFinite(throughput) || !Number.isFinite(utilization) || throughput <= 0 || throughput >= 2 || utilization >= 60) {
                    return false;
                }
            }

            // Search filter
            if (searchVal) {
                const searchable = `${row.Site_Name || ''} ${row.Site_ID || ''} ${row.LNCEL || ''} ${row.Operator || ''} ${row.Band || ''}`.toLowerCase();
                if (!searchable.includes(searchVal)) return false;
            }

            return true;
        });

        updateDataSourceBanner();
        updateKpiScorecards();
        updateCharts();
        updateAnomalyAlert();
        renderTable();
        renderDaywiseSummary();
    }

    function renderDaywiseSummary() {
        const body = document.getElementById('daywiseSummaryBody');
        if (!body || state.viewMode === 'hourly') return;

        let recordsForSummary = state.filteredRecords;
        if (state.viewMode === 'weekly' || state.viewMode === 'monthly') {
            const searchVal = elements.tableSearch ? elements.tableSearch.value.trim().toLowerCase() : '';
            recordsForSummary = state.rawRecords.filter(record => {
                if (state.filters.operator !== 'ALL' && record.Operator !== state.filters.operator) return false;
                if (state.filters.band !== 'ALL' && record.Band !== state.filters.band) return false;
                if (state.filters.site !== 'ALL' &&
                    (record.Site_Name || record.Site_ID) !== state.filters.site) return false;
                if (state.filters.lncel !== 'ALL' &&
                    String(record.LNCEL) !== String(state.filters.lncel)) return false;
                if (searchVal) {
                    const searchable = `${record.Site_Name || ''} ${record.Site_ID || ''} ${record.LNCEL || ''} ${record.Operator || ''} ${record.Band || ''}`.toLowerCase();
                    if (!searchable.includes(searchVal)) return false;
                }
                return true;
            });
        }
        const summaryMode = state.viewMode;
        const availableDates = [...new Set(recordsForSummary.map(r => r.Date).filter(Boolean))].sort();
        const allKeys = summaryMode === 'daily'
            ? availableDates
            : getPeriodKeys(availableDates, summaryMode);
        const groups = allKeys.slice(-(summaryMode === 'monthly' ? 12 : 8)).map(key => ({
            key,
            dates: summaryMode === 'daily'
                ? [key]
                : availableDates.filter(date => getPeriodKey(date, summaryMode) === key)
        }));
        const numeric = (records, key) => {
            const values = records.map(r => Number(r[key])).filter(Number.isFinite);
            return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
        };
        const sum = (records, key) => records.reduce((total, r) => total + (Number(r[key]) || 0), 0);
        const distinctCells = records => new Set(records.map(r => `${r.Site_Name || r.Site_ID}@@${r.LNCEL}`)).size;
        const getBusyHourRecords = dates => {
            const dateSet = new Set(dates);
            const searchVal = elements.tableSearch ? elements.tableSearch.value.trim().toLowerCase() : '';
            return state.hourlyRecords.filter(record => {
                if (!dateSet.has(record.Date) || Number(record.Hour) !== 18) return false;
                if (state.filters.operator !== 'ALL' && record.Operator !== state.filters.operator) return false;
                if (state.filters.band !== 'ALL' && record.Band !== state.filters.band) return false;
                if (state.filters.site !== 'ALL' && (record.Site_Name || record.Site_ID) !== state.filters.site) return false;
                if (state.filters.lncel !== 'ALL' && String(record.LNCEL) !== String(state.filters.lncel)) return false;
                if (searchVal) {
                    const searchable = `${record.Site_Name || ''} ${record.Site_ID || ''} ${record.LNCEL || ''} ${record.Operator || ''} ${record.Band || ''}`.toLowerCase();
                    if (!searchable.includes(searchVal)) return false;
                }
                return true;
            });
        };
        const format = (value, digits = 2) => value === null ? '—' : value.toFixed(digits);
        const label = group => {
            if (summaryMode === 'daily') return `${formatChartDate(group.key)}-${group.key.slice(2, 4)}`;
            if (summaryMode === 'weekly') {
                const start = group.key.slice(5);
                const [year, month, day] = start.split('-').map(Number);
                const end = new Date(Date.UTC(year, month - 1, day + 6)).toISOString().slice(0, 10);
                return `${formatChartDate(start)}–${formatChartDate(end)}`;
            }
            return getPeriodLabel(group.key, summaryMode);
        };
        const rows = [
            ['No Of cells — LNCEL', records => distinctCells(records), 0],
            ['Cell Availability LTE-4G_KPI_Report', records => numeric(records, 'Cell_Availability'), 2],
            ['NSN_Data Volume - Total_GB', records => sum(records, 'Total_Traffic_GB'), 2],
            ['Nokia_LTE_VoLTE Traffic Erl', records => sum(records, 'VoLTE_Traffic_Erl'), 2],
            ['Nokia_LTE_VoLTE Call Setup Success Rate', records => numeric(records, 'VoLTE_CSSR'), 2],
            ['E2E Call Setup Success Rate - All Bearer New', records => numeric(records, 'E2E_CSSR'), 2],
            ['NSN_LTE_VoLTE_Drop_Call_Rate', records => numeric(records, 'VoLTE_Drop_Rate'), 2],
            ['NSN_LTE_DL Throughput Per UE (NA)', records => numeric(records, 'DL_User_Throughput_Mbps'), 2],
            ['NSN_LTE_UL Throughput Per UE (NA)', records => numeric(records, 'UL_User_Throughput_Mbps'), 2],
            ['No Of cells DCR >2%', records => distinctCells(records.filter(r => Number(r.Drop_Rate) > 2)), 0],
            ['No Of cells Throughput <1.5 mbps_24 hrs', records => distinctCells(records.filter(r => getDlThroughputMbps(r) > 0 && getDlThroughputMbps(r) < 1.5)), 0],
            ['No Of cells Throughput <1.5 mbps_BBH', (_records, dates) => {
                const busyHourRecords = getBusyHourRecords(dates);
                if (!busyHourRecords.length) return null;
                return distinctCells(busyHourRecords.filter(r => getDlThroughputMbps(r) > 0 && getDlThroughputMbps(r) < 1.5));
            }, 0]
        ];

        const classifySummaryValue = (name, value) => {
            if (name === 'Cell Availability LTE-4G_KPI_Report') {
                return value >= 99.5 ? 'good' : (value >= 98.0 ? 'warning' : 'critical');
            }
            if (name === 'Nokia_LTE_VoLTE Call Setup Success Rate' ||
                name === 'E2E Call Setup Success Rate - All ******') {
                return value >= 99.0 ? 'good' : (value >= 98.0 ? 'warning' : 'critical');
            }
            if (name === 'NSN_LTE_VoLTE_Drop_Call_Rate') {
                return value <= 1.0 ? 'good' : (value <= 1.5 ? 'warning' : 'critical');
            }
            if (name === 'NSN_LTE_DL Throughput Per UE (NA)') {
                const throughputMbps = value / 1000;
                return throughputMbps >= 3.0 ? 'good' : (throughputMbps >= 2.0 ? 'warning' : 'critical');
            }
            if (name === 'NSN_LTE_UL Throughput Per UE (NA)') {
                return value < 200 ? 'critical' : (value > 500 ? 'good' : 'warning');
            }
            if (name === 'No Of cells DCR >2%' ||
                name === 'No Of cells Throughput <1.5 mbps_24 hrs' ||
                name === 'No Of cells Throughput <1.5 mbps_BBH') {
                return value === 0 ? 'good' : 'critical';
            }
            return '';
        };

        body.innerHTML = rows.map(([name, calculate, digits]) => {
            const values = groups.map(group => {
                const groupDates = new Set(group.dates);
                const value = calculate(recordsForSummary.filter(r => groupDates.has(r.Date)), group.dates);
                if (value === null) return '<td>—</td>';
                const displayValue = digits === 0 ? value : format(value, digits);
                const status = Number.isFinite(value) ? classifySummaryValue(name, value) : '';
                return status
                    ? `<td><span class="status-badge ${status}" aria-label="${status}">${displayValue}</span></td>`
                    : `<td>${displayValue}</td>`;
            }).join('');
            return `<tr><th>${name}</th>${values}</tr>`;
        }).join('');

        const header = body.parentElement.querySelector('thead tr');
        if (header) {
            header.innerHTML = `<th>Values</th>${groups.map(group => `<th>${label(group)}</th>`).join('')}`;
        }
        const summarySubtitle = document.querySelector('#daywiseSummarySection .table-header-row p');
        if (summarySubtitle && groups.length) {
            if (summaryMode === 'weekly') {
                summarySubtitle.textContent = `Daily KPI data grouped by calendar week (Monday–Sunday), through ${label(groups[groups.length - 1])}`;
            } else if (summaryMode === 'monthly') {
                summarySubtitle.textContent = `Daily KPI data grouped by month, through ${label(groups[groups.length - 1])}`;
            } else {
                summarySubtitle.textContent = `Day-wise dashboard summary from ${label(groups[0])} through ${label(groups[groups.length - 1])}`;
            }
        }
    }

    // Metric Calculations
    function getMetricAvg(records, key) {
        const valid = records.map(r => parseFloat(r[key])).filter(v => !isNaN(v) && v !== null);
        if (valid.length === 0) return 0;
        return valid.reduce((a, b) => a + b, 0) / valid.length;
    }

    function getMetricSum(records, key) {
        return records.map(r => parseFloat(r[key])).filter(v => !isNaN(v) && v !== null).reduce((a, b) => a + b, 0);
    }

    // Update Scorecard Cards
    function updateKpiScorecards() {
        const recs = state.filteredRecords;
        const currentOp = state.filters.operator;
        const opTag = currentOp === 'ALL' ? 'Combined' : currentOp;

        if (recs.length === 0) {
            setCard('kpiVolteCssr', '0.00', 'critical');
            setCard('kpiRrcSr', '0.00', 'critical');
            setCard('kpiVolteDrop', '0.00', 'critical');
            setCard('kpiServiceDrop', '0.00', 'critical');
            setCard('kpiDlThrpt', '0.00', 'critical');
            setCard('kpiTotalTraffic', '0.0', 'good');
            setCard('kpiVolteTraffic', '0.0', 'good');
            setCard('kpiPrbUtil', '0.0', 'good');
            return;
        }

        const avgVolteCssr = getMetricAvg(recs, 'VoLTE_CSSR');
        const avgRrcSr = getMetricAvg(recs, 'RRC_Setup_SR');
        const avgVolteDrop = getMetricAvg(recs, 'VoLTE_Drop_Rate');
        const avgServiceDrop = getMetricAvg(recs, 'Drop_Rate');
        const avgDlThrpt = getMetricAvg(recs, 'DL_User_Throughput_Mbps') / 1000;
        const totalTrafficGb = getMetricSum(recs, 'Total_Traffic_GB');
        const totalVolteErl = getMetricSum(recs, 'VoLTE_Traffic_Erl');
        const avgDlPrb = getMetricAvg(recs, 'DL_PRB_Util');
        const avgAvail = getMetricAvg(recs, 'Cell_Availability');
        const avgCqi = getMetricAvg(recs, 'CQI');
        const avgSinr = getMetricAvg(recs, 'SINR');

        // VoLTE Call Setup SR
        setCard('kpiVolteCssr', avgVolteCssr.toFixed(2), avgVolteCssr >= 99.0 ? 'good' : (avgVolteCssr >= 98.0 ? 'warning' : 'critical'));
        
        // RRC Setup SR
        setCard('kpiRrcSr', avgRrcSr.toFixed(2), avgRrcSr >= 99.0 ? 'good' : (avgRrcSr >= 98.0 ? 'warning' : 'critical'));

        // VoLTE Drop Rate
        setCard('kpiVolteDrop', avgVolteDrop.toFixed(2), avgVolteDrop <= 1.0 ? 'good' : (avgVolteDrop <= 1.5 ? 'warning' : 'critical'));

        // Service Drop Rate
        setCard('kpiServiceDrop', avgServiceDrop.toFixed(2), avgServiceDrop <= 0.8 ? 'good' : (avgServiceDrop <= 1.5 ? 'warning' : 'critical'));

        // DL User Throughput
        setCard('kpiDlThrpt', avgDlThrpt.toFixed(2), avgDlThrpt >= 3.0 ? 'good' : (avgDlThrpt >= 2.0 ? 'warning' : 'critical'));

        // Total Data Traffic
        if (totalTrafficGb >= 1000) {
            setCard('kpiTotalTraffic', (totalTrafficGb / 1000).toFixed(2), 'good', 'TB');
        } else {
            setCard('kpiTotalTraffic', totalTrafficGb.toFixed(1), 'good', 'GB');
        }

        // VoLTE Voice Traffic
        setCard('kpiVolteTraffic', totalVolteErl.toFixed(1), 'good', 'Erl');

        // DL PRB Utilization %
        setCard('kpiPrbUtil', avgDlPrb.toFixed(1), avgDlPrb < 75.0 ? 'good' : (avgDlPrb < 85.0 ? 'warning' : 'critical'));

        // Cell Availability %
        setCard('kpiAvail', avgAvail.toFixed(2), avgAvail >= 99.5 ? 'good' : (avgAvail >= 98.0 ? 'warning' : 'critical'));

        // CQI & SINR
        setCard('kpiCqi', avgCqi.toFixed(1), avgCqi >= 8.0 ? 'good' : 'warning');
        setCard('kpiSinr', avgSinr.toFixed(1), avgSinr >= 8.0 ? 'good' : 'warning');

        updateCardOperatorTags(opTag);
    }

    function setCard(cardId, value, status, customUnit) {
        const cardEl = document.getElementById(cardId);
        if (!cardEl) return;
        const valEl = cardEl.querySelector('.kpi-value');
        const badgeEl = cardEl.querySelector('.status-badge');
        const unitEl = cardEl.querySelector('.kpi-unit');

        if (valEl) valEl.innerText = value;
        if (customUnit && unitEl) unitEl.innerText = customUnit;

        if (badgeEl) {
            badgeEl.className = `status-badge ${status}`;
            badgeEl.innerText = status === 'good' ? 'OPTIMAL' : (status === 'warning' ? 'WARNING' : 'CRITICAL');
        }
    }

    function updateCardOperatorTags(opTag) {
        const indicators = document.querySelectorAll('.trend-indicator');
        indicators.forEach(ind => {
            const currentText = ind.innerText;
            if (!ind.hasAttribute('data-original')) {
                ind.setAttribute('data-original', currentText);
            }
            const orig = ind.getAttribute('data-original');
            ind.innerText = `${orig} • [${opTag}]`;
        });
    }

    // Anomaly Alert Banner (Clean Site and LNCEL representation)
    function updateAnomalyAlert() {
        if (!elements.alertBanner || !elements.alertText) return;

        let alertRecords = state.filteredRecords;
        if (state.filters.date === 'ALL' && !state.filters.alertDateFilterTouched) {
            const latestDate = state.rawRecords
                .map(r => r.Date)
                .filter(Boolean)
                .sort()
                .pop();
            if (latestDate) {
                alertRecords = alertRecords.filter(r => r.Date === latestDate);
            }
        }

        const cellDrops = {};
        alertRecords.forEach(r => {
            const sectorKey = `${r.Site_Name || r.Site_ID}_@@_${r.LNCEL}`;
            if (!cellDrops[sectorKey]) {
                cellDrops[sectorKey] = { count: 0, dropSum: 0, prbSum: 0, site: r.Site_Name || r.Site_ID, lncel: r.LNCEL, op: r.Operator };
            }
            const m = cellDrops[sectorKey];
            m.count++;
            m.dropSum += (r.VoLTE_Drop_Rate || r.Drop_Rate || 0);
            m.prbSum += (r.DL_PRB_Util || 0);
        });

        const degraded = [];
        Object.keys(cellDrops).forEach(k => {
            const m = cellDrops[k];
            const avgDrop = m.dropSum / m.count;
            const avgPrb = m.prbSum / m.count;
            if (avgDrop > 1.0) {
                degraded.push(`${m.site} (LNCEL ${m.lncel}) [${m.op} Drop: ${avgDrop.toFixed(2)}%]`);
            } else if (avgPrb > 80.0) {
                degraded.push(`${m.site} (LNCEL ${m.lncel}) [${m.op} PRB: ${avgPrb.toFixed(1)}%]`);
            }
        });

        if (degraded.length > 0) {
            elements.alertBanner.style.display = 'flex';
            elements.alertText.innerHTML = `<strong>${degraded.length} Degradation Alerts Detected:</strong> Issues on ${degraded.slice(0, 3).join(', ')}${degraded.length > 3 ? ` and ${degraded.length - 3} other sectors` : ''}.`;
        } else {
            elements.alertBanner.style.display = 'none';
        }
    }

    // Charts Engine (Chart.js)
    function updateCharts() {
        if (typeof Chart === 'undefined') return;
        renderThroughputPrbTrendChart();
        renderWorstCellsChart();
        renderLowThroughputCellsChart();
        renderHourlyTrafficProfileChart();
        renderVolteQualityTrendChart();
        renderSectorDistributionChart();
        renderDailyVolumeChart();
    }

    function getChartThemeColors() {
        const isDark = state.activeTheme === 'dark';
        return {
            textColor: isDark ? '#94a3b8' : '#475569',
            gridColor: isDark ? 'rgba(255, 255, 255, 0.06)' : 'rgba(0, 0, 0, 0.06)',
            tooltipBg: isDark ? 'rgba(17, 23, 40, 0.95)' : 'rgba(255, 255, 255, 0.95)',
            tooltipText: isDark ? '#f1f5f9' : '#0f172a'
        };
    }

    function formatChartDate(dateValue) {
        const parts = String(dateValue).split(/[-\/]/);
        if (parts.length === 3 && parts[0].length === 4) {
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            return `${parseInt(parts[2], 10)}-${months[parseInt(parts[1], 10) - 1] || parts[1]}`;
        }
        return String(dateValue);
    }

    function updateChartTheme() {
        if (Object.keys(state.charts).length === 0) return;
        updateCharts();
    }

    function getDlThroughputMbps(row) {
        const rawThroughput = parseFloat(row.DL_User_Throughput_Mbps);
        return Number.isFinite(rawThroughput) ? rawThroughput / 1000 : NaN;
    }

    function renderSectorDistributionChart() {
        const canvas = document.getElementById('chartSectorDistribution');
        const subtitle = document.getElementById('sectorDistributionSubtitle');
        if (!canvas) return;

        const sectorSets = new Map();
        const bandSet = new Set();
        state.filteredRecords.forEach(record => {
            const operator = record.Operator || 'Unknown';
            const band = record.Band || 'Unknown';
            const site = record.Site_ID || record.Site_Name || 'Unknown Site';
            const sector = `${site}@@${record.LNCEL || 'Unknown LNCEL'}`;
            bandSet.add(band);
            if (!sectorSets.has(operator)) sectorSets.set(operator, new Map());
            const operatorBands = sectorSets.get(operator);
            if (!operatorBands.has(band)) operatorBands.set(band, new Set());
            operatorBands.get(band).add(sector);
        });

        const operators = [...sectorSets.keys()].sort();
        const bands = [...bandSet].sort((a, b) => {
            const preferred = ['FDD', 'TDD'];
            const indexA = preferred.indexOf(a);
            const indexB = preferred.indexOf(b);
            if (indexA !== -1 || indexB !== -1) {
                return (indexA === -1 ? preferred.length : indexA) - (indexB === -1 ? preferred.length : indexB);
            }
            return a.localeCompare(b);
        });
        const countsByBand = Object.fromEntries(bands.map(band => [
            band,
            operators.map(operator => sectorSets.get(operator).get(band)?.size || 0)
        ]));
        const operatorTotals = operators.map((_, index) =>
            bands.reduce((sum, band) => sum + countsByBand[band][index], 0)
        );
        const total = operatorTotals.reduce((sum, count) => sum + count, 0);
        if (subtitle) subtitle.textContent = 'Sector distribution by operator and band';

        if (state.charts.sectorDistribution) {
            state.charts.sectorDistribution.destroy();
            delete state.charts.sectorDistribution;
        }
        if (!total) return;

        const theme = getChartThemeColors();
        const colors = {
            FDD: { fill: 'rgba(79, 112, 199, 0.92)', border: '#4f70c7', text: '#ffffff' },
            TDD: { fill: 'rgba(139, 201, 107, 0.92)', border: '#8bc96b', text: '#102018' }
        };
        const fallback = [
            { fill: 'rgba(240, 190, 75, 0.92)', border: '#e6b843', text: '#17202a' },
            { fill: 'rgba(225, 92, 99, 0.92)', border: '#d95c63', text: '#ffffff' },
            { fill: 'rgba(89, 184, 211, 0.92)', border: '#59b8d3', text: '#10202a' }
        ];
        const bandColors = Object.fromEntries(bands.map((band, index) => [
            band,
            colors[band] || fallback[index % fallback.length]
        ]));
        const segmentLabelsPlugin = {
            id: 'sectorDistributionSegmentLabels',
            afterDatasetsDraw(chart) {
                const context = chart.ctx;
                chart.data.datasets.forEach((dataset, datasetIndex) => {
                    const meta = chart.getDatasetMeta(datasetIndex);
                    meta.data.forEach((bar, dataIndex) => {
                        const count = Number(dataset.data[dataIndex]) || 0;
                        const properties = bar.getProps(['x', 'y', 'base'], true);
                        if (!count || Math.abs(properties.x - properties.base) < 26) return;
                        context.save();
                        context.fillStyle = bandColors[dataset.label].text;
                        context.font = '600 12px Inter, sans-serif';
                        context.textAlign = 'center';
                        context.textBaseline = 'middle';
                        context.fillText(count.toLocaleString(), (properties.x + properties.base) / 2, properties.y);
                        context.restore();
                    });
                });

                operatorTotals.forEach((operatorTotal, dataIndex) => {
                    if (!operatorTotal) return;
                    const barPositions = chart.data.datasets.map((_, datasetIndex) => {
                        const bar = chart.getDatasetMeta(datasetIndex).data[dataIndex];
                        return bar ? bar.getProps(['x', 'y'], true) : null;
                    }).filter(Boolean);
                    if (!barPositions.length) return;
                    const endX = Math.max(...barPositions.map(position => position.x));
                    const y = barPositions[0].y;
                    context.save();
                    context.fillStyle = theme.textColor;
                    context.font = '600 12px Inter, sans-serif';
                    context.textAlign = 'left';
                    context.textBaseline = 'middle';
                    context.fillText(operatorTotal.toLocaleString(), endX + 8, y);
                    context.restore();
                });
            }
        };

        state.charts.sectorDistribution = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: operators,
                datasets: bands.map(band => ({
                    label: band,
                    data: countsByBand[band],
                    backgroundColor: bandColors[band].fill,
                    borderColor: bandColors[band].border,
                    borderWidth: 1,
                    borderSkipped: false
                }))
            },
            plugins: [segmentLabelsPlugin],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            color: theme.textColor,
                            usePointStyle: true,
                            padding: 16
                        }
                    },
                    tooltip: {
                        backgroundColor: theme.tooltipBg,
                        titleColor: theme.tooltipText,
                        bodyColor: theme.tooltipText,
                        callbacks: {
                            label(context) {
                                const count = Number(context.raw) || 0;
                                const rowTotal = operatorTotals[context.dataIndex];
                                const percentage = rowTotal ? (count / rowTotal * 100).toFixed(1) : '0.0';
                                return ` ${context.dataset.label}: ${count.toLocaleString()} sectors (${percentage}% of ${context.label})`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        beginAtZero: true,
                        suggestedMax: Math.max(...operatorTotals, 1) * 1.25,
                        ticks: { color: theme.textColor, precision: 0 },
                        grid: { color: theme.gridColor }
                    },
                    y: {
                        stacked: true,
                        ticks: { color: theme.textColor },
                        grid: { display: false }
                    }
                }
            }
        });
    }

    // Keep the low-throughput chart populated even when its optional filter is off.
    function getLowThroughputChartRecords() {
        const { operator, date, band, hour, site, lncel, busyHourOnly, degradedOnly } = state.filters;
        const searchVal = (elements.tableSearch ? elements.tableSearch.value.trim().toLowerCase() : '');
        const availableDates = state.rawRecords
            .map(row => row.Date)
            .filter(Boolean)
            .sort();
        const chartDates = state.viewMode === 'weekly' || state.viewMode === 'monthly'
            ? availableDates.filter(recordDate => matchesSelectedPeriod(recordDate, date, state.viewMode))
            : [date !== 'ALL' ? date : availableDates[availableDates.length - 1]];
        const chartDateSet = new Set(chartDates);

        return state.rawRecords.filter(row => {
            if (operator !== 'ALL' && row.Operator !== operator) return false;
            if (chartDateSet.size && !chartDateSet.has(row.Date)) return false;
            if (band !== 'ALL' && row.Band !== band) return false;
            if (hour !== 'ALL' && parseInt(row.Hour, 10) !== parseInt(hour, 10)) return false;

            const chartHour = parseInt(row.Hour, 10);
            if (state.viewMode !== 'hourly') {
                if (isNaN(chartHour) || chartHour !== 24) return false;
            } else if (isNaN(chartHour) || chartHour < 6 || chartHour > 23) {
                return false;
            }

            const rowSite = row.Site_Name || row.Site_ID;
            if (site !== 'ALL' && rowSite !== site) return false;
            if (lncel !== 'ALL' && String(row.LNCEL) !== String(lncel)) return false;

            if (busyHourOnly && parseInt(row.Hour, 10) !== 18) return false;

            if (degradedOnly) {
                const isDegraded = (row.VoLTE_Drop_Rate > 0.8) ||
                    (row.Drop_Rate > 0.8) ||
                    (row.DL_User_Throughput_Mbps > 0 && row.DL_User_Throughput_Mbps < 6.0) ||
                    (row.DL_PRB_Util > 75.0);
                if (!isDegraded) return false;
            }

            if (searchVal) {
                const searchable = `${row.Site_Name || ''} ${row.Site_ID || ''} ${row.LNCEL || ''} ${row.Operator || ''} ${row.Band || ''}`.toLowerCase();
                if (!searchable.includes(searchVal)) return false;
            }

            return true;
        });
    }

    // Chart 1: Hourly DL User Speed vs DL PRB Utilization
    function renderThroughputPrbTrendChart() {
        const ctx = document.getElementById('chartTrendThroughput');
        if (!ctx) return;

        const isDaily = state.viewMode !== 'hourly';
        const grouped = {};
        state.filteredRecords.forEach(r => {
            const key = isDaily ? r.Date : parseInt(r.Hour, 10);
            if (!key || (!isDaily && (isNaN(key) || key < 0 || key >= 24))) return;
            if (!grouped[key]) grouped[key] = { dlThrptSum: 0, prbSum: 0, count: 0 };
            grouped[key].dlThrptSum += Number(r.DL_User_Throughput_Mbps) || 0;
            grouped[key].prbSum += Number(r.DL_PRB_Util) || 0;
            grouped[key].count++;
        });
        let keys = Object.keys(grouped).sort((a, b) => isDaily ? a.localeCompare(b) : Number(a) - Number(b));
        const pointLimit = state.viewMode === 'monthly' ? 31 : 7;
        if (isDaily) keys = keys.slice(-pointLimit);
        const labels = keys.map(key => isDaily ? formatChartDate(key) : `${String(key).padStart(2, '0')}:00`);
        const groups = keys.map(key => grouped[key]);
        const thrptData = groups.map(group => (group.dlThrptSum / group.count / (isDaily ? 1000 : 1)).toFixed(2));
        const prbData = groups.map(group => (group.prbSum / group.count).toFixed(1));

        if (elements.trendChartTitle) {
            elements.trendChartTitle.textContent = state.viewMode === 'weekly' ? 'Weekly User Speed vs. Radio Congestion'
                : state.viewMode === 'monthly' ? 'Monthly User Speed vs. Radio Congestion'
                : isDaily ? 'Day-wise User Speed vs. Radio Congestion'
                : 'Hourly User Speed vs. Radio Congestion';
        }
        if (elements.trendChartSubtitle) {
            elements.trendChartSubtitle.textContent = isDaily
                ? `${state.viewMode === 'monthly' ? 'Daily trend for the selected month' : state.viewMode === 'weekly' ? 'Daily trend for the selected week' : 'Latest 7 days'} tracking DL User Speed (Mbps) against average PDSCH PRB Utilization (%)`
                : '24-hour time-series tracking DL User Speed (Mbps) against PDSCH PRB Utilization (%)';
        }

        if (state.charts.throughputPrb) state.charts.throughputPrb.destroy();

        const theme = getChartThemeColors();
        const opLabel = state.filters.operator === 'ALL' ? 'Combined' : state.filters.operator;

        state.charts.throughputPrb = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: `DL Speed - ${opLabel} (Mbps)`,
                        data: thrptData,
                        borderColor: '#00f0ff',
                        backgroundColor: 'rgba(0, 240, 255, 0.12)',
                        fill: true,
                        tension: 0.35,
                        yAxisID: 'y',
                        borderWidth: 2.5
                    },
                    {
                        label: `PRB Util - ${opLabel} (%)`,
                        data: prbData,
                        borderColor: '#f59e0b',
                        borderDash: [4, 4],
                        tension: 0.35,
                        yAxisID: 'y1',
                        borderWidth: 2
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { labels: { color: theme.textColor } }
                },
                scales: {
                    x: { ticks: { color: theme.textColor }, grid: { color: theme.gridColor } },
                    y: {
                        position: 'left',
                        title: { display: true, text: 'Throughput (Mbps)', color: theme.textColor },
                        ticks: { color: theme.textColor },
                        grid: { color: theme.gridColor }
                    },
                    y1: {
                        position: 'right',
                        max: 100,
                        title: { display: true, text: 'PRB Util (%)', color: theme.textColor },
                        ticks: { color: theme.textColor },
                        grid: { drawOnChartArea: false }
                    }
                }
            }
        });
    }

    // Chart 2: Worst-Performing Sectors (Clean Site and LNCEL, not concatenated)
    function renderWorstCellsChart() {
        const ctx = document.getElementById('chartWorstCells');
        if (!ctx) return;

        const availableDates = state.filteredRecords
            .map(r => r.Date)
            .filter(Boolean)
            .sort();
        const chartDate = state.filters.date !== 'ALL'
            ? state.filters.date
            : availableDates[availableDates.length - 1];
        const chartRecords = state.viewMode === 'weekly' || state.viewMode === 'monthly'
            ? state.filteredRecords
            : chartDate
            ? state.filteredRecords.filter(r => r.Date === chartDate)
            : state.filteredRecords;

        const sectorDrops = {};
        chartRecords.forEach(r => {
            const key = `${r.Site_Name || r.Site_ID}_@@_${r.LNCEL}`;
            if (!sectorDrops[key]) {
                sectorDrops[key] = {
                    sum: 0,
                    prbSum: 0,
                    count: 0,
                    site: r.Site_Name || r.Site_ID,
                    lncel: r.LNCEL,
                    op: r.Operator
                };
            }
            sectorDrops[key].sum += (r.VoLTE_Drop_Rate || r.Drop_Rate || 0);
            sectorDrops[key].prbSum += (r.DL_PRB_Util || 0);
            sectorDrops[key].count++;
        });

        const ranked = Object.values(sectorDrops)
            .map(s => ({
                site: s.site,
                lncel: s.lncel,
                op: s.op,
                avgDrop: s.sum / s.count
            }))
            .sort((a, b) => b.avgDrop - a.avgDrop)
            .slice(0, 8);

        const issueCellCount = Object.values(sectorDrops).filter(s => {
            const avgDrop = s.sum / s.count;
            const avgPrb = s.prbSum / s.count;
            return avgDrop >= 1.0 || avgPrb > 80;
        }).length;

        if (elements.worstCellsChartSubtitle) {
            elements.worstCellsChartSubtitle.textContent =
                `Top sectors ranked by Drop Call Rate (%) | Issue cells: ${issueCellCount}`;
        }

        if (state.charts.worstCells) state.charts.worstCells.destroy();

        const theme = getChartThemeColors();
        state.charts.worstCells = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ranked.map(r => {
                    const shortSite = r.site.length > 18 ? r.site.substring(0, 16) + '..' : r.site;
                    return `${shortSite} [LNCEL: ${r.lncel}]`;
                }),
                datasets: [{
                    label: 'Drop Call Rate (%)',
                    data: ranked.map(r => r.avgDrop.toFixed(2)),
                    backgroundColor: ranked.map(r => r.avgDrop >= 1.0 ? 'rgba(239, 68, 68, 0.85)' : 'rgba(245, 158, 11, 0.75)'),
                    borderColor: ranked.map(r => r.avgDrop >= 1.0 ? '#ef4444' : '#f59e0b'),
                    borderWidth: 1.5,
                    borderRadius: 6
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            title: (items) => {
                                const idx = items[0].dataIndex;
                                const item = ranked[idx];
                                return `Site: ${item.site}\nLNCEL: ${item.lncel} (${item.op})`;
                            },
                            label: (ctx) => `Drop Rate: ${ctx.raw}% (Issue threshold >= 1.0%)`
                        }
                    }
                },
                scales: {
                    x: { ticks: { color: theme.textColor }, grid: { color: theme.gridColor } },
                    y: { ticks: { color: theme.textColor, font: { size: 11 } }, grid: { display: false } }
                }
            }
        });
    }

    // Chart 3: Low-throughput sectors with low PRB utilization
    function renderLowThroughputCellsChart() {
        const ctx = document.getElementById('chartLowThroughputCells');
        if (!ctx) return;

        const sectors = {};
        getLowThroughputChartRecords().forEach(r => {
            const throughput = getDlThroughputMbps(r);
            const utilization = parseFloat(r.DL_PRB_Util);
            if (!Number.isFinite(throughput) || !Number.isFinite(utilization) ||
                throughput <= 0 || throughput >= 2 || utilization >= 60) return;

            const key = String(r.Site_Name || r.Site_ID || 'Unknown Site');
            if (!sectors[key]) {
                sectors[key] = {
                    throughputMin: throughput,
                    utilizationSum: 0,
                    count: 0,
                    site: r.Site_Name || r.Site_ID || 'Unknown Site',
                    sectors: [],
                    op: r.Operator
                };
            }
            if (throughput < sectors[key].throughputMin) {
                sectors[key].throughputMin = throughput;
                sectors[key].sectors = [r.LNCEL || r.Cell_ID || 'Unknown Sector'];
            } else if (throughput === sectors[key].throughputMin) {
                const sectorId = r.LNCEL || r.Cell_ID || 'Unknown Sector';
                if (!sectors[key].sectors.includes(sectorId)) {
                    sectors[key].sectors.push(sectorId);
                }
            }
            sectors[key].utilizationSum += utilization;
            sectors[key].count++;
        });

        const ranked = Object.values(sectors)
            .map(s => ({
                site: s.site,
                sectors: s.sectors,
                op: s.op,
                avgThroughput: s.throughputMin,
                avgUtilization: s.utilizationSum / s.count
            }))
            .sort((a, b) => a.avgThroughput - b.avgThroughput)
            .slice(0, 8);

        if (elements.lowThroughputChartSubtitle) {
            const issueCellCount = Object.keys(sectors).length;
            elements.lowThroughputChartSubtitle.textContent =
                `Top sectors with DL throughput < 2 Mbps and PRB utilization < 60% | Issue cells: ${issueCellCount}`;
        }

        if (state.charts.lowThroughputCells) state.charts.lowThroughputCells.destroy();

        const theme = getChartThemeColors();
        state.charts.lowThroughputCells = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ranked.map(r => {
                    const shortSite = String(r.site).length > 20 ? String(r.site).substring(0, 18) + '..' : r.site;
                    return shortSite;
                }),
                datasets: [{
                    label: 'DL Throughput (Mbps)',
                    data: ranked.map(r => r.avgThroughput),
                    backgroundColor: 'rgba(239, 68, 68, 0.8)',
                    borderColor: '#ef4444',
                    borderWidth: 1.5,
                    borderRadius: 6,
                    minBarLength: 12
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            title: (items) => {
                                const item = ranked[items[0].dataIndex];
                                return [
                                    item.site,
                                    `Sector ID: ${item.sectors.join(', ')} (${item.op})`
                                ];
                            },
                            label: (ctx) => {
                                const item = ranked[ctx.dataIndex];
                                return `DL Throughput: ${Number(ctx.raw).toFixed(3)} Mbps | PRB Util: ${item.avgUtilization.toFixed(1)}%`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true,
                        max: 2,
                        ticks: { color: theme.textColor },
                        grid: { color: theme.gridColor },
                        title: { display: true, text: 'DL Throughput (Mbps)', color: theme.textColor }
                    },
                    y: { ticks: { color: theme.textColor, font: { size: 11 } }, grid: { display: false } }
                }
            }
        });
    }

    // Chart 4: 24-Hour Diurnal Payload (GB) vs Active Connected Users
    function renderHourlyTrafficProfileChart() {
        const ctx = document.getElementById('chartHourlyProfile');
        if (!ctx) return;

        const isDaily = state.viewMode !== 'hourly';
        const grouped = {};
        state.filteredRecords.forEach(r => {
            const key = isDaily ? r.Date : parseInt(r.Hour, 10);
            if (!key || (!isDaily && (isNaN(key) || key < 0 || key >= 24))) return;
            if (!grouped[key]) grouped[key] = { trafficSum: 0, userSum: 0, count: 0 };
            grouped[key].trafficSum += Number(r.Total_Traffic_GB) || 0;
            grouped[key].userSum += Number(r.Active_Users_Avg) || 0;
            grouped[key].count++;
        });

        let keys = Object.keys(grouped).sort((a, b) => isDaily ? a.localeCompare(b) : Number(a) - Number(b));
        if (isDaily) keys = keys.slice(-(state.viewMode === 'monthly' ? 31 : 7));
        const labels = keys.map(key => isDaily ? formatChartDate(key) : `${String(key).padStart(2, '0')}:00`);
        const groups = keys.map(key => grouped[key]);
        const trafficData = groups.map(group => group.trafficSum.toFixed(1));
        const userData = groups.map(group => group.count > 0 ? Math.round(group.userSum / group.count) : 0);

        if (elements.hourlyProfileChartTitle) {
            elements.hourlyProfileChartTitle.textContent = state.viewMode === 'monthly'
                ? 'Monthly Payload & Connected Users'
                : state.viewMode === 'weekly'
                    ? 'Weekly Payload & Connected Users'
                    : isDaily ? '7-Day Payload & Connected Users'
                        : '24-Hour Payload & Connected Users';
        }
        if (elements.hourlyProfileChartSubtitle) {
            elements.hourlyProfileChartSubtitle.textContent = isDaily
                ? `${state.viewMode === 'monthly' ? 'Daily totals for the selected month' : state.viewMode === 'weekly' ? 'Daily totals for the selected week' : 'Latest 7 days'} data payload (GB) and average connected UEs across all sites`
                : 'Hourly total data payload (GB) and average connected UEs across all sites';
        }

        if (state.charts.hourlyProfile) state.charts.hourlyProfile.destroy();

        const theme = getChartThemeColors();
        const opLabel = state.filters.operator === 'ALL' ? 'All Operators' : state.filters.operator;

        state.charts.hourlyProfile = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        type: 'bar',
                        label: `Payload - ${opLabel} (GB)`,
                        data: trafficData,
                        backgroundColor: 'rgba(0, 240, 255, 0.65)',
                        borderColor: '#00f0ff',
                        borderWidth: 1,
                        borderRadius: 4,
                        yAxisID: 'y'
                    },
                    {
                        type: 'line',
                        label: `Connected UEs - ${opLabel}`,
                        data: userData,
                        borderColor: '#a855f7',
                        borderWidth: 2.5,
                        tension: 0.3,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { labels: { color: theme.textColor } }
                },
                scales: {
                    x: { ticks: { color: theme.textColor }, grid: { color: theme.gridColor } },
                    y: {
                        position: 'left',
                        title: { display: true, text: 'Payload (GB)', color: theme.textColor },
                        ticks: { color: theme.textColor },
                        grid: { color: theme.gridColor }
                    },
                    y1: {
                        position: 'right',
                        title: { display: true, text: 'Active UEs', color: theme.textColor },
                        ticks: { color: theme.textColor },
                        grid: { drawOnChartArea: false }
                    }
                }
            }
        });
    }

    // Chart 4: VoLTE Traffic vs VoLTE Drop Rate
    function renderVolteQualityTrendChart() {
        const ctx = document.getElementById('chartAccessibilityTrend');
        if (!ctx) return;

        const isDaily = state.viewMode !== 'hourly';
        const grouped = {};
        state.filteredRecords.forEach(r => {
            const key = isDaily ? r.Date : parseInt(r.Hour, 10);
            if (!key || (!isDaily && (isNaN(key) || key < 0 || key >= 24))) return;
            if (!grouped[key]) grouped[key] = { volteErlSum: 0, volteDropSum: 0, count: 0 };
            grouped[key].volteErlSum += Number(r.VoLTE_Traffic_Erl) || 0;
            grouped[key].volteDropSum += Number(r.VoLTE_Drop_Rate) || 0;
            grouped[key].count++;
        });

        let keys = Object.keys(grouped).sort((a, b) => isDaily ? a.localeCompare(b) : Number(a) - Number(b));
        if (isDaily) keys = keys.slice(-(state.viewMode === 'monthly' ? 31 : 7));
        const labels = keys.map(key => isDaily ? formatChartDate(key) : `${String(key).padStart(2, '0')}:00`);
        const groups = keys.map(key => grouped[key]);
        const erlData = groups.map(group => group.volteErlSum.toFixed(1));
        const dropData = groups.map(group => group.count > 0 ? (group.volteDropSum / group.count).toFixed(2) : 0);

        if (elements.volteQualityChartTitle) {
            elements.volteQualityChartTitle.textContent = state.viewMode === 'monthly'
                ? 'Monthly VoLTE Traffic (Erl) vs. Drop Rate (%)'
                : state.viewMode === 'weekly'
                    ? 'Weekly VoLTE Traffic (Erl) vs. Drop Rate (%)'
                    : isDaily ? '7-Day VoLTE Traffic (Erl) vs. Drop Rate (%)'
                        : 'VoLTE Traffic (Erl) vs. Drop Rate (%)';
        }
        if (elements.volteQualityChartSubtitle) {
            elements.volteQualityChartSubtitle.textContent = isDaily
                ? `${state.viewMode === 'monthly' ? 'Daily trend for the selected month' : state.viewMode === 'weekly' ? 'Daily trend for the selected week' : 'Latest 7 days'} voice load profile and call drop correlation`
                : '24-hour voice load profile and call drop correlation';
        }

        if (state.charts.volteQuality) state.charts.volteQuality.destroy();

        const theme = getChartThemeColors();
        const opLabel = state.filters.operator === 'ALL' ? 'All Operators' : state.filters.operator;

        state.charts.volteQuality = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: `VoLTE Traffic - ${opLabel} (Erl)`,
                        data: erlData,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.1)',
                        fill: true,
                        tension: 0.3,
                        yAxisID: 'y'
                    },
                    {
                        label: `VoLTE Drop Rate - ${opLabel} (%)`,
                        data: dropData,
                        borderColor: '#ef4444',
                        borderDash: [5, 5],
                        tension: 0.3,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { labels: { color: theme.textColor } }
                },
                scales: {
                    x: { ticks: { color: theme.textColor }, grid: { color: theme.gridColor } },
                    y: {
                        position: 'left',
                        title: { display: true, text: 'Voice Traffic (Erlang)', color: theme.textColor },
                        ticks: { color: theme.textColor },
                        grid: { color: theme.gridColor }
                    },
                    y1: {
                        position: 'right',
                        title: { display: true, text: 'VoLTE Drop (%)', color: theme.textColor },
                        ticks: { color: theme.textColor },
                        grid: { drawOnChartArea: false }
                    }
                }
            }
        });
    }

    // Cell Data Grid: Clean Site Name & LNCEL as distinct separate columns
    function renderTable() {
        if (!elements.cellTableBody) return;

        const sectorMap = {};
        state.filteredRecords.forEach(r => {
            const key = `${r.Site_Name || r.Site_ID}_@@_${r.LNCEL}`;
            if (!sectorMap[key]) {
                sectorMap[key] = {
                    site: r.Site_Name || r.Site_ID || 'Unknown',
                    lncel: r.LNCEL || '-',
                    operator: r.Operator || 'Unknown',
                    band: r.Band || 'FDD',
                    date: r.Date || '',
                    volteCssrSum: 0,
                    rrcSum: 0,
                    volteDropSum: 0,
                    dropSum: 0,
                    dlThrptSum: 0,
                    prbSum: 0,
                    trafficSum: 0,
                    volteErlSum: 0,
                    cqiSum: 0,
                    sinrSum: 0,
                    count: 0
                };
            }
            const cm = sectorMap[key];
            cm.volteCssrSum += (r.VoLTE_CSSR || 0);
            cm.rrcSum += (r.RRC_Setup_SR || 0);
            cm.volteDropSum += (r.VoLTE_Drop_Rate || 0);
            cm.dropSum += (r.Drop_Rate || 0);
            cm.dlThrptSum += (r.DL_User_Throughput_Mbps || 0);
            cm.prbSum += (r.DL_PRB_Util || 0);
            cm.trafficSum += (r.Total_Traffic_GB || 0);
            cm.volteErlSum += (r.VoLTE_Traffic_Erl || 0);
            cm.cqiSum += (r.CQI || 0);
            cm.sinrSum += (r.SINR || 0);
            cm.count++;
        });

        const rows = Object.values(sectorMap).map(c => ({
            site: c.site,
            lncel: c.lncel,
            operator: c.operator,
            band: c.band,
            date: c.date,
            volteCssr: (c.volteCssrSum / c.count).toFixed(2),
            rrcSr: (c.rrcSum / c.count).toFixed(2),
            volteDrop: (c.volteDropSum / c.count).toFixed(2),
            dropRate: (c.dropSum / c.count).toFixed(2),
            dlThrpt: (c.dlThrptSum / c.count / 1000).toFixed(3),
            dlPrb: (c.prbSum / c.count).toFixed(1),
            totalTraffic: c.trafficSum.toFixed(2),
            volteErl: c.volteErlSum.toFixed(2),
            cqi: (c.cqiSum / c.count).toFixed(1),
            sinr: (c.sinrSum / c.count).toFixed(1)
        }));

        // Sorting
        const { sortColumn, sortDirection } = state.pagination;
        rows.sort((a, b) => {
            let valA = a[sortColumn] !== undefined ? a[sortColumn] : a.site;
            let valB = b[sortColumn] !== undefined ? b[sortColumn] : b.site;
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);
            if (!isNaN(numA) && !isNaN(numB)) {
                return sortDirection === 'asc' ? numA - numB : numB - numA;
            }
            return sortDirection === 'asc' ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
        });

        // Pagination
        const { currentPage, pageSize } = state.pagination;
        const totalRows = rows.length;
        const totalPages = Math.ceil(totalRows / pageSize) || 1;
        const startIndex = (currentPage - 1) * pageSize;
        const pagedRows = rows.slice(startIndex, startIndex + pageSize);

        if (elements.paginationInfo) {
            elements.paginationInfo.innerText = `Showing ${startIndex + 1}-${Math.min(startIndex + pageSize, totalRows)} of ${totalRows} sectors`;
        }
        if (elements.prevPageBtn) elements.prevPageBtn.disabled = currentPage <= 1;
        if (elements.nextPageBtn) elements.nextPageBtn.disabled = currentPage >= totalPages;

        elements.cellTableBody.innerHTML = '';
        if (pagedRows.length === 0) {
            elements.cellTableBody.innerHTML = `<tr><td colspan="13" style="text-align:center; padding: 2rem; color: var(--text-muted);">No sector data found matching current filters.</td></tr>`;
            return;
        }

        pagedRows.forEach(row => {
            const tr = document.createElement('tr');
            const opBadgeClass = row.operator === 'AIRTEL' ? 'badge-airtel' : (row.operator === 'VIL' ? 'badge-vil' : '');
            const volteDropClass = parseFloat(row.volteDrop) <= 1.0 ? 'good' : (parseFloat(row.volteDrop) <= 1.5 ? 'warning' : 'critical');
            const rrcClass = parseFloat(row.rrcSr) >= 99.0 ? 'good' : (parseFloat(row.rrcSr) >= 98.0 ? 'warning' : 'critical');
            const prbClass = parseFloat(row.dlPrb) < 70.0 ? 'good' : (parseFloat(row.dlPrb) < 85.0 ? 'warning' : 'critical');
            const dlThroughput = parseFloat(row.dlThrpt);
            const dlThroughputClass = dlThroughput >= 3.0 ? 'good' : (dlThroughput >= 2.0 ? 'warning' : 'critical');

            tr.innerHTML = `
                <td><strong>${row.site}</strong></td>
                <td><span class="status-badge" style="background:var(--bg-surface-elevated); font-weight:700;">${row.lncel}</span></td>
                <td><span class="status-badge ${opBadgeClass}">${row.operator}</span></td>
                <td><span class="status-badge" style="background:var(--bg-surface-elevated);">${row.band}</span></td>
                <td><span class="status-badge ${rrcClass}">${row.rrcSr}%</span></td>
                <td>${row.volteCssr}%</td>
                <td><span class="status-badge ${volteDropClass}">${row.volteDrop}%</span></td>
                <td>${row.dropRate}%</td>
                <td><span class="status-badge ${dlThroughputClass}">${row.dlThrpt} Mbps</span></td>
                <td><span class="status-badge ${prbClass}">${row.dlPrb}%</span></td>
                <td><strong>${row.totalTraffic} GB</strong></td>
                <td>${row.volteErl} Erl</td>
                <td>${row.cqi}</td>
            `;
            elements.cellTableBody.appendChild(tr);
        });
    }

    // Chart 5: Last 8 Days — Day-wise Total Data Volume (GB) per Operator
    function renderDailyVolumeChart() {
        const ctx = document.getElementById('chartDailyVolume');
        if (!ctx) return;

        // Aggregate total traffic GB by date and operator from filteredRecords
        const dateOpMap = {};
        const operatorSet = new Set();

        state.filteredRecords.forEach(r => {
            const d = r.Date;
            if (!d) return;
            const op = r.Operator || 'Combined';
            operatorSet.add(op);
            if (!dateOpMap[d]) dateOpMap[d] = {};
            if (!dateOpMap[d][op]) dateOpMap[d][op] = 0;
            dateOpMap[d][op] += (parseFloat(r.Total_Traffic_GB) || 0);
        });

        // Show the full selected week or month; retain the recent 8-day window otherwise.
        const allDates = Object.keys(dateOpMap).sort();
        const pointLimit = state.viewMode === 'monthly' ? 31
            : state.viewMode === 'weekly' ? 7
            : 8;
        const last8Dates = allDates.slice(-pointLimit);

        if (last8Dates.length === 0) {
            if (state.charts.dailyVolume) { state.charts.dailyVolume.destroy(); delete state.charts.dailyVolume; }
            return;
        }

        const volumeTitle = document.getElementById('dailyVolumeChartTitle');
        const volumeSubtitle = document.getElementById('dailyVolumeChartSubtitle');
        if (volumeTitle) {
            volumeTitle.textContent = state.viewMode === 'monthly'
                ? 'Monthly Data Volume by Operator'
                : state.viewMode === 'weekly'
                    ? 'Weekly Data Volume by Operator'
                    : '📅 Last 8 Days — Total Data Volume (GB) per Operator';
        }
        if (volumeSubtitle) {
            volumeSubtitle.textContent = state.viewMode === 'monthly'
                ? 'Daily payload for the selected month — AIRTEL & VIL'
                : state.viewMode === 'weekly'
                    ? 'Daily payload for the selected week — AIRTEL & VIL'
                    : 'Day-wise aggregated data payload (GB) — AIRTEL & VIL compared across the last 8 available dates';
        }

        const theme = getChartThemeColors();
        const operators = [...operatorSet].sort();
        const opColors = {
            'AIRTEL':  { bar: 'rgba(251, 146, 60, 0.85)',  border: '#f97316' },
            'VIL':     { bar: 'rgba(168, 85, 247, 0.85)',  border: '#a855f7' },
            'Combined':{ bar: 'rgba(0, 240, 255, 0.75)',   border: '#00f0ff' },
            'Sample_RAN': { bar: 'rgba(16, 185, 129, 0.75)', border: '#10b981' }
        };
        const fallbackColors = ['rgba(59,130,246,0.8)', 'rgba(245,158,11,0.8)', 'rgba(239,68,68,0.8)'];

        // Build datasets — one bar series per operator
        const datasets = operators.map((op, idx) => {
            const color = opColors[op] || { bar: fallbackColors[idx % fallbackColors.length], border: fallbackColors[idx % fallbackColors.length] };
            return {
                label: `${op} (GB)`,
                data: last8Dates.map(d => {
                    const v = dateOpMap[d][op] || 0;
                    return parseFloat(v.toFixed(2));
                }),
                backgroundColor: color.bar,
                borderColor: color.border,
                borderWidth: 2,
                barPercentage: 0.7,
                categoryPercentage: 0.7,
                maxBarThickness: 40,
                borderRadius: 6,
                borderSkipped: false
            };
        });

        // Short date labels: "21-Sep"
        const shortDates = last8Dates.map(d => {
            try {
                const parts = String(d).split(/[-\/]/);
                if (parts.length === 3) {
                    // Try both DD-MM-YYYY and YYYY-MM-DD
                    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
                    let day, mon;
                    if (parts[0].length === 4) { // YYYY-MM-DD
                        day = parseInt(parts[2], 10);
                        mon = months[parseInt(parts[1], 10) - 1] || parts[1];
                    } else { // DD-MM-YYYY
                        day = parseInt(parts[0], 10);
                        mon = months[parseInt(parts[1], 10) - 1] || parts[1];
                    }
                    return `${day}-${mon}`;
                }
            } catch(e) {}
            return d;
        });

        if (state.charts.dailyVolume) state.charts.dailyVolume.destroy();

        state.charts.dailyVolume = new Chart(ctx, {
            type: 'bar',
            data: { labels: shortDates, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 600, easing: 'easeOutQuart' },
                plugins: {
                    legend: {
                        display: true,
                        labels: { color: theme.textColor, font: { size: 11 }, padding: 16 }
                    },
                    tooltip: {
                        backgroundColor: theme.tooltipBg,
                        titleColor: theme.tooltipText,
                        bodyColor: theme.tooltipText,
                        borderColor: 'rgba(0,240,255,0.3)',
                        borderWidth: 1,
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString()} GB`,
                            afterBody: (items) => {
                                // Show total across all operators for this date
                                const dateIdx = items[0].dataIndex;
                                const dateKey = last8Dates[dateIdx];
                                const total = operators.reduce((s, op) => s + (dateOpMap[dateKey][op] || 0), 0);
                                return [`Total All: ${total.toFixed(2)} GB`];
                            }
                        }
                    },
                    // Annotate the Busy Hour column with a note on the chart title (no plugin needed)
                },
                scales: {
                    x: {
                        stacked: false,
                        ticks: { color: theme.textColor, font: { size: 11 } },
                        grid: { color: theme.gridColor }
                    },
                    y: {
                        beginAtZero: true,
                        ticks: {
                            color: theme.textColor,
                            font: { size: 11 },
                            callback: v => v >= 1000 ? `${(v/1000).toFixed(1)}TB` : `${v}GB`
                        },
                        grid: { color: theme.gridColor },
                        title: { display: true, text: 'Total Data Volume (GB)', color: theme.textColor, font: { size: 11 } }
                    }
                }
            }
        });
    }

    // Export Table Data to CSV
    function exportFilteredDataToCsv() {
        if (state.filteredRecords.length === 0) {
            alert('No data to export.');
            return;
        }

        const headers = ['Date', 'Hour', 'Operator', 'Band', 'Site_Name', 'Site_ID', 'LNCEL', 'Cell_Availability', 'VoLTE_CSSR', 'RRC_Setup_SR', 'VoLTE_Drop_Rate', 'Drop_Rate', 'DL_User_Throughput_Mbps', 'Total_Traffic_GB', 'VoLTE_Traffic_Erl', 'DL_PRB_Util', 'CQI', 'SINR'];
        let csvContent = 'data:text/csv;charset=utf-8,' + headers.join(',') + '\n';

        state.filteredRecords.forEach(r => {
            const rowValues = headers.map(h => {
                let val = r[h] !== undefined ? r[h] : '';
                if (typeof val === 'string' && val.includes(',')) {
                    val = `"${val}"`;
                }
                return val;
            });
            csvContent += rowValues.join(',') + '\n';
        });

        const opPrefix = state.filters.operator !== 'ALL' ? `${state.filters.operator}_` : '';
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement('a');
        link.setAttribute('href', encodedUri);
        link.setAttribute('download', `${opPrefix}4G_VoLTE_KPI_Export_${new Date().toISOString().split('T')[0]}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // Universal Excel & CSV File Uploader Engine
    function handleFileUpload(file) {
        if (!file) return;
        const reader = new FileReader();
        const isCsv = file.name.endsWith('.csv');

        reader.onload = function (e) {
            try {
                let parsedRows = [];
                if (isCsv) {
                    parsedRows = parseCsvString(e.target.result);
                } else if (typeof XLSX !== 'undefined') {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    parsedRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
                } else {
                    alert('Excel parser library (SheetJS) is initializing. Please try again.');
                    return;
                }

                if (parsedRows.length === 0) {
                    alert('The uploaded file appears to be empty.');
                    return;
                }

                processUploadedData(parsedRows, file.name);
            } catch (err) {
                console.error('Error parsing file:', err);
                alert('Failed to parse the file: ' + err.message);
            }
        };

        if (isCsv) {
            reader.readAsText(file);
        } else {
            reader.readAsArrayBuffer(file);
        }
    }

    function parseCsvString(csvText) {
        const lines = csvText.split(/\r\n|\n/).filter(line => line.trim().length > 0);
        if (lines.length < 2) return [];

        const headers = parseCsvLine(lines[0]);
        const rows = [];

        for (let i = 1; i < lines.length; i++) {
            const values = parseCsvLine(lines[i]);
            const rowObj = {};
            headers.forEach((h, idx) => {
                rowObj[h.trim()] = values[idx] !== undefined ? values[idx].trim() : '';
            });
            rows.push(rowObj);
        }
        return rows;
    }

    function parseCsvLine(text) {
        const result = [];
        let cur = '';
        let inQuotes = false;
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(cur);
                cur = '';
            } else {
                cur += char;
            }
        }
        result.push(cur);
        return result;
    }

    function detectColumnMapping(sampleRow) {
        const fileColumns = Object.keys(sampleRow);
        const mapping = {};

        Object.keys(KPI_DEFINITIONS).forEach(stdKey => {
            const def = KPI_DEFINITIONS[stdKey];
            const cleanStdKey = stdKey.toLowerCase().replace(/[^a-z0-9]/g, '');

            let matchedCol = fileColumns.find(c => {
                const cleanC = c.toLowerCase().replace(/[^a-z0-9]/g, '');
                return cleanC === cleanStdKey;
            });

            if (!matchedCol && def.aliases) {
                matchedCol = fileColumns.find(c => {
                    const cleanC = c.toLowerCase().replace(/[^a-z0-9]/g, '');
                    return def.aliases.some(alias => {
                        const cleanAlias = alias.toLowerCase().replace(/[^a-z0-9]/g, '');
                        return cleanC === cleanAlias || cleanC.includes(cleanAlias);
                    });
                });
            }

            if (matchedCol) {
                mapping[stdKey] = matchedCol;
            }
        });

        return mapping;
    }

    function processUploadedData(rawRows, fileName) {
        const detected = detectColumnMapping(rawRows[0]);
        state.detectedColumns = detected;
        state.lastUploadedRawRows = rawRows;
        state.lastUploadedFileName = fileName;

        openMappingModal(rawRows[0], detected);
    }

    function applyMappingAndLoad() {
        const rawRows = state.lastUploadedRawRows;
        const mapping = state.detectedColumns;

        const normalized = rawRows.map((r, idx) => {
            const norm = {};
            Object.keys(KPI_DEFINITIONS).forEach(stdKey => {
                const sourceCol = mapping[stdKey];
                let val = sourceCol && r[sourceCol] !== undefined ? r[sourceCol] : null;

                if (['Cell_Availability', 'VoLTE_CSSR', 'E2E_CSSR', 'RRC_Setup_SR', 'ERAB_Setup_SR', 'Total_Traffic_GB', 'VoLTE_Traffic_Erl', 'VoLTE_Drop_Rate', 'Drop_Rate', 'DL_User_Throughput_Mbps', 'UL_User_Throughput_Mbps', 'DL_PRB_Util', 'UL_PRB_Util', 'CQI', 'SINR', 'RSSI', 'Handover_SR'].includes(stdKey)) {
                    val = parseFloat(val);
                    if (isNaN(val)) val = 0;
                } else if (['Active_Users_Avg', 'Max_Users'].includes(stdKey)) {
                    val = parseInt(val, 10);
                    if (isNaN(val)) val = 0;
                } else if (val === null || val === undefined) {
                    val = '';
                }
                norm[stdKey] = val;
            });

            let timeVal = norm.Timestamp || norm.Date || '';
            if (typeof timeVal === 'number' && timeVal > 30000) {
                const frac = timeVal - Math.floor(timeVal);
                norm.Hour = Math.floor(frac * 24 + 0.0001);
            } else if (typeof timeVal === 'string' && timeVal.includes(':')) {
                const match = timeVal.match(/(\d{1,2}):\d{2}/);
                if (match) norm.Hour = parseInt(match[1], 10);
            }
            if (norm.Hour === undefined) norm.Hour = 0;

            // Keep Site Name and LNCEL completely separate
            norm.Site_Name = norm.Site_Name || norm.Site_ID || `Site_${idx + 1}`;
            norm.LNCEL = norm.LNCEL || '';
            if (!norm.Operator) norm.Operator = 'Unknown';
            norm.Band = resolveRecordBand(norm);

            return norm;
        });

        closeMappingModal();
        loadProcessedRecords(normalized, state.lastUploadedFileName, false);
    }

    function openMappingModal(sampleRow, detectedMapping) {
        if (!elements.mappingModal || !elements.mappingGrid) return;
        const fileColumns = Object.keys(sampleRow);

        elements.mappingGrid.innerHTML = '';
        Object.keys(KPI_DEFINITIONS).forEach(stdKey => {
            const def = KPI_DEFINITIONS[stdKey];
            const currentSelected = detectedMapping[stdKey] || '';

            const rowDiv = document.createElement('div');
            rowDiv.className = 'mapping-row';

            let optionsHtml = '<option value="">-- Not Mapped --</option>';
            fileColumns.forEach(fc => {
                const isSel = fc === currentSelected ? 'selected' : '';
                optionsHtml += `<option value="${fc}" ${isSel}>${fc}</option>`;
            });

            rowDiv.innerHTML = `
                <div class="mapping-target">
                    ${def.label}
                    <small>Mapped to: <strong>${currentSelected || 'None'}</strong></small>
                </div>
                <select class="form-select" data-stdkey="${stdKey}">
                    ${optionsHtml}
                </select>
            `;
            elements.mappingGrid.appendChild(rowDiv);
        });

        elements.mappingModal.classList.add('active');
    }

    function closeMappingModal() {
        if (elements.mappingModal) {
            elements.mappingModal.classList.remove('active');
        }

    }

    // Event Listeners
    function setupEventListeners() {
        if (elements.themeToggleBtn) {
            elements.themeToggleBtn.addEventListener('click', toggleTheme);
        }
        if (elements.viewModeContainer) {
            elements.viewModeContainer.addEventListener('click', (event) => {
                const button = event.target.closest('.view-mode-btn');
                if (button) setViewMode(button.dataset.viewMode);
            });
        }
        document.querySelectorAll('[data-dashboard-nav]').forEach(link => {
            link.addEventListener('click', event => {
                const target = document.querySelector(link.getAttribute('href'));
                if (!target) return;
                event.preventDefault();

                const isInventoryNavigation = target.id === 'networkInventory' || Boolean(link.closest('.sidebar-subnav'));
                document.body.classList.toggle('inventory-only', isInventoryNavigation);
                document.body.classList.toggle('inventory-show-assets', target.id === 'inventorySiteAssets');
                document.body.classList.toggle('inventory-show-ip', target.id === 'inventoryIpAddresses');
                if (link.hasAttribute('data-requires-daily') && state.viewMode !== 'daily') {
                    setViewMode('daily');
                }
                document.querySelectorAll('[data-dashboard-nav]').forEach(navLink => {
                    const isInventoryParent = link.hasAttribute('data-inventory-nav') &&
                        navLink.getAttribute('href') === '#networkInventory';
                    const isActive = navLink === link || isInventoryParent;
                    navLink.classList.toggle('active', isActive);
                    if (navLink === link) navLink.setAttribute('aria-current', link.hasAttribute('data-inventory-nav') ? 'location' : 'page');
                    else if (isInventoryParent) navLink.setAttribute('aria-current', 'page');
                    else navLink.removeAttribute('aria-current');
                });
                history.replaceState(null, '', link.getAttribute('href'));
                requestAnimationFrame(() => target.scrollIntoView({ behavior: 'smooth', block: 'start' }));
            });
        });

        // Dropzone
        if (elements.dropzone && elements.fileInput) {
            elements.dropzone.addEventListener('click', () => elements.fileInput.click());
            elements.fileInput.addEventListener('change', (e) => {
                if (e.target.files && e.target.files[0]) {
                    handleFileUpload(e.target.files[0]);
                }
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                elements.dropzone.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    elements.dropzone.classList.add('dragover');
                });
            });

            ['dragleave', 'drop'].forEach(eventName => {
                elements.dropzone.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    elements.dropzone.classList.remove('dragover');
                });
            });

            elements.dropzone.addEventListener('drop', (e) => {
                if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                    handleFileUpload(e.dataTransfer.files[0]);
                }
            });
        }

        // Action Buttons
        if (elements.loadUserReportBtn) {
            elements.loadUserReportBtn.addEventListener('click', loadDefaultReportData);
        }
        if (elements.loadSampleBtn) {
            elements.loadSampleBtn.addEventListener('click', () => {
                if (window.SampleLTEData) {
                    const data = window.SampleLTEData.generateDataset(14).map(r => ({
                        ...r,
                        LNCEL: r.Cell_ID,
                        Operator: 'Sample_RAN'
                    }));
                    loadProcessedRecords(data, 'Preloaded 14-Day Simulation Sample', false);
                }
            });
        }

        // Modal Controls
        if (elements.closeModalBtn) elements.closeModalBtn.addEventListener('click', closeMappingModal);
        if (elements.cancelModalBtn) elements.cancelModalBtn.addEventListener('click', closeMappingModal);
        if (elements.saveMappingBtn) {
            elements.saveMappingBtn.addEventListener('click', () => {
                const selects = elements.mappingGrid.querySelectorAll('select[data-stdkey]');
                selects.forEach(sel => {
                    const stdKey = sel.getAttribute('data-stdkey');
                    state.detectedColumns[stdKey] = sel.value;
                });
                applyMappingAndLoad();
            });
        }

        // Operator Filter Dropdown
        if (elements.filterOperator) {
            elements.filterOperator.addEventListener('change', (e) => {
                setOperatorFilter(e.target.value);
            });
        }

        // Date Filter Dropdown
        if (elements.filterDate) {
            elements.filterDate.addEventListener('change', (e) => {
                state.filters.date = e.target.value;
                state.filters.alertDateFilterTouched = true;
                applyFilters();
            });
        }

        // Band Filter Dropdown
        if (elements.filterBand) {
            elements.filterBand.addEventListener('change', (e) => {
                state.filters.band = e.target.value;
                applyFilters();
            });
        }

        // Hour Filter Dropdown
        if (elements.filterHour) {
            elements.filterHour.addEventListener('change', (e) => {
                state.filters.hour = e.target.value;
                applyFilters();
            });
        }

        // Site Filter
        if (elements.filterSite) {
            elements.filterSite.addEventListener('change', (e) => {
                state.filters.site = e.target.value;
                applyFilters();
            });
        }

        // LNCEL Filter (clean, separate LNCEL selector)
        if (elements.filterLncel) {
            elements.filterLncel.addEventListener('change', (e) => {
                state.filters.lncel = e.target.value;
                applyFilters();
            });
        }

        // Toggles
        if (elements.filterBusyHour) {
            elements.filterBusyHour.addEventListener('change', (e) => {
                state.filters.busyHourOnly = e.target.checked;
                applyFilters();
            });
        }
        if (elements.filterDegraded) {
            elements.filterDegraded.addEventListener('change', (e) => {
                state.filters.degradedOnly = e.target.checked;
                applyFilters();
            });
        }
        if (elements.filterLowThroughputLowUtilization) {
            elements.filterLowThroughputLowUtilization.addEventListener('change', (e) => {
                state.filters.lowThroughputLowUtilization = e.target.checked;
                applyFilters();
            });
        }

        // Reset Filters Button
        if (elements.resetFiltersBtn) {
            elements.resetFiltersBtn.addEventListener('click', () => {
                state.filters.operator = 'ALL';
                const availableDates = [...new Set(state.rawRecords.map(r => r.Date).filter(Boolean))].sort();
                state.filters.date = state.viewMode === 'daily' || !availableDates.length
                    ? 'ALL'
                    : getPeriodKey(availableDates[availableDates.length - 1], state.viewMode);
                state.filters.band = 'ALL';
                state.filters.hour = 'ALL';
                state.filters.site = 'ALL';
                state.filters.lncel = 'ALL';
                state.filters.busyHourOnly = false;
                state.filters.degradedOnly = false;
                state.filters.alertDateFilterTouched = false;
                state.filters.lowThroughputLowUtilization = false;
                if (elements.filterOperator) elements.filterOperator.value = 'ALL';
                if (elements.filterDate) elements.filterDate.value = state.filters.date;
                if (elements.filterBand) elements.filterBand.value = 'ALL';
                if (elements.filterHour) elements.filterHour.value = 'ALL';
                if (elements.filterSite) elements.filterSite.value = 'ALL';
                if (elements.filterLncel) elements.filterLncel.value = 'ALL';
                if (elements.filterBusyHour) elements.filterBusyHour.checked = false;
                if (elements.filterDegraded) elements.filterDegraded.checked = false;
                if (elements.filterLowThroughputLowUtilization) elements.filterLowThroughputLowUtilization.checked = false;
                if (elements.tableSearch) elements.tableSearch.value = '';
                updateOperatorPillActive();
                updateSiteAndLncelDropdowns();
                applyFilters();
            });
        }

        // Table Search & Pagination
        if (elements.tableSearch) {
            elements.tableSearch.addEventListener('input', () => {
                state.pagination.currentPage = 1;
                applyFilters();
            });
        }
        if (elements.prevPageBtn) {
            elements.prevPageBtn.addEventListener('click', () => {
                if (state.pagination.currentPage > 1) {
                    state.pagination.currentPage--;
                    renderTable();
                }
            });
        }
        if (elements.nextPageBtn) {
            elements.nextPageBtn.addEventListener('click', () => {
                state.pagination.currentPage++;
                renderTable();
            });
        }

        // Table Sorting Header clicks
        const tableHeaders = document.querySelectorAll('.lte-table th[data-sort]');
        tableHeaders.forEach(th => {
            th.addEventListener('click', () => {
                const col = th.getAttribute('data-sort');
                if (state.pagination.sortColumn === col) {
                    state.pagination.sortDirection = state.pagination.sortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    state.pagination.sortColumn = col;
                    state.pagination.sortDirection = 'desc';
                }
                renderTable();
            });
        });

        // Export
        if (elements.exportCsvBtn) {
            elements.exportCsvBtn.addEventListener('click', exportFilteredDataToCsv);
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
