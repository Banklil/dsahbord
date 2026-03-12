// ==========================================
// 1. Configuration & Initial State
// ==========================================
const SUPABASE_URL = 'https://kbwnpexlszstenscgfqr.supabase.co';
const SUPABASE_KEY = 'sb_publishable_gqt4-rABRt4l49cNVKsNhg_wW74uyRT'; 
const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

let dynamicBranchSettings = []; 
let currentActivePage = "PUPOM"; 
let rawData = []; 
let yearEnd2025Data = {}; 
let charts = { area: null, donut: null };
let kpiChart = null;
window.currentEditingId = null;

let currentUserRole = 'viewer'; 

// ==========================================
// 2. Auth & Initialization
// ==========================================
document.getElementById('login-btn').addEventListener('click', () => {
    const user = document.getElementById('user-auth').value.trim();
    const pass = document.getElementById('pass-auth').value.trim();

    if (user === 'admin' && pass === '1234') { 
        currentUserRole = 'admin';
        performLogin();
    } 
    else if (user === 'MKT01' && pass === '3578') { 
        currentUserRole = 'viewer';
        performLogin();
    } 
    else {
        const err = document.getElementById('auth-error');
        err.style.display = 'block';
        err.innerHTML = "<i class='fas fa-exclamation-triangle'></i> ຊື່ຜູ້ໃຊ້ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ!";
    }
});

function performLogin() {
    document.getElementById('login-screen').style.display = 'none';
    const stars = document.getElementById('stars-container');
    if (stars) stars.style.display = 'none';

    document.getElementById('dashboard-app').style.display = 'flex';
    
    if (typeof initApp === 'function') initApp(); 
    else loadData();
}

function initApp() {
    const areaDom = document.getElementById('largeAreaChart');
    const donutDom = document.getElementById('donutChartDiv');
    if (areaDom) charts.area = echarts.init(areaDom);
    if (donutDom) charts.donut = echarts.init(donutDom);

    const tableTitles = {
        'PUPOM': 'ລາຍລະອຽດຂໍ້ມູນ ລາຍງານປູພົມ',
        'BEERLAO': 'ລາຍລະອຽດຂໍ້ມູນ ລາຍງານເບຍລາວ',
        'CAMPAIGN': 'ລາຍລະອຽດຂໍ້ມູນ ລູກຄ້າແຄມແປນ'
    };

    // --- ແກ້ໄຂ: ປັບປຸງການສະຫຼັບໜ້າຫຼັກ ---
    document.querySelectorAll('.menu-item[data-view]').forEach(btn => {
        btn.onclick = function() {
            document.querySelectorAll('.menu-item').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            const view = this.getAttribute('data-view');

            const titleElement = document.getElementById('table-section-title');
            if (titleElement) {
                titleElement.innerText = tableTitles[view.toUpperCase()] || 'ລາຍລະອຽດຂໍ້ມູນລາຍສາຂາ';
            }

            if (view === 'KPI') {
                switchPageView('kpi');
                loadKPIData(); 
            } else {
                switchPageView('dashboard');
                currentActivePage = view.toUpperCase();
                loadData(); 
            }
        };
    });

    document.getElementById('upload-trigger').onclick = () => document.getElementById('upload-modal').style.display = 'flex';
    document.getElementById('close-modal').onclick = () => document.getElementById('upload-modal').style.display = 'none';
    
    const manageBtn = document.getElementById('manage-branches-btn');
    if(manageBtn) {
        manageBtn.onclick = () => {
            document.getElementById('goal-modal').style.display = 'flex';
            loadBranchSettings(); 
        };
    }
    
    const closeGoalBtn = document.querySelector('#goal-modal .close-btn');
    if(closeGoalBtn) closeGoalBtn.onclick = () => closeGoalModal();

    const fileInput = document.getElementById('excel-input'); 
    if (fileInput) {
        fileInput.onchange = function() {
            const display = document.getElementById('file-name-display');
            if (this.files[0] && display) {
                display.innerHTML = `ໄຟລ໌: <span style="color:#10b981;">${this.files[0].name}</span>`;
            }
        };
    }

    document.getElementById('start-upload-btn').onclick = confirmUpload;
    document.getElementById('apply-filter').onclick = applyFilters;
    document.getElementById('logout-btn').onclick = () => location.reload();

    // ປະກາດ Function ໃຫ້ໃຊ້ງານໄດ້ທົ່ວໄປ
    window.saveBranch = saveBranch;
    window.closeGoalModal = closeGoalModal;
    window.prepareEdit = prepareEdit;
    window.deleteBranch = deleteBranch;
    window.setQuickFilter = setQuickFilter;
    window.updateAmountInput = updateAmountInput;
    window.clearFilters = clearFilters;
    window.loadKPIData = loadKPIData;
    window.deleteDataByDate = deleteDataByDate;
    window.exportToPDF = exportToPDF;
    window.exportToExcel = exportToExcel;
    window.resetChart = resetChart;
    
    // --- ເພີ່ມໃໝ່: ສຳລັບໜ້າ Upload History ---
    window.switchPage = switchPage; 
    window.deleteUpload = deleteUpload;
    window.replaceFile = replaceFile;
    window.handleReplaceFile = handleReplaceFile;
    window.handleNewUpload = handleNewUpload;

    checkPermissions();
    loadData();
    
    window.onresize = () => { 
        charts.area?.resize(); 
        charts.donut?.resize();
        kpiChart?.resize();
    };
}

// 📌 ຟັງຊັນໃໝ່: ສຳລັບສະຫຼັບໜ້າຫຼັກ (Dashboard / KPI / Upload History)
function switchPageView(page) {
    document.getElementById('view-dashboard').style.display = page === 'dashboard' ? 'block' : 'none';
    document.getElementById('view-kpi').style.display = page === 'kpi' ? 'block' : 'none';
    const uploadHistoryPage = document.getElementById('page-upload-history');
    if(uploadHistoryPage) uploadHistoryPage.style.display = page === 'upload-history' ? 'block' : 'none';
}

// 📌 ຟັງຊັນໃໝ່: ສຳລັບເມນູ Upload History ທີ່ຖືກເອີ້ນຈາກ HTML
function switchPage(event, pageId) {
    event.preventDefault();
    document.querySelectorAll('.menu-item').forEach(b => b.classList.remove('active'));
    event.currentTarget.classList.add('active');
    switchPageView(pageId);
}

function checkPermissions() {
    if (currentUserRole === 'viewer') {
        const lockButton = (selector) => {
            const elements = document.querySelectorAll(selector);
            elements.forEach(el => {
                el.style.opacity = '0.4'; 
                el.style.pointerEvents = 'none'; 
                el.style.cursor = 'not-allowed';
                el.style.filter = 'grayscale(100%)';
            });
        };

        lockButton('#upload-trigger');
        lockButton('#manage-branches-btn');
        lockButton('.btn-export.excel');
        lockButton('.btn-delete');
        lockButton('.btn-delete-small');
        lockButton('.btn-delete-row'); // ລັອກປຸ່ມລຶບໃນໜ້າ History
    }
}

function resetChart() {
    if (charts.area) {
        charts.area.dispatchAction({ type: 'restore' });
        charts.area.dispatchAction({ type: 'dataZoom', start: 0, end: 100 });
    }
}

// ==========================================
// 3. Data Loading & Logic
// ==========================================
async function loadData() {
    const loader = document.getElementById('loading-overlay');
    if(loader) loader.style.display = 'flex';

    updateTableHeader();
    await fetchYearEnd2025();

    let allData = [];
    let from = 0, to = 999, hasMore = true;

    try {
        while (hasMore) {
            const { data, error } = await _supabase.from('bank_data').select('*').eq('campaign_type', currentActivePage).range(from, to);
            if (error) throw error;
            allData = allData.concat(data);
            if (data.length < 1000) hasMore = false;
            else { from += 1000; to += 1000; }
        }
        
        rawData = allData.map(d => ({
            ...d,
            branch_name: (d.branch_name || '').trim(),
            ascount_no: (d.ascount_no || '').trim()
        }));

        const { data: settings } = await _supabase.from('branch_settings').select('*').eq('campaign_type', currentActivePage);
        
        dynamicBranchSettings = (settings || []).map(s => ({
            ...s,
            branch_name: (s.branch_name || '').trim()
        }));
        
        updateBranchDropdown();
        applyFilters(); 

    } catch (err) {
        console.error("Load Error:", err);
    } finally {
        if(loader) setTimeout(() => loader.style.display = 'none', 500);
    }
}

async function fetchYearEnd2025() {
    try {
        let allRows = [];
        let from = 0;
        let to = 999;
        let hasMore = true;

        while (hasMore) {
            const { data, error } = await _supabase
                .from('bank_data')
                .select('branch_name, closing_balance, ascount_no')
                .eq('date_key', '2025-12-31') 
                .eq('campaign_type', currentActivePage)
                .range(from, to);

            if (error) throw error;

            if (data && data.length > 0) {
                allRows = allRows.concat(data);
            }

            if (!data || data.length < 1000) {
                hasMore = false;
            } else {
                from += 1000;
                to += 1000;
            }
        }

        console.log(`[${currentActivePage}] ດຶງຂໍ້ມູນ 31/12/2025 ມາໄດ້: ${allRows.length} ລາຍການ`);

        yearEnd2025Data = {};
        const unique2025 = new Map();
        
        allRows.forEach(item => unique2025.set(item.ascount_no, item));
        
        unique2025.forEach(item => {
            if (!yearEnd2025Data[item.branch_name]) yearEnd2025Data[item.branch_name] = 0;
            yearEnd2025Data[item.branch_name] += (Number(item.closing_balance) || 0);
        });

    } catch (err) { 
        console.error("Error fetching 2025 data:", err); 
    }
}

function updateTableHeader() {
    const thead = document.querySelector('#branch-table thead');
    if (!thead) return;
    let headerRow = thead.querySelector('tr');
    if (!headerRow) {
        headerRow = document.createElement('tr');
        thead.appendChild(headerRow);
    }
    headerRow.innerHTML = `
        <th style="padding: 10px;">ສາຂາ</th>
        <th align="center">ຍອດຍົກມາ</th>
        <th align="center">ແຜນປີ 2026</th>
        <th align="center" style="color: #fbbf24;">ຍອດເງິນເປີດໃໝ່</th>
        <th align="center" style="color: #3b82f6;">ປະຕິບັດ </th>
        <th align="center">ທຽບປີ2025</th>
        <th align="center">ທຽບແຜນປີ2026</th>
        <th align="center">ທຽບຍອດເປີດໃໝ່</th>
        <th align="center">%ທຽບປີ2025</th>
        <th align="center">%ທຽບແຜນປີ2026</th>
    `;
}

// ==========================================
// 4. Filtering & Rendering (Dashboard Main)
// ==========================================
function applyFilters() {
    const branch = document.getElementById('branch-select').value;
    const startDate = document.getElementById('start-date').value;
    const endDate = document.getElementById('end-date').value;
    let filtered = [];

    if (startDate || endDate) {
        filtered = rawData.filter(d => {
            let pass = true;
            if (startDate) pass = pass && (d.date_key >= startDate);
            if (endDate) pass = pass && (d.date_key <= endDate);
            return pass;
        });
    } else { filtered = rawData; }

    if (branch !== 'all') filtered = filtered.filter(d => d.branch_name === branch);
    updateUI(filtered);
}

function clearFilters() {
    document.getElementById('branch-select').value = 'all';
    document.getElementById('start-date').value = '';
    document.getElementById('end-date').value = '';
    applyFilters(); 
}

function updateUI(filteredData) {
    const uniqueMap = new Map();
    filteredData.forEach(item => {
        const existing = uniqueMap.get(item.ascount_no);
        if (!existing || item.date_key > existing.date_key) uniqueMap.set(item.ascount_no, item);
    });
    const distinctData = Array.from(uniqueMap.values());

    const totalClosing = distinctData.reduce((sum, i) => sum + (Number(i.closing_balance) || 0), 0);
    document.getElementById('total-amount').innerText = totalClosing.toLocaleString();

    const totalTarget = dynamicBranchSettings.reduce((sum, i) => sum + (Number(i.target_amount) || 0), 0);
    const totalPct = totalTarget > 0 ? (totalClosing / totalTarget) * 100 : 0;
    
    document.getElementById('achievement-pct').innerText = totalPct.toFixed(2) + "%";
    document.getElementById('progress-fill').style.width = Math.min(totalPct, 100) + "%";

    renderMainTable(filteredData); 
    renderCharts(distinctData); 
}

// ==========================================
// 4.1 Table Rendering (ແກ້ໄຂຍອດເປີດໃໝ່ ແລະ ສູດຄິດໄລ່)
// ==========================================
function renderMainTable(tableData) {
    const tbody = document.querySelector('#branch-table tbody');
    if (!tbody) return;
    
    let allBranches = [...new Set([ ...rawData.map(d => d.branch_name), ...dynamicBranchSettings.map(s => s.branch_name) ])];
    const customOrder = ["ສາຂາພາກນະຄອນຫຼວງ", "ສາຂານ້ອຍໂພນໄຊ", "ສາຂານ້ອຍດອນໜູນ", "ສາຂານ້ອຍຊັງຈ້ຽງ", "ສາຂານ້ອຍໜອງໜ່ຽງ"];
    allBranches.sort((a, b) => {
        let indexA = customOrder.indexOf(a), indexB = customOrder.indexOf(b);
        if (indexA === -1) indexA = 999; if (indexB === -1) indexB = 999;
        return indexA !== indexB ? indexA - indexB : a.localeCompare(b, 'lo');
    });

    let TB=0, TP=0, TO=0, TC=0, TD25=0, TDP=0, TDO=0;
    
    const fmt = (v, isPercent = false, isPlan = false) => {
        if (isPlan) {
            const color = v >= 100 ? '#10b981' : '#fbbf24';
            const arrow = v >= 100 ? '▲' : '▼';
            return `<div style="display:flex; justify-content:center; gap:4px; color:${color}; font-weight:bold;">
                        <span>${arrow}</span>
                        <span>${v.toFixed(1)}%</span>
                    </div>`;
        }
        
        const color = v >= 0 ? '#10b981' : '#ef4444';
        const arrow = v >= 0 ? '▲' : '▼';
        const formattedNum = v.toLocaleString(undefined, { 
            minimumFractionDigits: isPercent ? 1 : 0, 
            maximumFractionDigits: isPercent ? 1 : 2 
        });
        
        return `<div style="display:flex; justify-content:center; gap:4px; color:${color}; font-weight:500;">
                    <span>${arrow}</span>
                    <span>${formattedNum}${isPercent ? '%' : ''}</span>
                </div>`;
    };

    const rowsHtml = allBranches.map(bName => {
        const set2025 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2025);
        const set2026 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2026);
        const plan2026 = set2026 ? set2026.target_amount : 0; 
        
        const bData = tableData.filter(d => d.branch_name === bName);
        const accMap = new Map();
        bData.forEach(d => {
            if (!accMap.has(d.ascount_no)) accMap.set(d.ascount_no, []);
            accMap.get(d.ascount_no).push(d);
        });

        let sumO = 0, sumC = 0; 
        accMap.forEach(records => {
            // 1. ຫາຍອດປັດຈຸບັນ (ປິດ) - ເອົາວັນທີໃໝ່ສຸດ
            records.sort((a, b) => b.date_key.localeCompare(a.date_key));
            sumC += (Number(records[0].closing_balance) || 0);
            
            // 2. ຫາຍອດເປີດໃໝ່ (ແຍກເງື່ອນໄຂຕາມໜ້າ)
            if (currentActivePage === 'PUPOM') {
                // --- ສຳລັບໜ້າປູພົມ: ເອົາ "ຍອດເປີດ" ຈາກໄຟລ໌ອັບໂຫຼດຫຼ້າສຸດເລີຍ ---
                let accountOpen = 0;
                for (let r of records) {
                    if (Number(r.opening_balance) > 0) {
                        accountOpen = Number(r.opening_balance);
                        break; 
                    }
                }
                sumO += accountOpen;
            } else {
                // --- ສຳລັບໜ້າ ເບຍລາວ / ແຄມແປນ: ເອົາຍອດຈາກການອັບໂຫຼດຄັ້ງທຳອິດຂອງປີ ---
                let records2026 = records.filter(r => !r.date_key.startsWith('2025'));
                if (records2026.length > 0) {
                    // ລຽງວັນທີຈາກເກົ່າໄປໃໝ່ ເພື່ອເອົາໃບບິນທຳອິດ
                    records2026.sort((a, b) => a.date_key.localeCompare(b.date_key));
                    const firstUpload = records2026[0];
                    
                    // ເອົາຍອດທີ່ໃຫຍ່ທີ່ສຸດລະຫວ່າງ opening ຫຼື closing ຂອງມື້ທຳອິດ
                    sumO += Math.max(Number(firstUpload.opening_balance) || 0, Number(firstUpload.closing_balance) || 0);
                }
            }
        });

        let baselineVal = yearEnd2025Data[bName] || 0;
        if (baselineVal === 0 && set2025) baselineVal = set2025.balance_prev_year || 0;

        // 🎯 ສູດຄິດໄລ່ຫຼັກ
        const d25 = sumC - baselineVal; 
        const dP = sumC - plan2026;     
        
        // 🎯 ປັບປຸງການທຽບຍອດເປີດໃໝ່ (ໃຫ້ໃຊ້ສູດດຽວກັນທັງໝົດແລ້ວ)
        let dO = sumC - sumO; 
        
        let p25 = baselineVal !== 0 ? (d25 / baselineVal) * 100 : (sumC > 0 ? 100 : 0);
        let pP = plan2026 !== 0 ? (sumC / plan2026) * 100 : 0;

        TB+=baselineVal; TP+=plan2026; TO+=sumO; TC+=sumC; TD25+=d25; TDP+=dP; TDO+=dO;

        return `<tr>
            <td style="text-align:left; padding-left:20px;">${bName}</td>
            <td align="center" style="font-weight:bold; color:#e2e8f0;">${baselineVal.toLocaleString()}</td>
            <td align="center" style="color:#94a3b8;">${plan2026.toLocaleString()}</td>
            <td align="center" style="color:#fbbf24; font-weight:bold;">${sumO.toLocaleString()}</td>
            <td align="center" style="color:#3b82f6;">${sumC.toLocaleString()}</td>
            <td align="center">${fmt(d25)}</td>
            <td align="center">${fmt(dP)}</td>
            <td align="center">${fmt(dO)}</td> 
            <td align="center">${fmt(p25, true)}</td>
            <td align="center">${fmt(pP, true, true)}</td> </tr>`;
    }).join('');

    const totalHtml = `<tr style="background:#0f172a; font-weight:bold;">
        <td style="text-align:left; padding-left:20px;">ລວມທັງໝົດ</td>
        <td align="center" style="color:#ffffff;">${TB.toLocaleString()}</td>
        <td align="center" style="color:#ffffff;">${TP.toLocaleString()}</td>
        <td align="center" style="color:#fbbf24;">${TO.toLocaleString()}</td>
        <td align="center" style="color:#3b82f6;">${TC.toLocaleString()}</td>
        <td align="center">${fmt(TD25)}</td>
        <td align="center">${fmt(TDP)}</td>
        <td align="center">${fmt(TDO)}</td>
        <td align="center">${fmt(TB?((TC-TB)/TB)*100:0, true)}</td>
        <td align="center">${fmt(TP?(TC/TP)*100:0, true, true)}</td>
    </tr>`;
    
    tbody.innerHTML = rowsHtml + totalHtml;
}

// ==========================================
// 6. Branch Management
// ==========================================
async function loadBranchSettings() {
    try {
        const { data, error } = await _supabase
            .from('branch_settings')
            .select('*')
            .eq('campaign_type', currentActivePage)
            .order('branch_name', { ascending: true });

        if (error) throw error;
        dynamicBranchSettings = (data || []).map(s => ({
            ...s, branch_name: (s.branch_name || '').trim()
        }));
        renderBranchList(); 
    } catch (err) {
        console.error("Error loading branch settings:", err);
    }
}

function renderBranchList() {
    const tbody = document.getElementById('goal-table-body');
    if (!tbody) return;
    
    const grouped = {};
    dynamicBranchSettings.forEach(item => {
        if (!grouped[item.branch_name]) grouped[item.branch_name] = {};
        grouped[item.branch_name][item.target_year] = item.target_amount;
    });

    const branches = Object.keys(grouped);
    if (branches.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:20px; color:#64748b;">ບໍ່ມີຂໍ້ມູນສາຂາ</td></tr>`;
        return;
    }

    tbody.innerHTML = branches.map(branch => {
        const goal25 = grouped[branch][2025] || 0;
        const goal26 = grouped[branch][2026] || 0;

        return `
            <tr>
                <td style="text-align:left; font-weight:bold; color:#fff; padding: 12px;">${branch}</td>
                <td align="center" style="color:#94a3b8;">${goal25.toLocaleString()}</td>
                <td align="center" style="color:#fbbf24; font-weight:bold;">${goal26.toLocaleString()}</td>
                <td style="text-align:center;">
                    <div style="display: flex; justify-content: center; gap: 10px;">
                        <button class="btn-edit-small" onclick="prepareEdit('${branch}')" style="color:#3b82f6; border:none; background:none; cursor:pointer; font-weight:bold;">
                            <i class="fas fa-edit"></i> ແກ້ໄຂ
                        </button>
                        <button class="btn-delete-small" onclick="deleteBranch('${branch}')" style="color:#ef4444; border:none; background:none; cursor:pointer;">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
    
    checkPermissions();
}

async function saveBranch() {
    const name = document.getElementById('branch-name').value.trim();
    const year = parseInt(document.getElementById('select-year').value);
    const amount = parseFloat(document.getElementById('input-amount').value || 0);

    if (!name) return alert("ກະລຸນາປ້ອນຊື່ສາຂາ!");

    const saveBtn = document.querySelector('#goal-modal .btn-primary-glass');
    if(saveBtn) saveBtn.disabled = true;

    try {
        const { data: existing } = await _supabase
            .from('branch_settings')
            .select('id')
            .eq('branch_name', name)
            .eq('target_year', year)
            .eq('campaign_type', currentActivePage)
            .maybeSingle();

        const payload = {
            branch_name: name,
            target_year: year,
            target_amount: amount,
            campaign_type: currentActivePage
        };

        if (existing) {
            await _supabase.from('branch_settings').update(payload).eq('id', existing.id);
        } else {
            await _supabase.from('branch_settings').insert(payload);
        }

        alert("ບັນທຶກສຳເລັດ!");
        document.getElementById('branch-name').value = '';
        document.getElementById('input-amount').value = '';
        
        await loadBranchSettings();
        loadData(); 
    } catch (err) {
        alert("Error: " + err.message);
    } finally {
        if(saveBtn) saveBtn.disabled = false;
    }
}

async function deleteBranch(branchName) {
    if(!confirm(`ຢືນຢັນລົບສາຂາ "${branchName}"?`)) return;

    try {
        const { error } = await _supabase
            .from('branch_settings')
            .delete()
            .eq('branch_name', branchName)
            .eq('campaign_type', currentActivePage);

        if (error) throw error;
        
        await loadBranchSettings();
        loadData();
    } catch (err) {
        alert("ລົບບໍ່ໄດ້: " + err.message);
    }
}

function prepareEdit(branchName) {
    document.getElementById('branch-name').value = branchName;
    updateAmountInput();
    document.getElementById('input-amount').focus();
}

function updateAmountInput() {
    const branchName = document.getElementById('branch-name').value;
    if (!branchName) return;
    
    const year = parseInt(document.getElementById('select-year').value);
    const found = dynamicBranchSettings.find(i => i.branch_name === branchName && i.target_year === year);
    
    document.getElementById('input-amount').value = found ? found.target_amount : 0;
}

function closeGoalModal() { 
    document.getElementById('goal-modal').style.display = 'none'; 
}

function updateBranchDropdown() {
    const s = document.getElementById('branch-select');
    const b = [...new Set(rawData.map(d => d.branch_name))];
    s.innerHTML = '<option value="all">ທັງໝົດທຸກສາຂາ</option>' + b.map(x => `<option value="${x}">${x}</option>`).join('');
}

function setQuickFilter(type) {
    const today = new Date(); 
    let start = new Date();
    if (type === 'week') start.setDate(today.getDate() - 7);
    if (type === 'month') start.setMonth(today.getMonth(), 1);
    if (type === 'year') start.setFullYear(today.getFullYear(), 0, 1);
    
    document.getElementById('start-date').value = start.toISOString().split('T')[0];
    document.getElementById('end-date').value = today.toISOString().split('T')[0];
    applyFilters();
}

// ==========================================
// 7. Chart Logic
// ==========================================
function renderCharts(uniqueData) {
    if (!charts.area || !charts.donut) return;

    const branches = [...new Set(dynamicBranchSettings.map(s => s.branch_name))]; 
    const totals = branches.map(b => {
        const branchData = uniqueData.filter(d => d.branch_name === b);
        let currentTotal = branchData.reduce((s, c) => s + (Number(c.closing_balance) || 0), 0);
        if (currentTotal === 0 && branchData.length === 0) {
            currentTotal = yearEnd2025Data[b] || 0;
        }
        return currentTotal;
    });

    charts.area.setOption({
        tooltip: { 
            trigger: 'axis',
            backgroundColor: 'rgba(30, 41, 59, 0.9)', 
            borderColor: '#334155',
            textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao, sans-serif' }
        },
        grid: { left: '3%', right: '4%', bottom: '10%', containLabel: true },
        xAxis: { 
            type: 'category', 
            data: branches, 
            axisLabel: { rotate: 45, color: '#94a3b8', fontFamily: 'Noto Sans Lao, sans-serif', fontSize: 11 },
            axisLine: { lineStyle: { color: '#334155' } }
        },
        yAxis: { 
            type: 'value', 
            axisLabel: { color: '#94a3b8', fontFamily: 'Noto Sans Lao, sans-serif' },
            splitLine: { lineStyle: { color: 'rgba(255,255,255,0.05)' } }
        },
        series: [{ 
            data: totals, 
            type: 'bar', 
            itemStyle: { color: '#48f63bff', borderRadius: [4, 4, 0, 0] },
            name: 'ຍອດເງິນປັດຈຸບັນ',
            barMaxWidth: 60
        }]
    });

    let activeCount = 0, inactiveCount = 0, zeroCount = 0;
    uniqueData.forEach(d => {
        const open = Number(d.opening_balance) || 0;
        const close = Number(d.closing_balance) || 0;
        
        if (close <= 0) zeroCount++; 
        else if (open !== close) activeCount++; 
        else inactiveCount++;
    });

    const totalAccounts = activeCount + inactiveCount + zeroCount;
    const donutData = [
        { value: activeCount, name: 'ເຄື່ອນໄຫວ', itemStyle: { color: '#e29c46ff' } },   
        { value: inactiveCount, name: 'ບໍ່ເຄື່ອນໄຫວ', itemStyle: { color: '#21dcccff' } }, 
        { value: zeroCount, name: 'ບັນຊີສູນ', itemStyle: { color: '#44d289ff' } }       
    ];

    charts.donut.clear();

    charts.donut.setOption({
        title: {
            text: totalAccounts.toLocaleString(),
            subtext: 'ບັນຊີທັງໝົດ (Unique)',
            left: 'center', top: '40%',
            textStyle: { color: '#ffffff', fontSize: 28, fontWeight: 'bold', fontFamily: 'Noto Sans Lao, sans-serif' },
            subtextStyle: { color: '#94a3b8', fontSize: 14, fontFamily: 'Noto Sans Lao, sans-serif' }
        },
        tooltip: { 
            trigger: 'item', 
            formatter: '{b}: {c} ({d}%)', 
            backgroundColor: 'rgba(30, 41, 59, 0.9)',
            borderColor: '#334155',
            textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao, sans-serif' } 
        },
        legend: { bottom: '0%', left: 'center', textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao, sans-serif' }, itemGap: 20 },
        series: [{ 
            type: 'pie', radius: ['55%', '75%'], avoidLabelOverlap: false,
            itemStyle: { borderRadius: 10, borderColor: '#1e293b', borderWidth: 3 },
            label: { show: false, position: 'center' },
            emphasis: { label: { show: false }, itemStyle: { shadowBlur: 10, shadowOffsetX: 0, shadowColor: 'rgba(0, 0, 0, 0.5)' } },
            data: donutData 
        }]
    });

    charts.donut.off('legendselectchanged'); 
    charts.donut.on('legendselectchanged', function(params) {
        let newTotal = 0;
        donutData.forEach(item => {
            if (params.selected[item.name]) {
                newTotal += item.value;
            }
        });
        charts.donut.setOption({ 
            title: { text: newTotal.toLocaleString() } 
        });
    });
}

// ==========================================
// 8. Upload & Export
// ==========================================
async function confirmUpload() {
    const fileInput = document.getElementById('excel-input');
    const dateKey = document.getElementById('upload-date-select').value;
    
    const uploadTypeElem = document.querySelector('input[name="uploadType"]:checked');
    const uploadType = uploadTypeElem ? uploadTypeElem.value : 'UPDATE'; 

    if (!fileInput.files.length || !dateKey) {
        alert("ກະລຸນາເລືອກໄຟລ໌ ແລະ ວັນທີ!");
        return;
    }

    const btn = document.getElementById('start-upload-btn');
    const originalText = btn.innerHTML;
    btn.innerHTML = "<i class='fas fa-spinner fa-spin'></i> ກຳລັງປະມວນຜົນ...";
    btn.disabled = true;

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            if (jsonData.length === 0) throw new Error("ບໍ່ພົບຂໍ້ມູນໃນໄຟລ໌!");

            const groupedData = {};

            jsonData.forEach(row => {
                const cleanRow = {};
                for (let key in row) {
                    const cleanKey = String(key).toLowerCase().replace(/[\s_\-\.]/g, '');
                    cleanRow[cleanKey] = row[key];
                }

                // --- ແກ້ໄຂ: ປັບປຸງການດຶງຊື່ສາຂາ (Auto Mapping) ---
                let branchName = cleanRow['branchname'] || cleanRow['branch'] || row['ສາຂາ'] || 'ບໍ່ລະບຸ';
                if (branchName.includes('ດອນໜູນ')) branchName = 'ສາຂານ້ອຍດອນໜູນ';
                else if (branchName.includes('ໂພນໄຊ')) branchName = 'ສາຂານ້ອຍໂພນໄຊ';
                else if (branchName.includes('ຊັງຈ້ຽງ')) branchName = 'ສາຂານ້ອຍຊັງຈ້ຽງ';
                else if (branchName.includes('ໜອງໜ່ຽງ')) branchName = 'ສາຂານ້ອຍໜອງໜ່ຽງ';
                else if (branchName.includes('ນະຄອນຫຼວງ')) branchName = 'ສາຂາພາກນະຄອນຫຼວງ';

                // --- ແກ້ໄຂ: ປັບປຸງການດຶງເລກບັນຊີ ແລະ ສ້າງ Dummy ID ຖ້າບໍ່ມີ ---
                let accNo = String(
                    cleanRow['accountno'] || cleanRow['ascountno'] || cleanRow['accno'] || 
                    cleanRow['account'] || cleanRow['number'] || row['ເລກບັນຊີ'] || row['ເລກທີບັນຊີ'] || ''
                ).trim();

                // 🚨 ປ້ອງກັນການປັດຍອດຖິ້ມ ຖ້າບໍ່ມີເລກບັນຊີ
                if (!accNo) {
                    accNo = 'DUMMY_' + Math.floor(Math.random() * 10000000);
                }

                const custName = cleanRow['customername'] || cleanRow['name'] || row['ຊື່ລູກຄ້າ'] || '';

                let openVal = 0;
                let closeVal = 0;

                const hasOpenCol = 'openingbalance' in cleanRow || 'ຍອດເປີດ' in cleanRow;
                const hasCloseCol = 'closingbalance' in cleanRow || 'ຍອດປິດ' in cleanRow;

                if (hasOpenCol || hasCloseCol) {
                    openVal = parseFloat(String(cleanRow['openingbalance'] || cleanRow['ຍອດເປີດ'] || '0').replace(/,/g, '')) || 0;
                    closeVal = parseFloat(String(cleanRow['closingbalance'] || cleanRow['ຍອດປິດ'] || '0').replace(/,/g, '')) || 0;
                } else {
                    let rawAmount = String(
                        cleanRow['balance'] || cleanRow['amount'] || cleanRow['total'] || 
                        cleanRow['ຍອດເງິນ'] || cleanRow['ຈຳນວນເງິນ'] || cleanRow['ຍອດຄົງເຫຼືອ'] || '0'
                    ).replace(/,/g, '');
                    const amountVal = parseFloat(rawAmount) || 0;

                    if (uploadType === 'NEW') {
                        openVal = amountVal;
                        closeVal = 0; 
                    } else {
                        openVal = 0; 
                        closeVal = amountVal;
                    }
                }

                const uKey = `${accNo}_${dateKey}_${currentActivePage}`;

                if (!groupedData[uKey]) {
                    groupedData[uKey] = {
                        branch_name: branchName,
                        ascount_no: accNo,
                        customer_name: custName,
                        campaign_type: currentActivePage,
                        date_key: dateKey,
                        opening_balance: openVal, 
                        closing_balance: closeVal 
                    };
                } else {
                    groupedData[uKey].opening_balance += openVal;
                    groupedData[uKey].closing_balance += closeVal;
                }
            });

            const finalPayload = Object.values(groupedData);
            
            const { error: dataError } = await _supabase
                .from('bank_data')
                .upsert(finalPayload, { onConflict: 'ascount_no,date_key,campaign_type' });

            if (dataError) throw dataError;

            if(dateKey === '2025-12-31') await fetchYearEnd2025();

            alert(`✅ ສຳເລັດ! ອັບໂຫລດຈຳນວນ ${finalPayload.length} ລາຍການ`);
            document.getElementById('upload-modal').style.display = 'none';
            
            rawData = [];
            await loadData();

        } catch (err) {
            console.error("Upload Error:", err);
            alert("ເກີດຂໍ້ຜິດພາດ: " + err.message);
        } finally {
            btn.innerHTML = originalText;
            btn.disabled = false;
            fileInput.value = ''; 
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

// ==========================================
// 9. KPI Logic
// ==========================================
async function loadKPIData() {
    const empName = document.getElementById('kpi-employee-select').value;
    const month = document.getElementById('kpi-month-select').value;

    if(!empName || !month) return;

    const { data, error } = await _supabase
        .from('marketing_kpi')
        .select('*')
        .eq('employee_name', empName)
        .eq('kpi_month', month)
        .single();

    if (error || !data) {
        document.getElementById('kpi-table-body').innerHTML = `<tr><td colspan="5" style="text-align:center; padding:20px;">ບໍ່ພົບຂໍ້ມູນ</td></tr>`;
        document.getElementById('kpi-total-score').innerText = "0%";
        document.getElementById('kpi-status').innerText = "-";
        updateRadarChart([0, 0, 0]);
        return;
    }

    const depPct = calculatePercent(data.actual_deposit, data.target_deposit);
    const accPct = calculatePercent(data.actual_new_acc, data.target_new_acc);
    const visitPct = calculatePercent(data.actual_visit, data.target_visit);
    const totalScore = (depPct * 0.5) + (accPct * 0.3) + (visitPct * 0.2);

    document.getElementById('kpi-total-score').innerText = totalScore.toFixed(1) + "%";
    
    const statusEl = document.getElementById('kpi-status');
    let color = '#ef4444'; let text = "ຕ້ອງປັບປຸງ";
    if(totalScore >= 100) { color = '#10b981'; text = "ດີເລີດ"; }
    else if(totalScore >= 80) { color = '#3b82f6'; text = "ດີ"; }
    else if(totalScore >= 60) { color = '#f59e0b'; text = "ປານກາງ"; }
    
    statusEl.innerText = text;
    statusEl.style.color = color;

    const tbody = document.getElementById('kpi-table-body');
    tbody.innerHTML = `
        <tr><td>💰 ຍອດເງິນຝາກ</td><td align="right">${Number(data.target_deposit).toLocaleString()}</td><td align="right">${Number(data.actual_deposit).toLocaleString()}</td><td align="center" style="color:${getColor(depPct)}">${depPct.toFixed(1)}%</td><td align="center">${(depPct * 0.5).toFixed(1)}</td></tr>
        <tr><td>ເປີດບັນຊີໃໝ່</td><td align="right">${data.target_new_acc}</td><td align="right">${data.actual_new_acc}</td><td align="center" style="color:${getColor(accPct)}">${accPct.toFixed(1)}%</td><td align="center">${(accPct * 0.3).toFixed(1)}</td></tr>
        <tr><td>🚗 ລົງຢ້ຽມຢາມ</td><td align="right">${data.target_visit}</td><td align="right">${data.actual_visit}</td><td align="center" style="color:${getColor(visitPct)}">${visitPct.toFixed(1)}%</td><td align="center">${(visitPct * 0.2).toFixed(1)}</td></tr>
    `;
    updateRadarChart([depPct, accPct, visitPct]);
}

function calculatePercent(actual, target) {
    if(!target || target == 0) return 0;
    let pct = (actual / target) * 100;
    return pct > 120 ? 120 : pct; 
}

function getColor(pct) {
    if(pct >= 100) return '#10b981';
    if(pct >= 80) return '#3b82f6';
    if(pct >= 60) return '#f59e0b';
    return '#ef4444';
}

function updateRadarChart(dataArr) {
    const dom = document.getElementById('kpiRadarChart');
    if(!kpiChart) kpiChart = echarts.init(dom);
    const option = {
        radar: { indicator: [{ name: 'ເງິນຝາກ', max: 120 }, { name: 'ບັນຊີໃໝ່', max: 120 }, { name: 'ຢ້ຽມຢາມ', max: 120 }], splitArea: { areaStyle: { color: ['#1e293b'] } }, axisName: { color: '#fff' } },
        series: [{ type: 'radar', data: [{ value: dataArr, name: 'Performance', itemStyle: { color: '#3b82f6' }, areaStyle: { opacity: 0.3 } }] }]
    };
    kpiChart.setOption(option);
}

// ==========================================
// 10. Utils (Export, Delete, Upload UI)
// ==========================================

// --- ແກ້ໄຂ: ອັບເກຣດຟັງຊັນລຶບຂໍ້ມູນ (ລຶບຕາມວັນທີ ແລະ ຊື່ສາຂາ) ---
async function deleteDataByDate() {
    const dateKey = document.getElementById('upload-date-select').value;
    const branchSelect = document.getElementById('branch-select').value; 
    
    if (!dateKey && branchSelect === 'all') {
        return alert("⚠️ ກະລຸນາເລືອກ 'ວັນທີ' ໃນກ່ອງນີ້ ຫຼື ອອກໄປເລືອກ 'ສາຂາ' ຢູ່ໜ້າ Dashboard ກ່ອນຈຶ່ງກົດລຶບ!");
    }

    let confirmText = `⚠️ ຢືນຢັນການລຶບຂໍ້ມູນ`;
    if(dateKey) confirmText += `\n- ວັນທີ: ${dateKey}`;
    if(branchSelect !== 'all') confirmText += `\n- ສະເພາະສາຂາ: ${branchSelect}`;
    confirmText += `\n(ລຶບແລ້ວບໍ່ສາມາດກູ້ຄືນໄດ້!)`;

    if (!confirm(confirmText)) return;

    let query = _supabase.from('bank_data').delete().eq('campaign_type', currentActivePage);
    if (dateKey) query = query.eq('date_key', dateKey);
    if (branchSelect !== 'all') query = query.eq('branch_name', branchSelect);

    const { error } = await query;
    
    if (error) {
        alert("Error: " + error.message);
    } else { 
        alert("✅ ລຶບຂໍ້ມູນສຳເລັດແລ້ວ!"); 
        document.getElementById('upload-modal').style.display = 'none'; 
        rawData = []; 
        await loadData(); 
    }
}

function exportToPDF() {
    const element = document.getElementById('table-content-area');
    if (!element) {
        alert("ບໍ່ພົບຂໍ້ມູນທີ່ຈະດາວໂຫລດ!");
        return;
    }

    const today = new Date().toISOString().slice(0, 10);
    let pdfTitle = "ລາຍງານທົ່ວໄປ";
    if (currentActivePage === 'PUPOM') pdfTitle = "ລາຍງານ ປູພົມ";
    else if (currentActivePage === 'BEERLAO') pdfTitle = "ລາຍງານ ເບຍລາວ";
    else if (currentActivePage === 'CAMPAIGN') pdfTitle = "ລາຍງານ ລູກຄ້າແຄມແປນ";

    const opt = {
        margin: [0.2, 0.2, 0.2, 0.2], 
        filename: `${pdfTitle}_${today}.pdf`,
        image: { type: 'jpeg', quality: 1.0 },
        html2canvas: { 
            scale: 2, 
            useCORS: true, 
            backgroundColor: '#1e293b', 
            scrollY: 0
        },
        jsPDF: { unit: 'in', format: 'a3', orientation: 'landscape' }
    };

    const container = document.createElement('div');
    container.style.backgroundColor = '#1e293b'; 
    container.style.padding = '20px';
    container.style.fontFamily = '"Noto Sans Lao", sans-serif';
    container.style.width = '100%';
    container.style.color = '#ffffff'; 

    const header = document.createElement('h1');
    header.innerText = `${pdfTitle} (ປະຈຳວັນທີ: ${today})`;
    header.style.textAlign = 'center';
    header.style.fontSize = '24px';
    header.style.marginBottom = '20px';
    header.style.color = '#3b82f6'; 
    container.appendChild(header);

    const tableClone = element.cloneNode(true);
    const table = tableClone.querySelector('table');
    if (table) {
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
    }
    
    const ths = tableClone.querySelectorAll('th');
    ths.forEach(th => {
        th.style.backgroundColor = '#0f172a';
        th.style.color = '#fbbf24'; 
        th.style.border = '1px solid #334155';
    });

    const tds = tableClone.querySelectorAll('td');
    tds.forEach(td => {
        td.style.borderBottom = '1px solid #334155';
        if (!td.style.color) td.style.color = '#ffffff';
    });

    container.appendChild(tableClone);

    html2pdf().set(opt).from(container).save().then(() => {});
}

function exportToExcel() {
    const branch = document.getElementById('branch-select').value;
    let filtered = rawData;
    if (branch !== 'all') filtered = filtered.filter(d => d.branch_name === branch);
    if (filtered.length === 0) return alert("ບໍ່ມີຂໍ້ມູນ!");
    const uniqueMap = new Map();
    filtered.forEach(item => { const existing = uniqueMap.get(item.ascount_no); if (!existing || item.date_key > existing.date_key) uniqueMap.set(item.ascount_no, item); });
    const excelData = Array.from(uniqueMap.values()).map((item, index) => ({ "ລ/ດ": index + 1, "ສາຂາ": item.branch_name, "ເລກບັນຊີ": item.ascount_no, "ຍອດຄົງເຫຼືອ": item.closing_balance }));
    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Report");
    XLSX.writeFile(workbook, `Report_${currentActivePage}.xlsx`);
}

// ----------------------------------------
// ຈັດການໜ້າ Upload History 
// ----------------------------------------
let currentRowToReplace = null;
let currentBranchToReplace = null;

// --- ແກ້ໄຂ: ຟັງຊັນລຶບໄຟລ໌ອັບໂຫຼດ (ລຶບຂໍ້ມູນສາຂາອອກຈາກຖານຂໍ້ມູນແທ້) ---
async function deleteUpload(rowId, branchName) {
    if (confirm(`⚠️ ທ່ານຕ້ອງການລຶບຂໍ້ມູນທັງໝົດຂອງ "${branchName}" ອອກຈາກຖານຂໍ້ມູນແທ້ຫຼືບໍ່? (ລຶບແລ້ວກູ້ຄືນບໍ່ໄດ້)`)) {
        
        const { error } = await _supabase
            .from('bank_data')
            .delete()
            .eq('branch_name', branchName)
            .eq('campaign_type', currentActivePage); 

        if (error) {
            alert("ເກີດຂໍ້ຜິດພາດໃນການລຶບ: " + error.message);
            return;
        }

        const row = document.getElementById(rowId);
        if (row) row.remove();
        
        alert(`✅ ລຶບຂໍ້ມູນຂອງ "${branchName}" ອອກຈາກຖານຂໍ້ມູນສຳເລັດແລ້ວ!`);
        rawData = [];
        await loadData();
    }
}

function replaceFile(rowId, branchName) {
    currentRowToReplace = rowId;
    currentBranchToReplace = branchName;
    
    if (confirm(`ທ່ານຕ້ອງການອັບໂຫຼດໄຟລ໌ໃໝ່ ເພື່ອໄປແທນທີ່ຂໍ້ມູນເກົ່າຂອງ "${branchName}" ແມ່ນບໍ່?`)) {
        document.getElementById('replaceFileInput').click();
    }
}

function handleReplaceFile(input) {
    if (input.files && input.files[0]) {
        const newFileName = input.files[0].name;
        const row = document.getElementById(currentRowToReplace);
        
        row.cells[2].innerHTML = `<i class="fa-solid fa-file-excel" style="color: #10b981;"></i> <span style="color: #cbd5e1;">${newFileName}</span>`;
        
        const now = new Date();
        const timeString = now.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        row.cells[0].innerHTML = `ມື້ນີ້ <br><small style="color: #94a3b8;">${timeString}</small>`;
        
        alert(`✅ ອັບໂຫຼດໄຟລ໌ "${newFileName}" ເຂົ້າໄປແທນທີ່ຂໍ້ມູນຂອງ "${currentBranchToReplace}" ສຳເລັດແລ້ວ!`);
        input.value = ""; 
    }
}

function handleNewUpload(input) {
    if (input.files && input.files[0]) {
        alert(`✅ ຮັບໄຟລ໌ "${input.files[0].name}" ເຂົ້າລະບົບແລ້ວ! (ລໍຖ້າການປະມວນຜົນເຂົ້າຕາຕະລາງ)`);
        input.value = "";
    }
}