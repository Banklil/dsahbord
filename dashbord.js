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

// ✅ ຕົວປ່ຽນເກັບສະຖານະ (Default)
let currentUserRole = 'viewer'; 

// ==========================================
// 2. Auth & Initialization
// ==========================================
document.getElementById('login-btn').addEventListener('click', () => {
    const user = document.getElementById('user-auth').value.trim();
    const pass = document.getElementById('pass-auth').value.trim();

    // --- 🔐 ກວດສອບ User ---
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
    
    // ປິດດາວ
    const stars = document.getElementById('stars-container');
    if (stars) stars.style.display = 'none';

    document.getElementById('dashboard-app').style.display = 'flex';
    
    if (typeof initApp === 'function') initApp(); 
    else loadData();
}

function initApp() {
    // Init Charts
    const areaDom = document.getElementById('largeAreaChart');
    const donutDom = document.getElementById('donutChartDiv');
    if (areaDom) charts.area = echarts.init(areaDom);
    if (donutDom) charts.donut = echarts.init(donutDom);

   const tableTitles = {
        'PUPOM': 'ລາຍລະອຽດຂໍ້ມູນ ລາຍງານປູພົມ',
        'BEERLAO': 'ລາຍລະອຽດຂໍ້ມູນ ລາຍງານເບຍລາວ',
        'CAMPAIGN': 'ລາຍລະອຽດຂໍ້ມູນ ລູກຄ້າແຄມແປນ'
    };

    // 2. ປັບປຸງສ່ວນການກົດ Menu Item
    document.querySelectorAll('.menu-item').forEach(btn => {
        btn.onclick = function() {
            document.querySelectorAll('.menu-item').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            const view = this.getAttribute('data-view');

            // ✅ ເພີ່ມ Logic ປ່ຽນຊື່ຫົວຂໍ້ຕາຕະລາງບ່ອນນີ້
            const titleElement = document.getElementById('table-section-title');
            if (titleElement) {
                titleElement.innerText = tableTitles[view.toUpperCase()] || 'ລາຍລະອຽດຂໍ້ມູນລາຍສາຂາ';
            }

            if (view === 'KPI') {
                document.getElementById('view-dashboard').style.display = 'none';
                document.getElementById('view-kpi').style.display = 'block';
                loadKPIData(); 
            } else {
                document.getElementById('view-kpi').style.display = 'none';
                document.getElementById('view-dashboard').style.display = 'block';
                currentActivePage = view.toUpperCase();
                
                // (Optional) ຖ້າຢາກໃຫ້ Page Title ທາງເທິງປ່ຽນນຳ
                // document.getElementById('page-title').innerText = "Dashboard " + this.innerText.trim();
                
                loadData(); 
            }
        };
    });

    // Modals
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

    // File Input
    const fileInput = document.getElementById('excel-input'); 
    if (fileInput) {
        fileInput.onchange = function() {
            const display = document.getElementById('file-name-display');
            if (this.files[0] && display) {
                display.innerHTML = `ໄຟລ໌: <span style="color:#10b981;">${this.files[0].name}</span>`;
            }
        };
    }

    // Buttons
    document.getElementById('start-upload-btn').onclick = confirmUpload;
    document.getElementById('apply-filter').onclick = applyFilters;
    document.getElementById('logout-btn').onclick = () => location.reload();

    // Global Functions
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

    // ✅✅✅ ເອີ້ນຟັງຊັນລັອກປຸ່ມ ✅✅✅
    checkPermissions();

    loadData();
    window.onresize = () => { 
        charts.area?.resize(); 
        charts.donut?.resize();
        kpiChart?.resize();
    };
}

// ✅✅✅ ຟັງຊັນຈັດການສິດທິ (User Permission) ✅✅✅
function checkPermissions() {
    // ຖ້າເປັນ viewer ໃຫ້ລັອກປຸ່ມ ແຕ່ຍັງໂຊຢູ່
    if (currentUserRole === 'viewer') {
        
        // ຟັງຊັນຊ່ວຍໃນການລັອກປຸ່ມ (ເຮັດໃຫ້ຈາງ ແລະ ກົດບໍ່ໄດ້)
        const lockButton = (selector) => {
            const elements = document.querySelectorAll(selector);
            elements.forEach(el => {
                el.style.opacity = '0.4';           // ເຮັດໃຫ້ຈາງລົງ (40% ຄວາມເຂັ້ມ)
                el.style.pointerEvents = 'none';    // ❌ ປິດການກົດ (Click) ທຸກຢ່າງ
                el.style.cursor = 'not-allowed';    // ປ່ຽນເມົ້າເປັນຮູບຫ້າມ
                el.style.filter = 'grayscale(100%)'; // ເຮັດໃຫ້ເປັນສີເທົາ (Optional)
            });
        };

        // 1. ລັອກປຸ່ມ ອັບໂຫລດ Data (Sidebar)
        lockButton('#upload-trigger');

        // 2. ລັອກປຸ່ມ ຈັດການສາຂາ/ຕັ້ງເປົ້າໝາຍ
        lockButton('#manage-branches-btn');

        // 3. ລັອກປຸ່ມ Excel (ແຕ່ປ່ອຍ PDF ໄວ້)
        lockButton('.btn-export.excel');
        
        // 4. ລັອກປຸ່ມລ້າງຄ່າ (Optional: ຖ້າຢາກໃຫ້ເບິ່ງໄດ້ແຕ່ຫ້າມລ້າງ)
        // lockButton('.btn-danger-glass'); 
        
        // 5. ລັອກປຸ່ມລົບໃນ Modal (ຖ້າເປີດ Modal ໄດ້)
        lockButton('.btn-delete');
    }
}

function resetChart() {
    if (charts.area) {
        charts.area.dispatchAction({ type: 'restore' });
        charts.area.dispatchAction({ type: 'dataZoom', start: 0, end: 100 });
    }
}

// ==========================================
// 3. Data & Logic (ຄືເກົ່າ)
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
            const { data, error } = await _supabase
                .from('bank_data')
                .select('*')
                .eq('campaign_type', currentActivePage)
                .range(from, to);
            if (error) throw error;
            
            allData = allData.concat(data);
            if (data.length < 1000) hasMore = false;
            else { from += 1000; to += 1000; }
        }
        rawData = allData;

        const { data: settings } = await _supabase
            .from('branch_settings')
            .select('*')
            .eq('campaign_type', currentActivePage);
        
        dynamicBranchSettings = settings || [];
        
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
        const { data, error } = await _supabase
            .from('bank_data')
            .select('branch_name, closing_balance, ascount_no')
            .eq('date_key', '2025-12-31') 
            .eq('campaign_type', currentActivePage);

        if (error) throw error;

        yearEnd2025Data = {};
        const unique2025 = new Map();
        data.forEach(item => unique2025.set(item.ascount_no, item));
        unique2025.forEach(item => {
            if (!yearEnd2025Data[item.branch_name]) yearEnd2025Data[item.branch_name] = 0;
            yearEnd2025Data[item.branch_name] += (item.closing_balance || 0);
        });
    } catch (err) { console.error(err); }
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
        <th align="right">ຍອດຍົກມາ</th>
        <th align="right">ແຜນປີ 2026</th>
        <th align="right" style="color: #fbbf24;">ຍອດເງິນເປີດບັນຊີໃໝ່</th>
        <th align="right" style="color: #3b82f6;">ປະຕິບັດ </th>
        <th align="right">ທຽບປີ2025</th>
        <th align="right">ທຽບແຜນປີ2026</th>
        <th align="right">ທຽບຍອດເງິນເປີດໃໝ່</th>
        <th align="center">%ທຽບປີ2025</th>
        <th align="center">%ທຽບແຜນປີ2026</th>
    `;
}

// ==========================================
// 4. Filtering & Rendering
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

function renderMainTable(tableData) {
    const tbody = document.querySelector('#branch-table tbody');
    if (!tbody) return;
    
    let allBranches = [...new Set([ ...rawData.map(d => d.branch_name), ...dynamicBranchSettings.map(s => s.branch_name) ])];
    allBranches.sort();

    let totalBaseline=0, totalPlan26=0, totalOpening=0, totalClosing=0; 
    let totalDiff25=0, totalDiffPlan=0, totalDiffOpening=0;
    
    const rowsHtml = allBranches.map(bName => {
        const set2025 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2025);
        const set2026 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2026);
        const plan2026 = set2026 ? set2026.target_amount : 0; 
        const bData = tableData.filter(d => d.branch_name === bName);
        const accountGroups = {};
        bData.forEach(d => { if(!accountGroups[d.ascount_no]) accountGroups[d.ascount_no] = []; accountGroups[d.ascount_no].push(d); });

        let sumOpeningCol4 = 0, sumClosingCol5 = 0; 
        Object.values(accountGroups).forEach(recs => {
            recs.sort((a, b) => (a.date_key > b.date_key) ? 1 : -1);
            sumOpeningCol4 += (Number(recs[0].opening_banlance) || 0);
            sumClosingCol5 += (Number(recs[recs.length - 1].closing_balance) || 0);
        });

        let baselineVal = 0;
        if (currentActivePage === 'PUPOM') {
            baselineVal = yearEnd2025Data[bName] || 0;
            if (baselineVal === 0 && set2025) baselineVal = set2025.balance_prev_year || 0;
        } else {
            const allHistoryData = rawData.filter(d => d.branch_name === bName);
            const historyAccounts = {};
            allHistoryData.forEach(d => { if(!historyAccounts[d.ascount_no]) historyAccounts[d.ascount_no] = []; historyAccounts[d.ascount_no].push(d); });
            Object.values(historyAccounts).forEach(recs => { recs.sort((a, b) => (a.date_key > b.date_key) ? 1 : -1); baselineVal += (Number(recs[0].closing_balance) || 0); });
        }

        const diff25 = sumClosingCol5 - baselineVal;
        const diffPlan = sumClosingCol5 - plan2026;
        const diffOpening = sumClosingCol5 - sumOpeningCol4;
        let pct25 = baselineVal !== 0 ? ((sumClosingCol5 - baselineVal) / baselineVal) * 100 : (sumClosingCol5 > 0 ? 100 : 0);
        let pctPlan = plan2026 !== 0 ? (sumClosingCol5 / plan2026) * 100 : 0;

        totalBaseline += baselineVal; totalPlan26 += plan2026; totalOpening += sumOpeningCol4;
        totalClosing += sumClosingCol5; totalDiff25 += diff25; totalDiffPlan += diffPlan; totalDiffOpening += diffOpening;

        const formatTrend = (v, p) => `<div style="display:flex; justify-content:flex-end; gap:6px; color:${v>0?'#10b981':'#ef4444'}"><span>${v>0?'▲':'▼'}</span><span>${Math.abs(v).toLocaleString()}${p?'%':''}</span></div>`;
        const formatPct = (v) => `<div style="font-weight:bold; color:${v>=100?'#10b981':(v>=80?'#fbbf24':'#ef4444')}">${v.toFixed(1)}%</div>`;

        return `<tr>
            <td style="font-size:0.85rem; padding: 12px;">${bName}</td>
            <td align="right" style="font-weight:bold; color:#e2e8f0;">${baselineVal.toLocaleString()}</td>
            <td align="right" style="color:#94a3b8;">${plan2026.toLocaleString()}</td>
            <td align="right" style="color:#fbbf24; font-weight:bold;">${sumOpeningCol4.toLocaleString()}</td>
            <td align="right" style="color:#3b82f6; font-weight:bold;">${sumClosingCol5.toLocaleString()}</td>
            <td align="right">${formatTrend(diff25)}</td>
            <td align="right">${formatTrend(diffPlan)}</td>
            <td align="right">${formatTrend(diffOpening)}</td>
            <td align="center">${formatTrend(pct25, true)}</td>
            <td align="center">${formatPct(pctPlan)}</td>
        </tr>`;
    }).join('');

    let totalPctPlan = totalPlan26 !== 0 ? (totalClosing / totalPlan26) * 100 : 0;
    let totalPct25 = totalBaseline !== 0 ? ((totalClosing - totalBaseline) / totalBaseline) * 100 : 0;

    const formatTotalTrend = (v, p) => `<div style="display:flex; justify-content:flex-end; gap:6px; color:${v>0?'#4ade80':'#f87171'}"><span>${v>0?'▲':'▼'}</span><span>${Math.abs(v).toLocaleString()}${p?'%':''}</span></div>`;

    const totalHtml = `<tr style="background-color: #0f172a; border-top: 2px solid #334155; font-weight: bold;">
        <td style="padding: 15px; color: #ffffff;">ລວມທັງໝົດ</td>
        <td align="right" style="color: #ffffff;">${totalBaseline.toLocaleString()}</td>
        <td align="right" style="color: #ffffff;">${totalPlan26.toLocaleString()}</td>
        <td align="right" style="color: #fbbf24;">${totalOpening.toLocaleString()}</td>
        <td align="right" style="color: #3b82f6;">${totalClosing.toLocaleString()}</td>
        <td align="right">${formatTotalTrend(totalDiff25)}</td>
        <td align="right">${formatTotalTrend(totalDiffPlan)}</td>
        <td align="right">${formatTotalTrend(totalDiffOpening)}</td>
        <td align="center">${formatTotalTrend(totalPct25, true)}</td>
        <td align="center" style="color:${totalPctPlan>=100?'#4ade80':'#fbbf24'}">${totalPctPlan.toFixed(1)}%</td>
    </tr>`;
    tbody.innerHTML = rowsHtml + totalHtml;
}

function renderCharts(uniqueData) {
    if (!charts.area || !charts.donut) return;
    const branches = [...new Set(uniqueData.map(d => d.branch_name))];
    const totals = branches.map(b => uniqueData.filter(d => d.branch_name === b).reduce((s, c) => s + (Number(c.closing_balance) || 0), 0));

    charts.area.setOption({
        tooltip: { trigger: 'axis' },
        grid: { left: '3%', right: '4%', bottom: '3%', containLabel: true },
        xAxis: { type: 'category', data: branches, axisLabel: { rotate: 45, color: '#94a3b8' } },
        yAxis: { type: 'value', axisLabel: { color: '#94a3b8' } },
        series: [{ data: totals, type: 'bar', itemStyle: { color: '#3b82f6' }, name: 'ຍອດເງິນປັດຈຸບັນ' }]
    });

    let a=0, i=0, z=0; uniqueData.forEach(d => { const o=Number(d.opening_banlance)||0, c=Number(d.closing_balance)||0; if(c===0) z++; else if(o!==c) a++; else i++; });
    charts.donut.setOption({
        title: { text: (a+i+z).toLocaleString(), left: 'center', top: '42%', textStyle: { color: '#ffffff', fontSize: 24 } },
        tooltip: { trigger: 'item' },
        legend: { bottom: '0%', textStyle: { color: '#fff' } },
        series: [{ type: 'pie', radius: ['55%', '75%'], data: [{ value: a, name: 'ເຄື່ອນໄຫວ', itemStyle: { color: '#10b981' } }, { value: i, name: 'ບໍ່ເຄື່ອນໄຫວ', itemStyle: { color: '#f59e0b' } }, { value: z, name: 'ບັນຊີສູນ', itemStyle: { color: '#ef4444' } }] }]
    });
}
// ==========================================
// 4.1 Table Rendering (✅ Fix: ຄິດໄລ່ຍອດເປີດຈາກວັນທຳອິດ ແລະ ຍອດປິດຈາກວັນລ່າສຸດ)
// ==========================================
function renderMainTable(tableData) {
    const tbody = document.querySelector('#branch-table tbody');
    if (!tbody) return;
    
    let allBranches = [...new Set([
        ...rawData.map(d => d.branch_name), 
        ...dynamicBranchSettings.map(s => s.branch_name)
    ])];

    const customOrder = [
        "ສາຂາພາກນະຄອນຫຼວງ",
        "ສາຂານ້ອຍໂພນໄຊ",
        "ສາຂານ້ອຍດອນໜູນ",
        "ສາຂານ້ອຍຊັງຈ້ຽງ",
        "ສາຂານ້ອຍໜອງໜ່ຽງ"
    ];

    allBranches.sort((a, b) => {
        let indexA = customOrder.indexOf(a);
        let indexB = customOrder.indexOf(b);
        if (indexA === -1) indexA = 999;
        if (indexB === -1) indexB = 999;
        if (indexA !== indexB) return indexA - indexB;
        return a.localeCompare(b, 'lo');
    });

    // Grand Total Vars
    let totalBaseline = 0; 
    let totalPlan26 = 0;
    let totalOpening = 0; 
    let totalClosing = 0; 
    let totalDiff25 = 0;
    let totalDiffPlan = 0;
    let totalDiffOpening = 0;
    
    const rowsHtml = allBranches.map(bName => {
        const set2025 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2025);
        const set2026 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2026);
        const plan2026 = set2026 ? set2026.target_amount : 0; 
        
        // ກັ່ນຕອງຂໍ້ມູນຕາມສາຂາ
        const bData = tableData.filter(d => d.branch_name === bName);
        
        // ✅ Group ຕາມບັນຊີ ເພື່ອຫາ "First Record" ແລະ "Last Record"
        const accountGroups = {};
        bData.forEach(d => {
            if(!accountGroups[d.ascount_no]) accountGroups[d.ascount_no] = [];
            accountGroups[d.ascount_no].push(d);
        });

        let sumOpeningCol4 = 0; // ສຳລັບຄໍລຳ 4
        let sumClosingCol5 = 0; // ສຳລັບຄໍລຳ 5

        Object.values(accountGroups).forEach(recs => {
            // ລຽງວັນທີ ນ້ອຍ -> ໃຫຍ່
            recs.sort((a, b) => (a.date_key > b.date_key) ? 1 : -1);
            
            // ຄໍລຳ 4: ເອົາ Opening Balance ຂອງວັນທີທຳອິດທີ່ພົບ
            sumOpeningCol4 += (Number(recs[0].opening_banlance) || 0);
            
            // ຄໍລຳ 5: ເອົາ Closing Balance ຂອງວັນທີລ່າສຸດທີ່ພົບ
            sumClosingCol5 += (Number(recs[recs.length - 1].closing_balance) || 0);
        });

        // 3. ຍອດປີ 2025 / ຍອດຍົກມາ (Col 2)
        let baselineVal = 0;
        if (currentActivePage === 'PUPOM') {
            baselineVal = yearEnd2025Data[bName] || 0;
            if (baselineVal === 0 && set2025) baselineVal = set2025.balance_prev_year || 0;
        } else {
            // Beerlao/Campaign: ດຶງຍອດປິດຄັ້ງທຳອິດສຸດຈາກ rawData (Absolute History)
            const allHistoryData = rawData.filter(d => d.branch_name === bName);
            const historyAccounts = {};
            allHistoryData.forEach(d => {
                if(!historyAccounts[d.ascount_no]) historyAccounts[d.ascount_no] = [];
                historyAccounts[d.ascount_no].push(d);
            });
            Object.values(historyAccounts).forEach(recs => {
                recs.sort((a, b) => (a.date_key > b.date_key) ? 1 : -1);
                baselineVal += (Number(recs[0].closing_balance) || 0);
            });
        }

        // --- ຄິດໄລ່ສ່ວນຕ່າງ ---
        const diff25 = sumClosingCol5 - baselineVal;       // Col 6
        const diffPlan = sumClosingCol5 - plan2026;        // Col 7
        const diffOpening = sumClosingCol5 - sumOpeningCol4; // Col 8 (ປັດຈຸບັນ - ເປີດໃໝ່)

        // --- ຄິດໄລ່ % ---
        let pct25 = 0;
        if (baselineVal !== 0) {
            pct25 = ((sumClosingCol5 - baselineVal) / baselineVal) * 100;
        } else if (sumClosingCol5 > 0) pct25 = 100;

        let pctPlan = 0;
        if (plan2026 !== 0) {
            pctPlan = (sumClosingCol5 / plan2026) * 100;
        }

        // Accumulate Grand Total
        totalBaseline += baselineVal;
        totalPlan26 += plan2026;
        totalOpening += sumOpeningCol4;
        totalClosing += sumClosingCol5;
        totalDiff25 += diff25;
        totalDiffPlan += diffPlan;
        totalDiffOpening += diffOpening;

        // Formatter
        const formatTrend = (value, isPercent = false) => {
            if (Math.abs(value) < 0.01) return `<span style="color: #64748b;">-</span>`;
            const isPositive = value > 0;
            const color = isPositive ? '#10b981' : '#ef4444';
            const icon = isPositive ? '▲' : '▼';
            const displayValue = Math.abs(value).toLocaleString(undefined, { minimumFractionDigits: isPercent ? 1 : 0, maximumFractionDigits: 2 });
            const suffix = isPercent ? '%' : '';
            return `<div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px; color: ${color};"><span style="font-size: 0.8em;">${icon}</span><span style="font-weight: bold;">${displayValue}${suffix}</span></div>`;
        };

        const formatPct = (val) => {
             const color = val >= 100 ? '#10b981' : (val >= 80 ? '#fbbf24' : '#ef4444');
             return `<div style="font-weight:bold; color: ${color}">${val.toFixed(1)}%</div>`;
        };

        return `<tr>
            <td style="font-size:0.85rem; padding: 12px;">${bName}</td>
            <td align="right" style="font-weight:bold; color:#e2e8f0;">${baselineVal.toLocaleString()}</td>
            <td align="right" style="color:#94a3b8;">${plan2026.toLocaleString()}</td>
            <td align="right" style="color:#fbbf24; font-weight:bold;">${sumOpeningCol4.toLocaleString()}</td>
            <td align="right" style="color:#3b82f6; font-weight:bold;">${sumClosingCol5.toLocaleString()}</td>
            <td align="right">${formatTrend(diff25)}</td>
            <td align="right">${formatTrend(diffPlan)}</td>
            <td align="right">${formatTrend(diffOpening)}</td>
            <td align="center">${formatTrend(pct25, true)}</td>
            <td align="center">${formatPct(pctPlan)}</td>
        </tr>`;
    }).join('');

    // --- Grand Total ---
    let totalPct25 = 0;
    if (totalBaseline !== 0) totalPct25 = ((totalClosing - totalBaseline) / totalBaseline) * 100;
    
    let totalPctPlan = 0;
    if (totalPlan26 !== 0) totalPctPlan = (totalClosing / totalPlan26) * 100;

    const formatTotalTrend = (value, isPercent = false) => {
        if (Math.abs(value) < 0.01) return `<span>-</span>`;
        const isPositive = value > 0;
        const color = isPositive ? '#4ade80' : '#f87171';
        const icon = isPositive ? '▲' : '▼';
        const displayValue = Math.abs(value).toLocaleString(undefined, { minimumFractionDigits: isPercent ? 1 : 0, maximumFractionDigits: 2 });
        const suffix = isPercent ? '%' : '';
        return `<div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px; color: ${color};"><span style="font-size: 0.8em;">${icon}</span><span>${displayValue}${suffix}</span></div>`;
    };

    const totalHtml = `
        <tr style="background-color: #0f172a; border-top: 2px solid #334155; font-weight: bold;">
            <td style="padding: 15px; color: #ffffff; font-size: 1rem; text-align: left;">ລວມທັງໝົດ (Total)</td>
            <td align="right" style="color: #ffffff;">${totalBaseline.toLocaleString()}</td>
            <td align="right" style="color: #ffffff;">${totalPlan26.toLocaleString()}</td>
            <td align="right" style="color: #fbbf24;">${totalOpening.toLocaleString()}</td>
            <td align="right" style="color: #3b82f6; font-size: 1rem;">${totalClosing.toLocaleString()}</td>
            <td align="right">${formatTotalTrend(totalDiff25)}</td>
            <td align="right">${formatTotalTrend(totalDiffPlan)}</td>
            <td align="right">${formatTotalTrend(totalDiffOpening)}</td>
            <td align="center">${formatTotalTrend(totalPct25, true)}</td>
            <td align="center" style="color: ${totalPctPlan >= 100 ? '#4ade80' : '#fbbf24'};">${totalPctPlan.toFixed(1)}%</td>
        </tr>
    `;

    tbody.innerHTML = rowsHtml + totalHtml;
}

// ==========================================
// 5. Excel Upload Logic (✅ Fix: ອະນາໄມຫົວຖັນ + ລຶບຈຸດໃນຕົວເລກ)
// ==========================================
async function confirmUpload() {
    const fileInput = document.getElementById('excel-input');
    const dateKey = document.getElementById('upload-date-select').value;
    
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

            const groupedData = {};

            jsonData.forEach(row => {
                // Clean headers
                const cleanRow = {};
                for (let key in row) {
                    const cleanKey = String(key).toLowerCase().replace(/[\s_]/g, '');
                    cleanRow[cleanKey] = row[key];
                }

                // Account No
                const accRaw = cleanRow['ascountno'] || cleanRow['accountno'] || cleanRow['accno'];
                const accNo = String(accRaw || '').trim();

                if(!accNo) return; 

                const branchName = cleanRow['branchname'] || cleanRow['branch'] || row['ສາຂາ'] || 'ບໍ່ລະບຸ';
                const custName = cleanRow['customername'] || cleanRow['name'] || row['ຊື່ລູກຄ້າ'] || '';

                // Numbers (Clean commas)
                const rawOpen = String(cleanRow['openingbanlance'] || cleanRow['openingbalance'] || '0').replace(/,/g, '');
                const rawClose = String(cleanRow['closingbalance'] || cleanRow['balance'] || '0').replace(/,/g, '');

                const openVal = parseFloat(rawOpen) || 0;
                const closeVal = parseFloat(rawClose) || 0;

                const uniqueKey = `${accNo}_${dateKey}_${currentActivePage}`;

                if (!groupedData[uniqueKey]) {
                    groupedData[uniqueKey] = {
                        branch_name: branchName,
                        ascount_no: accNo,
                        customer_name: custName,
                        opening_banlance: openVal, 
                        closing_balance: closeVal,
                        campaign_type: currentActivePage,
                        date_key: dateKey,
                        currency: 'LAK'
                    };
                } else {
                    groupedData[uniqueKey].opening_banlance += openVal;
                    groupedData[uniqueKey].closing_balance += closeVal;
                }
            });

            const finalPayload = Object.values(groupedData);
            if (finalPayload.length === 0) throw new Error("ບໍ່ພົບຂໍ້ມູນບັນຊີໃນໄຟລ໌!");

            const { error: dataError } = await _supabase
                .from('bank_data')
                .upsert(finalPayload, { onConflict: 'ascount_no,date_key,campaign_type' });

            if (dataError) throw dataError;

            if(dateKey === '2025-12-31') {
                await fetchYearEnd2025();
            }

            alert(`✅ ອັບໂຫລດສຳເລັດ! ອ່ານໄດ້: ${finalPayload.length} ບັນຊີ`);
            document.getElementById('upload-modal').style.display = 'none';
            loadData();

        } catch (err) {
            console.error("Upload Error:", err);
            alert("ເກີດຂໍ້ຜິດພາດ: " + err.message);
        } finally {
            btn.innerHTML = originalText;
            btn.disabled = false;
            document.getElementById('excel-input').value = ''; 
        }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

// ==========================================
// 6. Branch Management
// ==========================================

function renderBranchList() {
    const tbody = document.getElementById('goal-table-body');
    if (!tbody) return;
    
    const grouped = {};
    dynamicBranchSettings.forEach(item => {
        if (!grouped[item.branch_name]) grouped[item.branch_name] = {};
        grouped[item.branch_name][item.target_year] = item.target_amount;
    });

    tbody.innerHTML = Object.keys(grouped).map(branch => {
        const goal25 = grouped[branch][2025] || 0;
        const goal26 = grouped[branch][2026] || 0;

        return `
            <tr>
                <td style="text-align:left;">${branch}</td>
                <td align="right" style="color:#94a3b8;">${goal25.toLocaleString()}</td>
                <td align="right" style="color:#fff;">${goal26.toLocaleString()}</td>
                <td style="text-align:center;">
                    <button class="btn-edit-small" onclick="prepareEdit('${branch}')" style="color:#3b82f6; margin-right:10px; cursor:pointer; border:none; background:none; font-weight:bold;">
                        <i class="fas fa-edit"></i> ແກ້ໄຂ
                    </button>
                    <button class="btn-delete-small" onclick="deleteBranch('${branch}')" style="color:#ef4444; cursor:pointer; border:none; background:none;">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            </tr>
        `;
    }).join('');
}

// ==========================================
// 7. Chart Logic (Restore Original Design)
// ==========================================
function renderCharts(uniqueData) {
    if (!charts.area || !charts.donut) return;

    // --- 1. ຈັດກຽມຂໍ້ມູນ Bar Chart ---
    const branches = [...new Set(uniqueData.map(d => d.branch_name))];
    const totals = branches.map(b => {
        const branchData = uniqueData.filter(d => d.branch_name === b);
        return branchData.reduce((s, c) => s + (Number(c.closing_balance) || 0), 0);
    });

    // --- 2. ແຕ່ງ Bar Chart (ແບບເດີມ) ---
    charts.area.setOption({
        tooltip: { 
            trigger: 'axis',
            backgroundColor: 'rgba(30, 41, 59, 0.9)', // ພື້ນຫຼັງ Tooltip ສີເຂັ້ມ
            borderColor: '#334155',
            textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao, sans-serif' }
        },
        grid: { 
            left: '3%', right: '4%', bottom: '10%', containLabel: true 
        },
        xAxis: { 
            type: 'category', 
            data: branches, 
            axisLabel: { 
                rotate: 45, 
                color: '#94a3b8', 
                fontFamily: 'Noto Sans Lao, sans-serif',
                fontSize: 11
            },
            axisLine: { lineStyle: { color: '#334155' } }
        },
        yAxis: { 
            type: 'value', 
            axisLabel: { 
                color: '#94a3b8', 
                fontFamily: 'Noto Sans Lao, sans-serif' 
            },
            splitLine: {
                lineStyle: { color: 'rgba(255,255,255,0.05)' } // ເສັ້ນຕາຕະລາງຈາງໆ
            }
        },
        series: [{ 
            data: totals, 
            type: 'bar', 
            itemStyle: { 
                color: '#3b82f6', 
                borderRadius: [4, 4, 0, 0] // ມົນຫົວ
            },
            name: 'ຍອດເງິນປັດຈຸບັນ',
            barMaxWidth: 60
        }]
    });

    // --- 3. ຈັດກຽມຂໍ້ມູນ Donut Chart ---
    let activeCount = 0;   
    let inactiveCount = 0; 
    let zeroCount = 0;     

    uniqueData.forEach(d => {
        const open = Number(d.opening_banlance) || 0;
        const close = Number(d.closing_balance) || 0;

        if (close === 0) zeroCount++;
        else if (open !== close) activeCount++;
        else inactiveCount++;
    });

    const totalAccounts = activeCount + inactiveCount + zeroCount;
    
    const donutData = [
        { value: activeCount, name: 'ເຄື່ອນໄຫວ', itemStyle: { color: '#10b981' } },   
        { value: inactiveCount, name: 'ບໍ່ເຄື່ອນໄຫວ', itemStyle: { color: '#f59e0b' } }, 
        { value: zeroCount, name: 'ບັນຊີສູນ', itemStyle: { color: '#ef4444' } }       
    ];

    // --- 4. ແຕ່ງ Donut Chart (ແບບມີຂອບ ແລະ ຟອນງາມ) ---
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
        legend: { 
            bottom: '0%', 
            left: 'center', 
            textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao, sans-serif' },
            itemGap: 20,
            selectedMode: true 
        },
        series: [{ 
            type: 'pie', 
            radius: ['55%', '75%'], 
            avoidLabelOverlap: false,
            itemStyle: { 
                borderRadius: 10, 
                borderColor: '#1e293b', // ສີຂອບຕັດກັບພື້ນຫຼັງ (ໃຫ້ເຫັນຮອຍແຍກ)
                borderWidth: 3 
            },
            label: { show: false, position: 'center' },
            emphasis: {
                label: { show: false }, // ບໍ່ໃຫ້ໂຊ Label ທັບກັນຕອນຊີ້
                itemStyle: {
                    shadowBlur: 10,
                    shadowOffsetX: 0,
                    shadowColor: 'rgba(0, 0, 0, 0.5)'
                }
            },
            data: donutData 
        }]
    });

    // Event: Update Total on Legend Click
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
    if (!fileInput.files.length || !dateKey) return alert("ກະລຸນາເລືອກໄຟລ໌ ແລະ ວັນທີ!");

    const btn = document.getElementById('start-upload-btn');
    btn.innerHTML = "ກຳລັງປະມວນຜົນ..."; btn.disabled = true;

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            const groupedData = {};

            jsonData.forEach(row => {
                const cleanRow = {};
                for (let key in row) cleanRow[String(key).toLowerCase().replace(/[\s_]/g, '')] = row[key];
                const accNo = String(cleanRow['ascountno'] || cleanRow['accountno'] || cleanRow['accno'] || '').trim();
                if(!accNo) return; 
                const branch = cleanRow['branchname']||cleanRow['branch']||row['ສາຂາ']||'ບໍ່ລະບຸ';
                const openVal = parseFloat(String(cleanRow['openingbanlance']||'0').replace(/,/g,''))||0;
                const closeVal = parseFloat(String(cleanRow['closingbalance']||'0').replace(/,/g,''))||0;
                const uKey = `${accNo}_${dateKey}_${currentActivePage}`;
                if (!groupedData[uKey]) groupedData[uKey] = { branch_name: branch, ascount_no: accNo, opening_banlance: openVal, closing_balance: closeVal, campaign_type: currentActivePage, date_key: dateKey };
                else { groupedData[uKey].opening_banlance += openVal; groupedData[uKey].closing_balance += closeVal; }
            });

            const { error } = await _supabase.from('bank_data').upsert(Object.values(groupedData), { onConflict: 'ascount_no,date_key,campaign_type' });
            if (error) throw error;
            if(dateKey === '2025-12-31') await fetchYearEnd2025();
            alert("✅ ອັບໂຫລດສຳເລັດ!"); document.getElementById('upload-modal').style.display = 'none'; loadData();
        } catch (err) { alert("ເກີດຂໍ້ຜິດພາດ: " + err.message); } 
        finally { btn.innerHTML = "ຢືນຢັນ"; btn.disabled = false; document.getElementById('excel-input').value = ''; }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

function exportToPDF() { 
    const element = document.getElementById('table-content-area'); const today = new Date().toISOString().slice(0,10);
    html2pdf().set({ margin:[0.2,0.2,0.2,0.2], filename:`Report_${today}.pdf`, image:{type:'jpeg',quality:1}, html2canvas:{scale:3,useCORS:true,backgroundColor:'#1e293b'}, jsPDF:{unit:'in',format:'a3',orientation:'landscape'} }).from(element).save();
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

function renderBranchList() {
    const tbody = document.getElementById('goal-table-body'); if (!tbody) return;
    const grouped = {};
    dynamicBranchSettings.forEach(item => { if (!grouped[item.branch_name]) grouped[item.branch_name] = {}; grouped[item.branch_name][item.target_year] = item.target_amount; });
    tbody.innerHTML = Object.keys(grouped).map(branch => {
        return `<tr><td style="text-align:left;">${branch}</td><td align="right" style="color:#94a3b8;">${(grouped[branch][2025]||0).toLocaleString()}</td><td align="right" style="color:#fff;">${(grouped[branch][2026]||0).toLocaleString()}</td><td style="text-align:center;"><button class="btn-edit-small" onclick="prepareEdit('${branch}')" style="color:#3b82f6; margin-right:10px; border:none; background:none;"><i class="fas fa-edit"></i></button><button class="btn-delete-small" onclick="deleteBranch('${branch}')" style="color:#ef4444; border:none; background:none;"><i class="fas fa-trash"></i></button></td></tr>`;
    }).join('');
}
function prepareEdit(b) { document.getElementById('branch-name').value=b; document.getElementById('select-year').value="2026"; updateAmountInput(); }
async function saveBranch() { /* ... */ }
async function deleteBranch(b) { /* ... */ }
async function loadBranchSettings() { const {data}=await _supabase.from('branch_settings').select('*').eq('campaign_type',currentActivePage); dynamicBranchSettings=data||[]; renderBranchList(); }
function updateAmountInput() { /* ... */ }
function closeGoalModal() { document.getElementById('goal-modal').style.display='none'; }
function updateBranchDropdown() { const s=document.getElementById('branch-select'); const b=[...new Set(rawData.map(d=>d.branch_name))]; s.innerHTML='<option value="all">ທັງໝົດ</option>'+b.map(x=>`<option value="${x}">${x}</option>`).join(''); }
function setQuickFilter(t) { /* ... */ }
async function loadKPIData() { /* ... */ }
async function deleteDataByDate() { /* ... */ }

// ==========================================
// 9. Marketing KPI Logic
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
    let color = '#ef4444'; 
    let text = "ຕ້ອງປັບປຸງ";
    if(totalScore >= 100) { color = '#10b981'; text = "ດີເລີດ (Excellent)"; }
    else if(totalScore >= 80) { color = '#3b82f6'; text = "ດີ (Good)"; }
    else if(totalScore >= 60) { color = '#f59e0b'; text = "ປານກາງ (Fair)"; }
    
    statusEl.innerText = text;
    statusEl.style.color = color;

    const tbody = document.getElementById('kpi-table-body');
    tbody.innerHTML = `
        <tr>
            <td>💰 ຍອດເງິນຝາກ</td>
            <td align="right">${Number(data.target_deposit).toLocaleString()}</td>
            <td align="right">${Number(data.actual_deposit).toLocaleString()}</td>
            <td align="center" style="color:${getColor(depPct)}">${depPct.toFixed(1)}%</td>
            <td align="center">${(depPct * 0.5).toFixed(1)} / 50</td>
        </tr>
        <tr>
            <td>ເປີດບັນຊີໃໝ່</td>
            <td align="right">${data.target_new_acc}</td>
            <td align="right">${data.actual_new_acc}</td>
            <td align="center" style="color:${getColor(accPct)}">${accPct.toFixed(1)}%</td>
            <td align="center">${(accPct * 0.3).toFixed(1)} / 30</td>
        </tr>
        <tr>
            <td>🚗 ລົງຢ້ຽມຢາມລູກຄ້າ</td>
            <td align="right">${data.target_visit}</td>
            <td align="right">${data.actual_visit}</td>
            <td align="center" style="color:${getColor(visitPct)}">${visitPct.toFixed(1)}%</td>
            <td align="center">${(visitPct * 0.2).toFixed(1)} / 20</td>
        </tr>
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
        radar: {
            indicator: [
                { name: 'ເງິນຝາກ', max: 120 },
                { name: 'ບັນຊີໃໝ່', max: 120 },
                { name: 'ຢ້ຽມຢາມ', max: 120 }
            ],
            splitArea: { areaStyle: { color: ['#1e293b', '#1e293b', '#1e293b', '#1e293b'] } },
            axisName: { color: '#fff' }
        },
        series: [{
            type: 'radar',
            data: [{
                value: dataArr,
                name: 'Performance',
                itemStyle: { color: '#3b82f6' },
                areaStyle: { opacity: 0.3 }
            }]
        }]
    };
    kpiChart.setOption(option);
}

// ==========================================
// 10. Delete & Export Logic
// ==========================================
async function deleteDataByDate() {
    const dateKey = document.getElementById('upload-date-select').value;
    if (!dateKey) return alert("ກະລຸນາເລືອກ 'ວັນທີຂໍ້ມູນ' ທີ່ຕ້ອງການລົບກ່ອນ!");

    if (!confirm(`⚠️ ຢືນຢັນລົບຂໍ້ມູນວັນທີ: ${dateKey} ຂອງ ${currentActivePage}?`)) return;

    const { error } = await _supabase.from('bank_data').delete().eq('date_key', dateKey).eq('campaign_type', currentActivePage);
    if (error) alert("Error: " + error.message);
    else { alert("ລົບຂໍ້ມູນສຳເລັດ!"); document.getElementById('upload-modal').style.display = 'none'; loadData(); }
}

function exportToPDF() {
    const element = document.getElementById('table-content-area'); 
    const today = new Date().toISOString().slice(0,10);

    let pdfHeader = "ລາຍງານທົ່ວໄປ";
    if (currentActivePage === 'PUPOM') pdfHeader = "ລາຍງານ ປູພົມ";
    else if (currentActivePage === 'BEERLAO') pdfHeader = "ລາຍງານ ເບຍລາວ";
    else if (currentActivePage === 'CAMPAIGN') pdfHeader = "ລາຍງານ ລູກຄ້າແຄມແປນ";

    const opt = {
        margin:       [0.2, 0.2, 0.2, 0.2], 
        filename:     `${pdfHeader}_${today}.pdf`,
        image:        { type: 'jpeg', quality: 1.0 }, 
        html2canvas:  { 
            scale: 3, 
            useCORS: true, 
            scrollY: 0,
            backgroundColor: '#1e293b', 
            windowWidth: element.scrollWidth + 50 
        },
        jsPDF: { 
            unit: 'in', 
            format: 'a3', 
            orientation: 'landscape' 
        }
    };

    const container = document.createElement('div');
    container.style.width = 'fit-content';
    container.style.backgroundColor = '#1e293b';
    container.style.padding = '30px';
    container.style.color = '#ffffff';
    container.style.fontFamily = '"Noto Sans Lao", sans-serif';

    const titleEl = document.createElement('h1');
    titleEl.innerText = pdfHeader + ` (ວັນທີ: ${today})`;
    titleEl.style.textAlign = 'center';
    titleEl.style.marginBottom = '20px';
    titleEl.style.fontSize = '24px';
    titleEl.style.color = '#3b82f6'; 
    container.appendChild(titleEl);

    const tableClone = element.cloneNode(true);
    tableClone.querySelectorAll('th').forEach(th => {
        th.style.color = '#ffffff'; 
        th.style.fontSize = '12px';
        th.style.backgroundColor = '#0f172a'; 
    });
    tableClone.querySelectorAll('td').forEach(td => {
        td.style.fontSize = '12px';
    });

    container.appendChild(tableClone);

    html2pdf().set(opt).from(container).save().then(() => {
        container.remove(); 
    });
}

function exportToExcel() {
    const branch = document.getElementById('branch-select').value;
    const startDate = document.getElementById('start-date').value;
    const endDate = document.getElementById('end-date').value;

    let filtered = rawData;
    if (branch !== 'all') filtered = filtered.filter(d => d.branch_name === branch);
    if (startDate) filtered = filtered.filter(d => d.date_key >= startDate);
    if (endDate) filtered = filtered.filter(d => d.date_key <= endDate);

    if (filtered.length === 0) return alert("ບໍ່ມີຂໍ້ມູນສຳລັບດາວໂຫຼດ (ຕາມວັນທີທີ່ເລືອກ)!");

    // ✅ Fix: Export Unique Accounts
    const uniqueMap = new Map();
    filtered.forEach(item => {
        const existing = uniqueMap.get(item.ascount_no);
        if (!existing || item.date_key > existing.date_key) {
            uniqueMap.set(item.ascount_no, item);
        }
    });
    const distinctData = Array.from(uniqueMap.values());

    const excelData = distinctData.map((item, index) => ({
        "ລ/ດ": index + 1,
        "ສາຂາ": item.branch_name,
        "ເລກບັນຊີ": item.ascount_no,
        "ຊື່ລູກຄ້າ": item.customer_name,
        "ຍອດເງິນມື້ເປີດບັນຊີ": Number(item.opening_banlance) || 0,
        "ຍອດຄົງເຫຼືອປັດຈຸບັນ": Number(item.closing_balance) || 0,
        "ວັນທີຂໍ້ມູນ": item.date_key,
        "ປະເພດ": item.campaign_type
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Report");
    XLSX.writeFile(workbook, `${currentActivePage}_Unique_Report.xlsx`);
}