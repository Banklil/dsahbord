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

    // Menu Navigation
    document.querySelectorAll('.menu-item').forEach(btn => {
        btn.onclick = function() {
            document.querySelectorAll('.menu-item').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            const view = this.getAttribute('data-view');

            // ປ່ຽນຫົວຂໍ້ຕາຕະລາງ
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
            loadBranchSettings(); // ✅ ໂຫຼດຂໍ້ມູນສາຂາເມື່ອເປີດ Modal
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

    // Export Global Functions (ສຳຄັນສຳລັບ HTML onclick)
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
        // ລັອກປຸ່ມລົບໃນ Modal ດ້ວຍ
        lockButton('.btn-delete-small');
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
        rawData = allData;

        // ໂຫຼດ Setting ມາເພື່ອໃຊ້ຄຳນວນໃນ Dashboard
        const { data: settings } = await _supabase.from('branch_settings').select('*').eq('campaign_type', currentActivePage);
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
        const { data } = await _supabase.from('bank_data').select('branch_name, closing_balance, ascount_no').eq('date_key', '2025-12-31').eq('campaign_type', currentActivePage);
        yearEnd2025Data = {};
        const unique2025 = new Map();
        data?.forEach(item => unique2025.set(item.ascount_no, item));
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

function renderMainTable(tableData) {
    const tbody = document.querySelector('#branch-table tbody');
    if (!tbody) return;
    
    let allBranches = [...new Set([ ...rawData.map(d => d.branch_name), ...dynamicBranchSettings.map(s => s.branch_name) ])];
    allBranches.sort();

    let TB=0, TP=0, TO=0, TC=0, TD25=0, TDP=0, TDO=0;
    
    const rowsHtml = allBranches.map(bName => {
        const set2025 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2025);
        const set2026 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2026);
        const plan2026 = set2026 ? set2026.target_amount : 0; 
        const bData = tableData.filter(d => d.branch_name === bName);
        const accGroups = {};
        bData.forEach(d => { if(!accGroups[d.ascount_no]) accGroups[d.ascount_no] = []; accGroups[d.ascount_no].push(d); });

        let sumO = 0, sumC = 0; 
        Object.values(accGroups).forEach(recs => {
            recs.sort((a, b) => a.date_key.localeCompare(b.date_key));
            sumO += (Number(recs[0].opening_banlance) || 0);
            sumC += (Number(recs[recs.length - 1].closing_balance) || 0);
        });

        let base = yearEnd2025Data[bName] || (set2025?.balance_prev_year || 0);
        const d25 = sumC - base, dP = sumC - plan2026, dO = sumC - sumO;
        let p25 = base !== 0 ? (d25/base)*100 : (sumC>0?100:0), pP = plan2026 !== 0 ? (sumC/plan2026)*100 : 0;

        TB+=base; TP+=plan2026; TO+=sumO; TC+=sumC; TD25+=d25; TDP+=dP; TDO+=dO;

        const fmt = (v, p) => `<div style="display:flex; justify-content:center; gap:4px; color:${v>=0?'#10b981':'#ef4444'}"><span>${v>=0?'▲':'▼'}</span><span>${Math.abs(v).toLocaleString()}${p?'%':''}</span></div>`;

        return `<tr>
            <td style="text-align:left; padding-left:20px;">${bName}</td>
            <td align="center">${base.toLocaleString()}</td>
            <td align="center">${plan2026.toLocaleString()}</td>
            <td align="center" style="color:#fbbf24;">${sumO.toLocaleString()}</td>
            <td align="center" style="color:#3b82f6;">${sumC.toLocaleString()}</td>
            <td align="center">${fmt(d25)}</td>
            <td align="center">${fmt(dP)}</td>
            <td align="center">${fmt(dO)}</td>
            <td align="center">${fmt(p25, true)}</td>
            <td align="center" style="color:${pP>=100?'#10b981':'#fbbf24'}">${pP.toFixed(1)}%</td>
        </tr>`;
    }).join('');

    const totalHtml = `<tr style="background:#0f172a; font-weight:bold;">
        <td style="text-align:left; padding-left:20px;">ລວມທັງໝົດ</td>
        <td align="center">${TB.toLocaleString()}</td><td align="center">${TP.toLocaleString()}</td>
        <td align="center" style="color:#fbbf24;">${TO.toLocaleString()}</td>
        <td align="center" style="color:#3b82f6;">${TC.toLocaleString()}</td>
        <td align="center">${fmt(TD25)}</td>
        <td align="center">${fmt(TDP)}</td>
        <td align="center">${fmt(TDO)}</td>
        <td align="center">${(TB?((TC-TB)/TB)*100:0).toFixed(1)}%</td>
        <td align="center" style="color:${TP && (TC/TP)*100>=100?'#10b981':'#fbbf24'}">${(TP?(TC/TP)*100:0).toFixed(1)}%</td>
    </tr>`;
    tbody.innerHTML = rowsHtml + totalHtml;
    
    // ຟັງຊັນຊ່ວຍໃນການ Format Total (ເພີ່ມເຂົ້າໄປເພື່ອບໍ່ໃຫ້ Error)
    function fmt(v, p) { 
        return `<div style="display:flex; justify-content:center; gap:4px; color:${v>=0?'#4ade80':'#f87171'}"><span>${v>=0?'▲':'▼'}</span><span>${Math.abs(v).toLocaleString()}${p?'%':''}</span></div>`;
    }
}

// ==========================================
// 6. Branch Management (Logic ສົມບູນ)
// ==========================================

async function loadBranchSettings() {
    try {
        const { data, error } = await _supabase
            .from('branch_settings')
            .select('*')
            .eq('campaign_type', currentActivePage)
            .order('branch_name', { ascending: true });

        if (error) throw error;
        dynamicBranchSettings = data || [];
        renderBranchList(); // ✅ ເອີ້ນໃຫ້ມັນ Render
    } catch (err) {
        console.error("Error loading branch settings:", err);
    }
}

function renderBranchList() {
    const tbody = document.getElementById('goal-table-body');
    if (!tbody) return;
    
    // Group ຂໍ້ມູນຕາມສາຂາ
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
    
    // ເອີ້ນ Permission Check ອີກຄັ້ງ ເພື່ອລັອກປຸ່ມລົບໃນຕາຕະລາງຖ້າເປັນ viewer
    checkPermissions();
}

async function saveBranch() {
    const name = document.getElementById('branch-name').value.trim();
    const year = parseInt(document.getElementById('select-year').value);
    const amount = parseFloat(document.getElementById('input-amount').value || 0);

    if (!name) return alert("ກະລຸນາປ້ອນຊື່ສາຂາ!");

    // ປ້ອງກັນການກົດຊ້ຳ
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
        loadData(); // ອັບເດດ Dashboard ຫຼັງຈາກບັນທຶກ
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
    const found = dynamicBranchSettings.find(i => 
        i.branch_name === branchName && i.target_year === year
    );
    
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

    const branches = [...new Set(uniqueData.map(d => d.branch_name))];
    const totals = branches.map(b => {
        const branchData = uniqueData.filter(d => d.branch_name === b);
        return branchData.reduce((s, c) => s + (Number(c.closing_balance) || 0), 0);
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
            itemStyle: { color: '#3b82f6', borderRadius: [4, 4, 0, 0] },
            name: 'ຍອດເງິນປັດຈຸບັນ',
            barMaxWidth: 60
        }]
    });

    let activeCount = 0, inactiveCount = 0, zeroCount = 0;
    uniqueData.forEach(d => {
        const open = Number(d.opening_banlance) || 0;
        const close = Number(d.closing_balance) || 0;
        if (close === 0) zeroCount++; else if (open !== close) activeCount++; else inactiveCount++;
    });

    const totalAccounts = activeCount + inactiveCount + zeroCount;
    const donutData = [
        { value: activeCount, name: 'ເຄື່ອນໄຫວ', itemStyle: { color: '#10b981' } },   
        { value: inactiveCount, name: 'ບໍ່ເຄື່ອນໄຫວ', itemStyle: { color: '#f59e0b' } }, 
        { value: zeroCount, name: 'ບັນຊີສູນ', itemStyle: { color: '#ef4444' } }       
    ];

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
            if (params.selected[item.name]) newTotal += item.value;
        });
        charts.donut.setOption({ title: { text: newTotal.toLocaleString() } });
    });
}

// ==========================================
// 8. Upload & Export (ສະບັບເລືອກປະເພດໄດ້)
// ==========================================
async function confirmUpload() {
    const fileInput = document.getElementById('excel-input');
    const dateKey = document.getElementById('upload-date-select').value;
    
    // ✅ ດຶງຄ່າຈາກ Radio Button ວ່າເລືອກອັນໃດ
    const uploadType = document.querySelector('input[name="uploadType"]:checked').value; // 'UPDATE' ຫຼື 'NEW'

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

                // 1. ຫາເລກບັນຊີ
                const accNo = String(
                    cleanRow['accountno'] || cleanRow['ascountno'] || cleanRow['accno'] || 
                    cleanRow['account'] || cleanRow['number'] || row['ເລກບັນຊີ'] || ''
                ).trim();

                if (!accNo) return;

                // 2. ຫາຍອດເງິນ (Amount)
                let rawAmount = String(
                    cleanRow['balance'] || cleanRow['closingbalance'] || cleanRow['openingbalance'] ||
                    cleanRow['amount'] || cleanRow['total'] || 
                    cleanRow['ຍອດເງິນ'] || cleanRow['ຈຳນວນເງິນ'] || cleanRow['ຍອດຄົງເຫຼືອ'] || '0'
                ).replace(/,/g, '');
                
                const amountVal = parseFloat(rawAmount) || 0;

                // 3. ຫາຊື່ສາຂາ ແລະ ລູກຄ້າ
                const branchName = cleanRow['branchname'] || cleanRow['branch'] || row['ສາຂາ'] || 'ບໍ່ລະບຸ';
                const custName = cleanRow['customername'] || cleanRow['name'] || row['ຊື່ລູກຄ້າ'] || '';

                const uKey = `${accNo}_${dateKey}_${currentActivePage}`;

                // ✅ LOGIC ການແຍກປະເພດ (ສຳຄັນ)
                let openVal = 0;
                let closeVal = 0;

                if (uploadType === 'NEW') {
                    // ຖ້າເລືອກ "ເພີ່ມບັນຊີໃໝ່" -> ໃຫ້ຍອດເງິນເຂົ້າ Opening Balance
                    // ແລະໃຫ້ Closing Balance ເທົ່າກັນ (ເພາະຫາກໍ່ເປີດ)
                    openVal = amountVal;
                    closeVal = amountVal; 
                } else {
                    // ຖ້າເລືອກ "ອັບເດດຍອດ" -> ໃຫ້ຍອດເງິນເຂົ້າ Closing Balance ຢ່າງດຽວ
                    openVal = 0; // ຫຼື ບໍ່ປ່ຽນຄ່າເກົ່າ (Supabase Upsert ຈະຈັດການ)
                    closeVal = amountVal;
                }

                if (!groupedData[uKey]) {
                    groupedData[uKey] = {
                        branch_name: branchName,
                        ascount_no: accNo,
                        customer_name: custName,
                        campaign_type: currentActivePage,
                        date_key: dateKey,
                        // ໃສ່ຄ່າຕາມທີ່ເລືອກ
                        opening_banlance: openVal, 
                        closing_balance: closeVal 
                    };
                } else {
                    // ກໍລະນີມີຫຼາຍແຖວ ໃຫ້ບວກຍອດເຂົ້າໄປ
                    if (uploadType === 'NEW') {
                        groupedData[uKey].opening_banlance += openVal;
                        groupedData[uKey].closing_balance += closeVal;
                    } else {
                        groupedData[uKey].closing_balance = closeVal; // Closing ຄວນເອົາຄ່າລ່າສຸດ (ຫຼື += ຖ້າເປັນ Transaction)
                    }
                }
            });

            const finalPayload = Object.values(groupedData);
            
            // ສົ່ງຂໍ້ມູນໄປ Supabase
            // ໝາຍເຫດ: ຖ້າເປັນ NEW ມັນຈະທັບ Opening, ຖ້າເປັນ UPDATE ມັນຈະທັບ Closing (ໂດຍທີ່ Opening ເກົ່າບໍ່ຫາຍ ຖ້າມີຂໍ້ມູນຢູ່ແລ້ວ)
            const { error: dataError } = await _supabase
                .from('bank_data')
                .upsert(finalPayload, { onConflict: 'ascount_no,date_key,campaign_type' });

            if (dataError) throw dataError;

            if(dateKey === '2025-12-31') await fetchYearEnd2025();

            alert(`✅ ສຳເລັດ! (${uploadType === 'NEW' ? 'ເພີ່ມບັນຊີໃໝ່' : 'ອັບເດດຍອດ'}) ຈຳນວນ ${finalPayload.length} ລາຍການ`);
            document.getElementById('upload-modal').style.display = 'none';
            loadData();

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
// 10. Utils (Export, Delete)
// ==========================================
async function deleteDataByDate() {
    const dateKey = document.getElementById('upload-date-select').value;
    if (!dateKey) return alert("ກະລຸນາເລືອກ 'ວັນທີຂໍ້ມູນ' ທີ່ຕ້ອງການລົບກ່ອນ!");
    if (!confirm(`⚠️ ຢືນຢັນລົບຂໍ້ມູນວັນທີ: ${dateKey}?`)) return;
    const { error } = await _supabase.from('bank_data').delete().eq('date_key', dateKey).eq('campaign_type', currentActivePage);
    if (error) alert("Error: " + error.message);
    else { alert("ລົບຂໍ້ມູນສຳເລັດ!"); document.getElementById('upload-modal').style.display = 'none'; loadData(); }
}

function exportToPDF() {
    const element = document.getElementById('table-content-area'); 
    const today = new Date().toISOString().slice(0,10);
    const opt = { margin: [0.2, 0.2, 0.2, 0.2], filename: `Report_${today}.pdf`, image: { type: 'jpeg', quality: 1.0 }, html2canvas: { scale: 3, useCORS: true, backgroundColor: '#1e293b' }, jsPDF: { unit: 'in', format: 'a3', orientation: 'landscape' } };
    html2pdf().set(opt).from(element).save();
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