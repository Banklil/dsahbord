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
let charts = { area: null, donut: null, bar: null };
let kpiChart = null;

let currentUserRole = 'viewer'; 

// ສ້າງລາຍຊື່ສາຂາຕາມລຳດັບທີ່ຕ້ອງການ
const CUSTOM_BRANCH_ORDER = [
    "ສາຂາພາກນະຄອນຫຼວງ", 
    "ສາຂານ້ອຍໂພນໄຊ", 
    "ສາຂານ້ອຍດອນໜູນ", 
    "ສາຂານ້ອຍຊັງຈ້ຽງ", 
    "ສາຂານ້ອຍໜອງໜ່ຽງ"
];

// ຟັງຊັນສຳລັບລຽງລຳດັບສາຂາ
function sortBranchesCustom(branchesArray) {
    return branchesArray.sort((a, b) => {
        let indexA = CUSTOM_BRANCH_ORDER.indexOf(a);
        let indexB = CUSTOM_BRANCH_ORDER.indexOf(b);
        if (indexA === -1) indexA = 999;
        if (indexB === -1) indexB = 999;
        return indexA !== indexB ? indexA - indexB : a.localeCompare(b, 'lo');
    });
}

// ==========================================
// 2. Auth & Initialization
// ==========================================
document.getElementById('login-btn').addEventListener('click', () => {
    const user = document.getElementById('user-auth').value.trim();
    const pass = document.getElementById('pass-auth').value.trim();

    if (user === 'admin' && pass === '1234') { currentUserRole = 'admin'; performLogin(); } 
    else if (user === 'MKT01' && pass === '3578') { currentUserRole = 'viewer'; performLogin(); } 
    else {
        const err = document.getElementById('auth-error');
        err.style.display = 'block';
        err.innerHTML = "⚠️ ຊື່ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ!";
    }
});

function performLogin() {
    document.getElementById('login-screen').style.opacity = '0';
    setTimeout(() => {
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('dashboard-app').style.display = 'flex';
        initApp();
    }, 500);
}

function showLoading(text = 'ກຳລັງປະມວນຜົນ...') {
    const loader = document.getElementById('loading-overlay');
    if(loader) { loader.style.display = 'flex'; void loader.offsetWidth; loader.style.opacity = '1'; }
}

function hideLoading() {
    const loader = document.getElementById('loading-overlay');
    if(loader) { loader.style.opacity = '0'; setTimeout(() => { loader.style.display = 'none'; }, 300); }
}

function initApp() {
    charts.area = echarts.init(document.getElementById('largeAreaChart'));
    charts.donut = echarts.init(document.getElementById('donutChartDiv'));
    charts.bar = echarts.init(document.getElementById('barChartDiv'));

    document.querySelectorAll('.menu-item[data-view]').forEach(btn => {
        btn.onclick = function() {
            document.querySelectorAll('.menu-item').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            const view = this.getAttribute('data-view');

            const titleElement = document.getElementById('page-title');
            if (titleElement) titleElement.innerText = "Dashboard " + (view === 'PUPOM' ? 'ປູພົມ' : view === 'BEERLAO' ? 'ເບຍລາວ' : 'ລູກຄ້າແຄມແປນ');

            if (view === 'KPI') { switchPageView('kpi'); loadKPIData(); } 
            else { switchPageView('dashboard'); currentActivePage = view.toUpperCase(); loadData(); }
        };
    });

    document.getElementById('upload-trigger').onclick = () => document.getElementById('upload-modal').style.display = 'flex';
    document.getElementById('close-modal').onclick = () => document.getElementById('upload-modal').style.display = 'none';
    
    const manageBtn = document.getElementById('manage-branches-btn');
    if(manageBtn) manageBtn.onclick = () => { document.getElementById('goal-modal').style.display = 'flex'; loadBranchSettings(); };
    
    const fileInput = document.getElementById('excel-input'); 
    if (fileInput) {
        fileInput.onchange = function() {
            const display = document.getElementById('file-name-display');
            if (this.files[0] && display) display.innerHTML = `<span class="text-success">📄 ໄຟລ໌:</span> <span class="text-main">${this.files[0].name}</span>`;
        };
    }

    document.getElementById('start-upload-btn').onclick = confirmUpload;
    document.getElementById('apply-filter').onclick = handleApplyFilterClick;
    document.getElementById('logout-btn').onclick = () => location.reload();

    window.saveBranch = saveBranch; 
    window.closeGoalModal = closeGoalModal;
    window.deleteBranch = deleteBranch; 
    window.setQuickFilter = setQuickFilter;
    window.updateAmountInput = updateAmountInput; 
    window.clearFilters = clearFilters;
    window.loadKPIData = loadKPIData; 
    window.deleteDataByDate = deleteDataByDate;
    window.exportToPDF = exportToPDF; 
    window.exportToExcel = exportToExcel;
    window.resetChart = resetChart;
    
    window.switchPage = switchPage;
    window.deleteUpload = deleteUpload;
    window.replaceFile = replaceFile;
    window.handleReplaceFile = handleReplaceFile;
    window.handleNewUpload = handleNewUpload;

    checkPermissions();
    loadData();
    
    window.onresize = () => { 
        charts.area?.resize(); charts.donut?.resize(); charts.bar?.resize(); kpiChart?.resize();
    };
}

function switchPageView(page) {
    document.getElementById('view-dashboard').style.display = page === 'dashboard' ? 'block' : 'none';
    document.getElementById('view-kpi').style.display = page === 'kpi' ? 'block' : 'none';
    if(page === 'dashboard') setTimeout(() => { charts.area?.resize(); charts.donut?.resize(); charts.bar?.resize(); }, 100);
    if(page === 'kpi') setTimeout(() => { kpiChart?.resize(); }, 100);
}

function switchPage(event, pageId) {
    event.preventDefault();
    document.querySelectorAll('.menu-item').forEach(b => b.classList.remove('active'));
    event.currentTarget.classList.add('active');
    switchPageView(pageId);
}

function checkPermissions() {
    if (currentUserRole === 'viewer') {
        const lockButton = (selector) => { document.querySelectorAll(selector).forEach(el => { el.style.opacity = '0.4'; el.style.pointerEvents = 'none'; }); };
        lockButton('#upload-trigger'); lockButton('#manage-branches-btn'); lockButton('.btn-export.excel'); lockButton('.btn-icon-small.delete');
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
    showLoading('ກຳລັງໂຫລດຂໍ້ມູນຈາກຖານຂໍ້ມູນ...');
    updateTableHeader();
    await fetchYearEnd2025();
    let allData = [], from = 0, to = 999, hasMore = true;

    try {
        while (hasMore) {
            const { data, error } = await _supabase.from('bank_data').select('*').eq('campaign_type', currentActivePage).range(from, to);
            if (error) throw error;
            allData = allData.concat(data);
            if (data.length < 1000) hasMore = false; else { from += 1000; to += 1000; }
        }
        rawData = allData.map(d => ({ ...d, branch_name: (d.branch_name || '').trim(), ascount_no: (d.ascount_no || '').trim() }));

        const { data: settings } = await _supabase.from('branch_settings').select('*').eq('campaign_type', currentActivePage);
        dynamicBranchSettings = (settings || []).map(s => ({ ...s, branch_name: (s.branch_name || '').trim() }));
        
        updateBranchDropdown(); applyFilters(); 
    } catch (err) {
        console.error("Load Error:", err);
    } finally { hideLoading(); }
}

async function fetchYearEnd2025() {
    try {
        let allRows = [], from = 0, to = 999, hasMore = true;
        while (hasMore) {
            const { data, error } = await _supabase.from('bank_data').select('branch_name, closing_balance, ascount_no').eq('date_key', '2025-12-31').eq('campaign_type', currentActivePage).range(from, to);
            if (error) throw error;
            if (data && data.length > 0) allRows = allRows.concat(data);
            if (!data || data.length < 1000) hasMore = false; else { from += 1000; to += 1000; }
        }
        yearEnd2025Data = {}; const unique2025 = new Map();
        allRows.forEach(item => unique2025.set(item.ascount_no, item));
        unique2025.forEach(item => {
            if (!yearEnd2025Data[item.branch_name]) yearEnd2025Data[item.branch_name] = 0;
            yearEnd2025Data[item.branch_name] += (Number(item.closing_balance) || 0);
        });
    } catch (err) { console.error(err); }
}

function updateTableHeader() {
    const thead = document.querySelector('#branch-table thead');
    if (!thead) return;
    thead.innerHTML = `<tr>
        <th style="color:#94a3b8;">ສາຂາ</th>
        <th style="text-align:right; color:#94a3b8;">ຍອດຍົກມາ</th>
        <th style="text-align:right; color:#94a3b8;">ແຜນປີ 2026</th>
        <th style="text-align:right; color:#f59e0b;">ຍອດເປີດໃໝ່</th>
        <th style="text-align:right; color:#2563eb;">ປະຕິບັດໄດ້</th>
        <th style="text-align:center; color:#94a3b8;">ທຽບ 2025</th>
        <th style="text-align:center; color:#94a3b8;">ທຽບແຜນ</th>
        <th style="text-align:center; color:#94a3b8;">ທຽບຍອດເປີດໃໝ່</th>
        <th style="text-align:center; color:#94a3b8;">% ທຽບ 2025</th>
        <th style="text-align:center; color:#94a3b8;">% ບັນລຸ</th>
    </tr>`;
}

// ==========================================
// 4. Filtering & Rendering
// ==========================================
function handleApplyFilterClick() {
    showLoading('ກັ່ນຕອງຂໍ້ມູນ...');
    setTimeout(() => { applyFilters(); const tbody = document.getElementById('table-body'); if(tbody) { tbody.classList.remove('fade-in-table'); void tbody.offsetWidth; tbody.classList.add('fade-in-table'); } hideLoading(); }, 300);
}

function applyFilters() {
    const branch = document.getElementById('branch-select').value;
    const startDate = document.getElementById('start-date').value;
    const endDate = document.getElementById('end-date').value;
    let filtered = [];

    if (startDate || endDate) {
        filtered = rawData.filter(d => { let pass = true; if (startDate) pass = pass && (d.date_key >= startDate); if (endDate) pass = pass && (d.date_key <= endDate); return pass; });
    } else { filtered = rawData; }

    if (branch !== 'all') filtered = filtered.filter(d => d.branch_name === branch);
    updateUI(filtered);
}

function clearFilters() { document.getElementById('branch-select').value = 'all'; document.getElementById('start-date').value = ''; document.getElementById('end-date').value = ''; applyFilters(); }

function updateUI(filteredData) {
    const uniqueMap = new Map();
    filteredData.forEach(item => { const existing = uniqueMap.get(item.ascount_no); if (!existing || item.date_key > existing.date_key) uniqueMap.set(item.ascount_no, item); });
    const distinctData = Array.from(uniqueMap.values());

    const totalClosing = distinctData.reduce((sum, i) => sum + (Number(i.closing_balance) || 0), 0);
    document.getElementById('total-amount').innerText = totalClosing.toLocaleString();

    const totalTarget = dynamicBranchSettings.reduce((sum, i) => sum + (Number(i.target_amount) || 0), 0);
    const totalPct = totalTarget > 0 ? (totalClosing / totalTarget) * 100 : 0;
    
    document.getElementById('achievement-pct').innerText = totalPct.toFixed(2) + "%";
    document.getElementById('progress-fill').style.width = Math.min(totalPct, 100) + "%";

    renderMainTable(filteredData); 
    renderCharts(distinctData, filteredData); 
}

function renderMainTable(tableData) {
    const tbody = document.querySelector('#branch-table tbody');
    if (!tbody) return;
    
    let allBranches = [...new Set([ ...rawData.map(d => d.branch_name), ...dynamicBranchSettings.map(s => s.branch_name) ])];
    sortBranchesCustom(allBranches);

    let TB=0, TP=0, TO=0, TC=0, TD25=0, TDP=0, TDO=0;
    
    const fmt = (v, isPercent = false, isPlan = false) => {
        if (isPlan) {
            const color = v >= 100 ? '#10b981' : '#f59e0b';
            return `<div style="color:${color}; font-weight:bold;">${v >= 100 ? '▲' : '▼'} ${v.toFixed(1)}%</div>`;
        }
        const color = v >= 0 ? '#10b981' : '#ef4444';
        return `<div style="color:${color}; font-weight:500;">${v >= 0 ? '▲' : '▼'} ${v.toLocaleString(undefined, { minimumFractionDigits: isPercent?1:0, maximumFractionDigits: isPercent?1:2 })}${isPercent?'%':''}</div>`;
    };

    if(allBranches.length === 0) { tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:30px; color:#ffffff;">ບໍ່ມີຂໍ້ມູນ</td></tr>`; return; }

    const rowsHtml = allBranches.map(bName => {
        const set2025 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2025);
        const set2026 = dynamicBranchSettings.find(s => s.branch_name === bName && s.target_year === 2026);
        const plan2026 = set2026 ? set2026.target_amount : 0; 
        
        const bData = tableData.filter(d => d.branch_name === bName);
        const accMap = new Map();
        bData.forEach(d => { if (!accMap.has(d.ascount_no)) accMap.set(d.ascount_no, []); accMap.get(d.ascount_no).push(d); });

        let sumO = 0, sumC = 0; 
        accMap.forEach(records => {
            records.sort((a, b) => b.date_key.localeCompare(a.date_key));
            sumC += (Number(records[0].closing_balance) || 0);
            if (currentActivePage === 'PUPOM') {
                let accountOpen = 0;
                for (let r of records) { if (Number(r.opening_balance) > 0) { accountOpen = Number(r.opening_balance); break; } }
                sumO += accountOpen;
            } else {
                let records2026 = records.filter(r => !r.date_key.startsWith('2025'));
                if (records2026.length > 0) {
                    records2026.sort((a, b) => a.date_key.localeCompare(b.date_key));
                    sumO += Math.max(Number(records2026[0].opening_balance) || 0, Number(records2026[0].closing_balance) || 0);
                }
            }
        });

        let baselineVal = yearEnd2025Data[bName] || 0;
        if (baselineVal === 0 && set2025) baselineVal = set2025.balance_prev_year || 0;

        const d25 = sumC - baselineVal; 
        const dP = sumC - plan2026; 
        let dO = sumC - sumO; 
        
        let p25 = baselineVal !== 0 ? (d25 / baselineVal) * 100 : (sumC > 0 ? 100 : 0);
        let pP = plan2026 !== 0 ? (sumC / plan2026) * 100 : 0;

        TB+=baselineVal; TP+=plan2026; TO+=sumO; TC+=sumC; TD25+=d25; TDP+=dP; TDO+=dO;

        return `<tr style="color:#ffffff;">
            <td style="text-align:left; font-weight:500;">${bName}</td>
            <td align="right">${baselineVal.toLocaleString()}</td>
            <td align="right" style="color:#94a3b8;">${plan2026.toLocaleString()}</td>
            <td align="right" style="color:#f59e0b;">${sumO.toLocaleString()}</td>
            <td align="right" style="color:#2563eb; font-weight:bold;">${sumC.toLocaleString()}</td>
            <td align="center">${fmt(d25)}</td>
            <td align="center">${fmt(dP)}</td>
            <td align="center">${fmt(dO)}</td>
            <td align="center">${fmt(p25, true)}</td>
            <td align="center">${fmt(pP, true, true)}</td> 
        </tr>`;
    }).join('');

    const totalHtml = `<tr style="background:#1f2937; font-weight:bold; color:#ffffff;">
        <td style="text-align:left; color:#2563eb;">ລວມທັງໝົດ</td>
        <td align="right">${TB.toLocaleString()}</td>
        <td align="right">${TP.toLocaleString()}</td>
        <td align="right" style="color:#f59e0b;">${TO.toLocaleString()}</td>
        <td align="right" style="color:#2563eb; font-size:1.05rem;">${TC.toLocaleString()}</td>
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
        const { data, error } = await _supabase.from('branch_settings').select('*').eq('campaign_type', currentActivePage).order('branch_name', { ascending: true });
        if (error) throw error;
        dynamicBranchSettings = (data || []).map(s => ({ ...s, branch_name: (s.branch_name || '').trim() }));
        
        const tbody = document.getElementById('goal-table-body');
        const theadTr = document.querySelector('#goal-modal thead tr');
        const grouped = {}; const yearsSet = new Set();
        
        dynamicBranchSettings.forEach(item => { 
            if (!grouped[item.branch_name]) grouped[item.branch_name] = {}; 
            grouped[item.branch_name][item.target_year] = item.target_amount; 
            yearsSet.add(item.target_year);
        });

        let years = Array.from(yearsSet).sort((a,b) => a - b);
        if (years.length === 0) years = [2025, 2026];

        if (theadTr) {
            let thHtml = `<th>ສາຂາ</th>`;
            years.forEach(y => thHtml += `<th style="text-align:right;">ເປົ້າ ${y}</th>`);
            thHtml += `<th style="text-align:center; width: 80px;">ຈັດການ</th>`;
            theadTr.innerHTML = thHtml;
        }

        const branches = Object.keys(grouped);
        sortBranchesCustom(branches);
        
        tbody.innerHTML = branches.length === 0 ? `<tr><td colspan="${years.length + 2}" style="text-align:center; padding:20px; color:var(--text-muted);">ບໍ່ມີຂໍ້ມູນ</td></tr>` : 
            branches.map(branch => {
                let tdHtml = `<td style="font-weight: 500;">${branch}</td>`;
                years.forEach(y => {
                    const amount = grouped[branch][y] || 0; const isCurrent = (y == new Date().getFullYear() || y == 2026);
                    tdHtml += `<td align="right" class="${isCurrent ? 'text-primary' : 'text-muted'}" style="${isCurrent ? 'font-weight:bold;' : ''}">${amount.toLocaleString()}</td>`;
                });
                tdHtml += `<td align="center"><button class="btn-icon-small delete" onclick="deleteBranch('${branch}')">ລຶບ</button></td>`;
                return `<tr>${tdHtml}</tr>`;
            }).join('');
    } catch (err) { console.error(err); }
}

async function saveBranch() {
    const name = document.getElementById('branch-name').value.trim(); const year = parseInt(document.getElementById('select-year').value); const amount = parseFloat(document.getElementById('input-amount').value || 0);
    if (!name) return alert("ກະລຸນາປ້ອນຊື່ສາຂາ!");
    try {
        const { data: existing } = await _supabase.from('branch_settings').select('id').eq('branch_name', name).eq('target_year', year).eq('campaign_type', currentActivePage).maybeSingle();
        const payload = { branch_name: name, target_year: year, target_amount: amount, campaign_type: currentActivePage };
        if (existing) await _supabase.from('branch_settings').update(payload).eq('id', existing.id); else await _supabase.from('branch_settings').insert(payload);
        document.getElementById('branch-name').value = ''; document.getElementById('input-amount').value = ''; await loadBranchSettings(); loadData(); 
    } catch (err) { alert("Error: " + err.message); }
}

async function deleteBranch(branchName) {
    if(!confirm(`ຢືນຢັນລົບສາຂາ "${branchName}"?`)) return;
    try { await _supabase.from('branch_settings').delete().eq('branch_name', branchName).eq('campaign_type', currentActivePage); await loadBranchSettings(); loadData(); } catch (err) { alert(err.message); }
}

function updateAmountInput() {
    const branchName = document.getElementById('branch-name').value; if (!branchName) return;
    const year = parseInt(document.getElementById('select-year').value);
    const found = dynamicBranchSettings.find(i => i.branch_name === branchName && i.target_year === year);
    document.getElementById('input-amount').value = found ? found.target_amount : 0;
}

function closeGoalModal() { document.getElementById('goal-modal').style.display = 'none'; }

function updateBranchDropdown() {
    const s = document.getElementById('branch-select'); 
    let b = [...new Set(rawData.map(d => d.branch_name))];
    sortBranchesCustom(b);
    s.innerHTML = '<option value="all">ທັງໝົດ (All Branches)</option>' + b.map(x => `<option value="${x}">${x}</option>`).join('');
}

function setQuickFilter(type) {
    const today = new Date(); let start = new Date();
    if (type === 'week') start.setDate(today.getDate() - 7); if (type === 'month') start.setMonth(today.getMonth(), 1); if (type === 'year') start.setFullYear(today.getFullYear(), 0, 1);
    document.getElementById('start-date').value = start.toISOString().split('T')[0]; document.getElementById('end-date').value = today.toISOString().split('T')[0]; applyFilters();
}

// ==========================================
// 7. Chart Logic (Area, Donut, Bar)
// ==========================================
function renderCharts(uniqueData, filteredData) {
    if (!charts.area || !charts.donut || !charts.bar) return;

    const sourceData = filteredData || uniqueData;
    const dateMap = new Map();
    sourceData.forEach(d => {
        if(!dateMap.has(d.date_key)) dateMap.set(d.date_key, 0);
        dateMap.set(d.date_key, dateMap.get(d.date_key) + (Number(d.closing_balance) || 0));
    });
    const sortedDates = Array.from(dateMap.keys()).sort();
    const trendValues = sortedDates.map(date => dateMap.get(date));

    charts.area.setOption({
        tooltip: { trigger: 'axis', backgroundColor: '#1f293b', borderColor: '#374151', textStyle: { color: '#f8fafc', fontFamily: 'Noto Sans Lao' } },
        grid: { left: '3%', right: '4%', bottom: '15%', top: '10%', containLabel: true },
        xAxis: { type: 'category', boundaryGap: false, data: sortedDates, axisLabel: { color: '#94a3b8' }, axisLine: { lineStyle: { color: '#374151' } } },
        yAxis: { type: 'value', min: 'dataMin', axisLabel: { color: '#94a3b8' }, splitLine: { lineStyle: { color: '#1f2937', type: 'dashed' } } },
        dataZoom: [{ type: 'inside', start: 0, end: 100 }, { start: 0, end: 100, height: 20, bottom: 0, textStyle: { color: '#94a3b8' } }],
        series: [{ 
            name: 'ຍອດເງິນລວມ', type: 'line', symbolSize: 0, sampling: 'lttb',
            itemStyle: { color: '#3b82f6' },
            areaStyle: { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: 'rgba(59, 130, 246, 0.4)' }, { offset: 1, color: 'rgba(59, 130, 246, 0.05)' }]) },
            data: trendValues 
        }]
    }, true);

    let activeCount = 0, inactiveCount = 0, zeroCount = 0;
    
    uniqueData.forEach(d => {
        const close = Number(d.closing_balance) || 0; 
        
        if (close <= 0) {
            zeroCount++;
        } else {
            const history = filteredData.filter(r => r.ascount_no === d.ascount_no);
            history.sort((a, b) => a.date_key.localeCompare(b.date_key)); 
            
            if (history.length > 1) {
                const startClose = Number(history[0].closing_balance) || 0;
                if (Math.abs(close - startClose) < 0.01) inactiveCount++;
                else activeCount++;
            } else {
                const open = Number(d.opening_balance) || 0;
                if (open > 0 && Math.abs(open - close) < 0.01) {
                    inactiveCount++; 
                } else {
                    activeCount++; 
                }
            }
        }
    });

    const totalAccounts = activeCount + inactiveCount + zeroCount;

    const donutData = [
        { value: activeCount, name: 'ເຄື່ອນໄຫວ', itemStyle: { color: '#2563eb' } },   
        { value: inactiveCount, name: 'ບໍ່ເຄື່ອນໄຫວ', itemStyle: { color: '#fbbf24' } }, 
        { value: zeroCount, name: 'ບັນຊີສູນ', itemStyle: { color: '#ef4444' } }       
    ];
    
    charts.donut.setOption({
        title: {
            text: totalAccounts.toLocaleString(),
            subtext: 'ບັນຊີທັງໝົດ',
            left: 'center', top: 'center',
            textStyle: { color: '#ffffff', fontSize: 24, fontWeight: 'bold', fontFamily: 'Noto Sans Lao' },
            subtextStyle: { color: '#94a3b8', fontSize: 12, fontFamily: 'Noto Sans Lao' }
        },
        tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)', backgroundColor: '#1f293b', borderColor: '#374151', textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao' } },
        legend: { bottom: '0', left: 'center', textStyle: { color: '#cbd5e1', fontFamily: 'Noto Sans Lao' }, itemGap: 15, icon: 'circle' },
        series: [{ type: 'pie', radius: ['55%', '80%'], avoidLabelOverlap: false, itemStyle: { borderColor: '#111827', borderWidth: 2 }, label: { show: false }, data: donutData }]
    });

    let branches = [...new Set(dynamicBranchSettings.map(s => s.branch_name))]; 
    sortBranchesCustom(branches);

    const targetValues = branches.map(b => {
        const set = dynamicBranchSettings.find(s => s.branch_name === b && s.target_year === 2026);
        return set ? set.target_amount : 0;
    });
    const actualValues = branches.map(b => {
        const branchData = uniqueData.filter(d => d.branch_name === b);
        return branchData.reduce((s, c) => s + (Number(c.closing_balance) || 0), 0);
    });

    charts.bar.setOption({
        tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, backgroundColor: '#1f293b', borderColor: '#374151', textStyle: { color: '#fff', fontFamily: 'Noto Sans Lao' } },
        legend: { data: ['ແຜນ (Target)', 'ຕົວຈິງ (Actual)'], bottom: 0, textStyle: { color: '#cbd5e1' }, icon: 'roundRect' },
        grid: { left: '3%', right: '4%', bottom: '15%', top: '10%', containLabel: true },
        xAxis: { type: 'category', data: branches, axisLabel: { color: '#94a3b8', rotate: 30, interval: 0, fontSize: 10 }, axisLine: { lineStyle: { color: '#374151' } } },
        yAxis: { type: 'value', axisLabel: { color: '#94a3b8' }, splitLine: { lineStyle: { color: '#1f2937', type: 'dashed' } } },
        series: [
            { name: 'ແຜນ (Target)', type: 'bar', data: targetValues, itemStyle: { color: '#374151', borderRadius: [4,4,0,0] } },
            { name: 'ຕົວຈິງ (Actual)', type: 'bar', data: actualValues, itemStyle: { color: '#10b981', borderRadius: [4,4,0,0] } }
        ]
    });
}

// ==========================================
// 8. Upload & Export
// ==========================================
async function confirmUpload() {
    const fileInput = document.getElementById('excel-input'); const dateKey = document.getElementById('upload-date-select').value;
    const uploadTypeElem = document.querySelector('input[name="uploadType"]:checked'); const uploadType = uploadTypeElem ? uploadTypeElem.value : 'UPDATE'; 
    if (!fileInput.files.length || !dateKey) return alert("ກະລຸນາເລືອກໄຟລ໌ ແລະ ວັນທີ!");

    showLoading('ກຳລັງປະມວນຜົນໄຟລ໌...');
    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result); const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]]; const jsonData = XLSX.utils.sheet_to_json(sheet);
            if (jsonData.length === 0) throw new Error("ບໍ່ພົບຂໍ້ມູນໃນໄຟລ໌!");
            const groupedData = {};

            jsonData.forEach(row => {
                const cleanRow = {}; for (let key in row) cleanRow[String(key).toLowerCase().replace(/[\s_\-\.]/g, '')] = row[key];
                let branchName = cleanRow['branchname'] || cleanRow['branch'] || row['ສາຂາ'] || 'ບໍ່ລະບຸ';
                if (branchName.includes('ດອນໜູນ')) branchName = 'ສາຂານ້ອຍດອນໜູນ'; else if (branchName.includes('ໂພນໄຊ')) branchName = 'ສາຂານ້ອຍໂພນໄຊ'; else if (branchName.includes('ຊັງຈ້ຽງ')) branchName = 'ສາຂານ້ອຍຊັງຈ້ຽງ'; else if (branchName.includes('ໜອງໜ່ຽງ')) branchName = 'ສາຂານ້ອຍໜອງໜ່ຽງ'; else if (branchName.includes('ນະຄອນຫຼວງ')) branchName = 'ສາຂາພາກນະຄອນຫຼວງ';
                let accNo = String(cleanRow['accountno'] || cleanRow['ascountno'] || cleanRow['accno'] || cleanRow['account'] || cleanRow['number'] || row['ເລກບັນຊີ'] || row['ເລກທີບັນຊີ'] || '').trim();
                if (!accNo) accNo = 'DUMMY_' + Math.floor(Math.random() * 10000000);
                const custName = cleanRow['customername'] || cleanRow['name'] || row['ຊື່ລູກຄ້າ'] || '';

                let openVal = 0; let closeVal = 0;
                
                const hasOpenCol = ['openingbalance', 'ຍອດເປີດ', 'ຍອດຍົກມາ', 'ຍົກມາ', 'openbalance'].some(k => k in cleanRow);
                const hasCloseCol = ['closingbalance', 'ຍອດປິດ', 'ຍອດຄົງເຫຼືອ', 'ຍອດເງິນ', 'closebalance'].some(k => k in cleanRow);

                if (hasOpenCol || hasCloseCol) { 
                    openVal = parseFloat(String(cleanRow['openingbalance'] || cleanRow['ຍອດເປີດ'] || cleanRow['ຍອດຍົກມາ'] || cleanRow['ຍົກມາ'] || '0').replace(/,/g, '')) || 0; 
                    closeVal = parseFloat(String(cleanRow['closingbalance'] || cleanRow['ຍອດປິດ'] || cleanRow['ຍອດຄົງເຫຼືອ'] || cleanRow['ຍອດເງິນ'] || '0').replace(/,/g, '')) || 0; 
                } 
                else {
                    let rawAmount = String(cleanRow['balance'] || cleanRow['amount'] || cleanRow['total'] || cleanRow['ຍອດເງິນ'] || cleanRow['ຈຳນວນເງິນ'] || cleanRow['ຍອດຄົງເຫຼືອ'] || '0').replace(/,/g, '');
                    const amountVal = parseFloat(rawAmount) || 0;
                    if (uploadType === 'NEW') { openVal = amountVal; closeVal = amountVal; } else { openVal = 0; closeVal = amountVal; }
                }

                const uKey = `${accNo}_${dateKey}_${currentActivePage}`;
                if (!groupedData[uKey]) groupedData[uKey] = { branch_name: branchName, ascount_no: accNo, customer_name: custName, campaign_type: currentActivePage, date_key: dateKey, opening_balance: openVal, closing_balance: closeVal };
                else { groupedData[uKey].opening_balance += openVal; groupedData[uKey].closing_balance += closeVal; }
            });

            const finalPayload = Object.values(groupedData);
            const { error: dataError } = await _supabase.from('bank_data').upsert(finalPayload, { onConflict: 'ascount_no,date_key,campaign_type' });
            if (dataError) throw dataError;
            if(dateKey === '2025-12-31') await fetchYearEnd2025();
            alert(`✅ ສຳເລັດ! ອັບໂຫລດຈຳນວນ ${finalPayload.length} ລາຍການ`);
            document.getElementById('upload-modal').style.display = 'none'; rawData = []; await loadData();
        } catch (err) { alert("ເກີດຂໍ້ຜິດພາດ: " + err.message); } finally { hideLoading(); fileInput.value = ''; }
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

async function deleteDataByDate() {
    const dateKey = document.getElementById('upload-date-select').value; const branchSelect = document.getElementById('branch-select').value; 
    if (!dateKey && branchSelect === 'all') return alert("⚠️ ກະລຸນາເລືອກ 'ວັນທີ' ໃນກ່ອງນີ້ ຫຼື ອອກໄປເລືອກ 'ສາຂາ' ກ່ອນ!");
    let confirmText = `⚠️ ຢືນຢັນການລຶບຂໍ້ມູນ`; if(dateKey) confirmText += `\n- ວັນທີ: ${dateKey}`; if(branchSelect !== 'all') confirmText += `\n- ສະເພາະສາຂາ: ${branchSelect}`;
    if (!confirm(confirmText)) return;

    let query = _supabase.from('bank_data').delete().eq('campaign_type', currentActivePage);
    if (dateKey) query = query.eq('date_key', dateKey); if (branchSelect !== 'all') query = query.eq('branch_name', branchSelect);

    const { error } = await query;
    if (error) alert("Error: " + error.message); else { alert("✅ ລຶບສຳເລັດ!"); document.getElementById('upload-modal').style.display = 'none'; rawData = []; await loadData(); }
}

function exportToPDF() {
    const originalTable = document.getElementById('branch-table');
    if (!originalTable) return alert("ບໍ່ພົບຂໍ້ມູນທີ່ຈະດາວໂຫຼດ!");
    
    showLoading('ກຳລັງສ້າງໄຟລ໌ PDF...');
    
    const today = new Date().toISOString().slice(0, 10);
    let pdfTitle = "ລາຍງານ";
    if (currentActivePage === 'PUPOM') pdfTitle += " ປູພົມ";
    else if (currentActivePage === 'BEERLAO') pdfTitle += " ເບຍລາວ";
    else if (currentActivePage === 'CAMPAIGN') pdfTitle += " ລູກຄ້າແຄມແປນ";

    // 1. ສ້າງ Container ຂຶ້ນມາໃໝ່ເພື່ອຫຼີກລ່ຽງ CSS ຂອງໜ້າຈໍຫຼັກມາກວນ
    const container = document.createElement('div');
    container.style.position = 'absolute';
    container.style.top = '0';
    container.style.left = '0';
    container.style.zIndex = '-9999'; // ໃຫ້ຢູ່ຫຼັງສຸດເພື່ອບໍ່ໃຫ້ລົບກວນໜ້າຈໍ
    container.style.width = '1400px'; // ບັງຄັບຄວາມກວ້າງບໍ່ໃຫ້ຂໍ້ຄວາມຕົກແຖວ
    container.style.backgroundColor = '#0b1120'; // ພື້ນຫຼັງສີດຳ
    container.style.padding = '40px';
    container.style.color = '#ffffff';
    container.style.fontFamily = '"Noto Sans Lao", sans-serif';

    // 2. ຫົວຂໍ້
    const titleEl = document.createElement('h2');
    titleEl.innerText = `ລາຍລະອຽດຂໍ້ມູນ ${pdfTitle} (ວັນທີ: ${today})`;
    titleEl.style.textAlign = 'center';
    titleEl.style.marginBottom = '20px';
    titleEl.style.color = '#3b82f6';
    container.appendChild(titleEl);

    // 3. ກັອບປີ້ຕາຕະລາງ
    const tableClone = originalTable.cloneNode(true);
    tableClone.style.width = '100%';
    tableClone.style.borderCollapse = 'collapse';
    
    // ບັງຄັບໃສ່ສີ Inline Styles ໃຫ້ທຸກໆຖັນ (ປ້ອງກັນ html2canvas ອ່ານສີຜິດ)
    const ths = tableClone.querySelectorAll('th');
    ths.forEach(th => {
        th.style.backgroundColor = '#1f2937';
        th.style.borderBottom = '1px solid #374151';
        th.style.padding = '12px';
    });

    const tds = tableClone.querySelectorAll('td');
    tds.forEach(td => {
        td.style.borderBottom = '1px solid #374151';
        td.style.padding = '12px';
        if(td.style.color === '') td.style.color = '#ffffff'; // ຖ້າບໍ່ມີສີບັງຄັບໃຫ້ເປັນຂາວ
    });

    container.appendChild(tableClone);
    document.body.appendChild(container); // ແປະໃສ່ DOM ເພື່ອໃຫ້ html2pdf ເຫັນ

    const opt = { 
        margin:       0.3,
        filename:     `${pdfTitle.replace(/\s+/g, '_')}_${today}.pdf`,
        image:        { type: 'jpeg', quality: 1.0 },
        html2canvas:  { scale: 2, useCORS: true, backgroundColor: '#0b1120', windowWidth: 1400 }, 
        jsPDF:        { unit: 'in', format: 'a3', orientation: 'landscape' }
    };

    // 4. ສ້າງ PDF ແລ້ວໂຫຼດ
    html2pdf().set(opt).from(container).output('blob').then(function(pdfBlob) {
        const blobUrl = URL.createObjectURL(pdfBlob);
        const downloadLink = document.createElement('a');
        downloadLink.href = blobUrl;
        downloadLink.download = opt.filename;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(blobUrl);
        
        document.body.removeChild(container); // ລຶບຖິ້ມຫຼັງຈາກໂຫຼດແລ້ວ
        hideLoading();
    }).catch(err => {
        console.error(err);
        document.body.removeChild(container);
        hideLoading();
        alert("ເກີດຂໍ້ຜິດພາດໃນການສ້າງ PDF: " + err.message);
    });
}

function exportToExcel() {
    showLoading('ກຳລັງສ້າງໄຟລ໌ Excel...');
    
    setTimeout(() => {
        try {
            const branch = document.getElementById('branch-select').value; 
            let filtered = rawData;
            
            if (branch !== 'all') filtered = filtered.filter(d => d.branch_name === branch);
            if (filtered.length === 0) {
                hideLoading();
                return alert("ບໍ່ມີຂໍ້ມູນທີ່ຈະດາວໂຫຼດ!");
            }
            
            const uniqueMap = new Map(); 
            filtered.forEach(item => { 
                const existing = uniqueMap.get(item.ascount_no); 
                if (!existing || item.date_key > existing.date_key) uniqueMap.set(item.ascount_no, item); 
            });
            
            const excelData = Array.from(uniqueMap.values()).map((item, index) => ({ 
                "ລ/ດ": index + 1, 
                "ສາຂາ": item.branch_name, 
                "ເລກບັນຊີ": item.ascount_no, 
                "ຊື່ລູກຄ້າ": item.customer_name || 'ບໍ່ລະບຸ',
                "ຍອດຍົກມາ": item.opening_balance || 0,
                "ຍອດຄົງເຫຼືອ": item.closing_balance || 0,
                "ວັນທີອັບເດດ": item.date_key
            }));
            
            const worksheet = XLSX.utils.json_to_sheet(excelData); 
            const workbook = XLSX.utils.book_new(); 
            XLSX.utils.book_append_sheet(workbook, worksheet, "Report"); 
            XLSX.writeFile(workbook, `Report_${currentActivePage}_${new Date().toISOString().slice(0, 10)}.xlsx`);
        } catch (err) {
            console.error(err);
            alert("ເກີດຂໍ້ຜິດພາດໃນການສ້າງ Excel: " + err.message);
        } finally {
            hideLoading();
        }
    }, 500);
}

// ==========================================
// 9. KPI Logic
// ==========================================
async function loadKPIData() {
    const empName = document.getElementById('kpi-employee-select').value; const month = document.getElementById('kpi-month-select').value;
    if(!empName || !month) return;
    const { data, error } = await _supabase.from('marketing_kpi').select('*').eq('employee_name', empName).eq('kpi_month', month).single();

    if (error || !data) {
        document.getElementById('kpi-table-body').innerHTML = `<tr><td colspan="5" style="text-align:center; padding:20px;">ບໍ່ພົບຂໍ້ມູນ</td></tr>`;
        document.getElementById('kpi-total-score').innerText = "0%"; document.getElementById('kpi-status').innerText = "-";
        updateRadarChart([0, 0, 0]); return;
    }

    const depPct = calculatePercent(data.actual_deposit, data.target_deposit); const accPct = calculatePercent(data.actual_new_acc, data.target_new_acc); const visitPct = calculatePercent(data.actual_visit, data.target_visit);
    const totalScore = (depPct * 0.5) + (accPct * 0.3) + (visitPct * 0.2);
    document.getElementById('kpi-total-score').innerText = totalScore.toFixed(1) + "%";
    const statusEl = document.getElementById('kpi-status'); let color = '#ef4444'; let text = "ຕ້ອງປັບປຸງ";
    if(totalScore >= 100) { color = '#10b981'; text = "ດີເລີດ"; } else if(totalScore >= 80) { color = '#3b82f6'; text = "ດີ"; } else if(totalScore >= 60) { color = '#f59e0b'; text = "ປານກາງ"; }
    statusEl.innerText = text; statusEl.style.color = color;

    document.getElementById('kpi-table-body').innerHTML = `
        <tr><td>ຍອດເງິນຝາກ</td><td align="right">${Number(data.target_deposit).toLocaleString()}</td><td align="right">${Number(data.actual_deposit).toLocaleString()}</td><td align="center" style="color:${getColor(depPct)}">${depPct.toFixed(1)}%</td><td align="center">${(depPct * 0.5).toFixed(1)}</td></tr>
        <tr><td>ເປີດບັນຊີໃໝ່</td><td align="right">${data.target_new_acc}</td><td align="right">${data.actual_new_acc}</td><td align="center" style="color:${getColor(accPct)}">${accPct.toFixed(1)}%</td><td align="center">${(accPct * 0.3).toFixed(1)}</td></tr>
        <tr><td>ລົງຢ້ຽມຢາມ</td><td align="right">${data.target_visit}</td><td align="right">${data.actual_visit}</td><td align="center" style="color:${getColor(visitPct)}">${visitPct.toFixed(1)}%</td><td align="center">${(visitPct * 0.2).toFixed(1)}</td></tr>
    `;
    updateRadarChart([depPct, accPct, visitPct]);
}
function calculatePercent(actual, target) { if(!target || target == 0) return 0; let pct = (actual / target) * 100; return pct > 120 ? 120 : pct; }
function getColor(pct) { if(pct >= 100) return '#10b981'; if(pct >= 80) return '#3b82f6'; if(pct >= 60) return '#f59e0b'; return '#ef4444'; }
function updateRadarChart(dataArr) {
    if(!kpiChart) kpiChart = echarts.init(document.getElementById('kpiRadarChart'));
    kpiChart.setOption({ radar: { indicator: [{ name: 'ເງິນຝາກ', max: 120 }, { name: 'ບັນຊີໃໝ່', max: 120 }, { name: 'ຢ້ຽມຢາມ', max: 120 }], splitArea: { areaStyle: { color: ['#1e293b'] } }, axisName: { color: '#fff' } }, series: [{ type: 'radar', data: [{ value: dataArr, name: 'Performance', itemStyle: { color: '#3b82f6' }, areaStyle: { opacity: 0.3 } }] }] });
}

// ==========================================
// 10. Upload History Utils
// ==========================================
let currentRowToReplace = null;
let currentBranchToReplace = null;

async function deleteUpload(rowId, branchName) {
    if (confirm(`⚠️ ທ່ານຕ້ອງການລຶບຂໍ້ມູນທັງໝົດຂອງ "${branchName}" ອອກຈາກຖານຂໍ້ມູນແທ້ຫຼືບໍ່? (ລຶບແລ້ວກູ້ຄືນບໍ່ໄດ້)`)) {
        
        showLoading('ກຳລັງລຶບ...');
        const { error } = await _supabase
            .from('bank_data')
            .delete()
            .eq('branch_name', branchName)
            .eq('campaign_type', currentActivePage); 

        hideLoading();
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
        
        if (row) {
            row.cells[2].innerHTML = `<span style="color: #10b981;">📄 </span><span class="text-body">${newFileName}</span>`;
            const now = new Date();
            const timeString = now.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            row.cells[0].innerHTML = `ມື້ນີ້ <br><small class="text-muted">${timeString}</small>`;
        }
        
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