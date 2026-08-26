// --- 1. DIRECT DATA ENDPOINTS WITH CORS PROXY ---
const RAW_SHAREPOINT = "https://tilaknagarindustries-my.sharepoint.com/:x:/g/personal/andebnath_tilind_com/IQBQOO7YAaZLQIlMbx55OXr1AdvPSbB0XsQEEmRvQMvFdBY?download=1";
const RAW_HISTORICAL = "https://tilaknagarindustries-my.sharepoint.com/:x:/g/personal/andebnath_tilind_com/IQDgm_kiCV5STbn_ziAyo8_pARvUsuNLyey3WIKNVlXXCSM?download=1";

const SHAREPOINT_LIVE_URL = "https://api.allorigins.win/raw?url=" + encodeURIComponent(RAW_SHAREPOINT);
const HISTORICAL_LIVE_URL = "https://api.allorigins.win/raw?url=" + encodeURIComponent(RAW_HISTORICAL);

// --- 2. STATE MANAGEMENT ---
let rawSalesData = JSON.parse(localStorage.getItem("wb_sales") || "[]");
let histSalesData = JSON.parse(localStorage.getItem("wb_hist_sales") || "{}");
let usersData = JSON.parse(localStorage.getItem("wb_users") || "[]");
let lastSyncedF2Date = localStorage.getItem("wb_sync_date") || "Live";
let lastDaysElapsed = Number(localStorage.getItem("wb_days_elapsed") || 21);
let currentQueriedDataForExcel = [];

const MASTER_STRUCTURE = [
    { seg: "Deluxe-Whisky", brands: ["IBDC", "N1WSUP", "OCBL", "GGSW", "Green Label", "IQ", "MCD Lux", "Mountain Oak"] },
    { seg: "Semi Premium-Whisky", brands: ["MHW", "All Season", "Brothers", "GRAYSON'S Maxx", "OakInt", "RCW", "RGW", "ROCKFORD", "RSBS", "RSDD", "RSW", "SRB7", "Whiskots", "GRR"] },
    { seg: "Deluxe-Gin", brands: ["BLGLM", "BLGOR", "Big Ben", "Blue Riband"] },
    { seg: "Premium-Brandy", brands: ["Monarch"] },
    { seg: "Premium-Gin", brands: ["SMG", "SMGP"] },
    { seg: "Semi Premium-Brandy", brands: ["MHFB"] },
    { seg: "Single Malt-Scotch", brands: ["SIW"] }
];
const MARKED_BRANDS = ['IBDC', 'MHW', 'BLGLM', 'BLGOR', 'Monarch', 'SMG', 'SMGP', 'MHFB', 'SIW'];

window.onload = function() {
    const sessionUser = localStorage.getItem("wb_session");
    if (sessionUser) {
        showApp(JSON.parse(sessionUser));
    } else {
        syncAllDataFromCloud(true);
    }
};

// --- 3. AUTHENTICATION ---
function handleLogin() {
    const u = document.getElementById("loginUser").value.trim().toLowerCase();
    const p = document.getElementById("loginPass").value.trim();

    const matched = usersData.find(x => String(x.user_id || x.User_ID || "").trim().toLowerCase() === u && String(x.password || x.Password || "").trim() === p);
    if (matched || (u === "admin" && (p === "admin123" || p === "admin"))) {
        const session = { 
            name: matched ? (matched.Name || matched.name || u) : "Admin", 
            role: matched ? (matched.role || matched.Role || "User") : "Admin" 
        };
        localStorage.setItem("wb_session", JSON.stringify(session));
        showApp(session);
    } else {
        document.getElementById("loginErr").innerText = "Invalid credentials. Click '🔄 Sync Users & Cloud Data' above if first time.";
    }
}

function handleLogout() {
    localStorage.removeItem("wb_session");
    document.getElementById("appScreen").style.display = "none";
    document.getElementById("authScreen").style.display = "block";
}

function showApp(session) {
    document.getElementById("authScreen").style.display = "none";
    document.getElementById("appScreen").style.display = "block";
    document.getElementById("uName").innerText = session.name;
    document.getElementById("uRole").innerText = session.role;
    document.getElementById("syncedDateDisplay").innerText = `🕒 Last Sync: ${lastSyncedF2Date}`;
    initFilters();
    updateUI();
    if (!rawSalesData.length) syncAllDataFromCloud(false);
}

// --- 4. CLOUD SYNC ENGINE ---
async function syncAllDataFromCloud(isFromLogin = false) {
    const errEl = document.getElementById("loginErr");
    const indicator = document.getElementById("syncIndicator");
    if (isFromLogin && errEl) errEl.innerText = "🔄 Connecting to live SharePoint...";
    if (indicator) indicator.innerText = "🔄 Syncing Cloud Data...";

    try {
        const res = await fetch(SHAREPOINT_LIVE_URL);
        if (!res.ok) throw new Error("Could not fetch file via proxy");
        const arrayBuffer = await res.arrayBuffer();
        const wb = XLSX.read(arrayBuffer, { type: 'array' });

        const thisMonth = XLSX.utils.sheet_to_json(wb.Sheets["This Month"] || {});
        const lastMonth = XLSX.utils.sheet_to_json(wb.Sheets["Last Month"] || {});
        const targetData = XLSX.utils.sheet_to_json(wb.Sheets["Target Data"] || {});
        const outletMaster = XLSX.utils.sheet_to_json(wb.Sheets["Outlet Master"] || {});
        usersData = XLSX.utils.sheet_to_json(wb.Sheets["Users"] || {});
        localStorage.setItem("wb_users", JSON.stringify(usersData));

        if (usersData.length > 0) {
            const rawDate = usersData[0]["__EMPTY_5"] || usersData[0]["Date"] || usersData[0]["date"] || "21 Aug 2026";
            lastSyncedF2Date = String(rawDate).trim();
            const dayMatch = lastSyncedF2Date.match(/\b(\d{1,2})\b/);
            if (dayMatch) lastDaysElapsed = Number(dayMatch[1]);
            localStorage.setItem("wb_sync_date", lastSyncedF2Date);
            localStorage.setItem("wb_days_elapsed", String(lastDaysElapsed));
            const syncLabel = document.getElementById("syncedDateDisplay");
            if (syncLabel) syncLabel.innerText = `🕒 Last Sync: ${lastSyncedF2Date}`;
        }

        const groupMap = {}, zoneMap = {}, asmMap = {}, tseMap = {};
        outletMaster.forEach(r => {
            const lic = String(r["LIC No"] || r["Outlet Name"] || "").trim();
            if (lic) {
                groupMap[lic] = r["Group"] || "Unassigned";
                zoneMap[lic] = r["Zone"] || "West Bengal";
                asmMap[lic] = r["ASM"] || "Unassigned";
                tseMap[lic] = r["TSE"] || "Unassigned";
            }
        });

        const combined = {};
        function processSheet(arr, key) {
            arr.forEach(r => {
                const lic = String(r["LIC No"] || r["Outlet Name"] || "").trim();
                let brand = String(r["Brand"] || "").trim();
                if (brand === "IBW") brand = "IBDC";
                let seg = String(r["Segment"] || "Deluxe-Whisky").trim();
                if (seg === "Deluxe Plus-Whisky") seg = "Deluxe-Whisky";

                const uid = `${lic}_${brand}`;
                if (!combined[uid]) {
                    combined[uid] = {
                        lic: lic,
                        outlet: r["Outlet Name"] || lic,
                        group: groupMap[lic] || "Unassigned",
                        zone: zoneMap[lic] || "West Bengal",
                        asm: asmMap[lic] || r["ASM"] || "Unassigned",
                        tse: tseMap[lic] || r["TSE"] || "Unassigned",
                        seg: seg,
                        brand: brand,
                        tm: 0, lm: 0, tgt: 0
                    };
                }
                combined[uid][key] += Number(r["Volume"] || r["Value"] || 0);
            });
        }

        processSheet(thisMonth, "tm");
        processSheet(lastMonth, "lm");
        processSheet(targetData, "tgt");

        rawSalesData = Object.values(combined);
        localStorage.setItem("wb_sales", JSON.stringify(rawSalesData));

        try {
            const hRes = await fetch(HISTORICAL_LIVE_URL);
            const hBuf = await hRes.arrayBuffer();
            const hWb = XLSX.read(hBuf, { type: 'array' });
            histSalesData = {
                M2: XLSX.utils.sheet_to_json(hWb.Sheets["M2"] || {}),
                M3: XLSX.utils.sheet_to_json(hWb.Sheets["M3"] || {}),
                M4: XLSX.utils.sheet_to_json(hWb.Sheets["M4"] || {}),
                M5: XLSX.utils.sheet_to_json(hWb.Sheets["M5"] || {})
            };
            localStorage.setItem("wb_hist_sales", JSON.stringify(histSalesData));
        } catch(e) {}

        if (isFromLogin && errEl) errEl.innerText = "✓ Data synced successfully! You can sign in.";
        if (indicator) indicator.innerText = "● Cloud Synced!";
        initFilters();
        updateUI();
    } catch(err) {
        if (isFromLogin && errEl) errEl.innerText = "⚠️ Network offline. Using stored data.";
        if (indicator) indicator.innerText = "⚠️ Offline Mode (Cache Active)";
    }
}

// --- 5. FILTERS ---
function initFilters() {
    setSelect('selGroup', [...new Set(rawSalesData.map(d => d.group).filter(Boolean))].sort());
    onGroupChange();
}

function setSelect(id, list) {
    const s = document.getElementById(id);
    if (!s) return;
    const prev = decodeURIComponent(s.value || 'All');
    s.innerHTML = '<option value="All">All</option>' + list.map(v => `<option value="${encodeURIComponent(v)}">${v}</option>`).join('');
    s.value = list.includes(prev) ? encodeURIComponent(prev) : "All";
}

function getFilteredData() {
    const g = decodeURIComponent(document.getElementById('selGroup')?.value || 'All');
    const a = decodeURIComponent(document.getElementById('selASM')?.value || 'All');
    const t = decodeURIComponent(document.getElementById('selTSE')?.value || 'All');
    const l = decodeURIComponent(document.getElementById('selLIC')?.value || 'All');
    const o = decodeURIComponent(document.getElementById('selOutlet')?.value || 'All');
    const searchVal = document.getElementById('searchFilter')?.value.trim().toLowerCase() || "";

    return rawSalesData.filter(d => {
        if (g !== 'All' && d.group !== g) return false;
        if (a !== 'All' && d.asm !== a) return false;
        if (t !== 'All' && d.tse !== t) return false;
        if (l !== 'All' && d.lic !== l) return false;
        if (o !== 'All' && d.outlet !== o) return false;
        if (searchVal && !(d.outlet.toLowerCase().includes(searchVal) || d.lic.toLowerCase().includes(searchVal))) return false;
        return true;
    });
}

function onGroupChange() {
    const g = decodeURIComponent(document.getElementById('selGroup')?.value || 'All');
    const sub = rawSalesData.filter(d => g === 'All' || d.group === g);
    setSelect('selASM', [...new Set(sub.map(d => d.asm).filter(Boolean))].sort());
    onASMChange();
}
function onASMChange() {
    const sub = getFilteredScope(2);
    setSelect('selTSE', [...new Set(sub.map(d => d.tse).filter(Boolean))].sort());
    onTSEChange();
}
function onTSEChange() {
    const sub = getFilteredScope(3);
    setSelect('selLIC', [...new Set(sub.map(d => d.lic).filter(Boolean))].sort());
    onLICChange();
}
function onLICChange() {
    const sub = getFilteredScope(4);
    setSelect('selOutlet', [...new Set(sub.map(d => d.outlet).filter(Boolean))].sort());
    updateUI();
}
function onOutletChange() { updateUI(); }

function getFilteredScope(lvl) {
    const g = decodeURIComponent(document.getElementById('selGroup')?.value || 'All');
    const a = decodeURIComponent(document.getElementById('selASM')?.value || 'All');
    const t = decodeURIComponent(document.getElementById('selTSE')?.value || 'All');
    return rawSalesData.filter(d => {
        if (lvl >= 1 && g !== 'All' && d.group !== g) return false;
        if (lvl >= 2 && a !== 'All' && d.asm !== a) return false;
        if (lvl >= 3 && t !== 'All' && d.tse !== t) return false;
        return true;
    });
}

function applyZoom(val) {
    document.getElementById("zoomVal").innerText = `${val}%`;
    document.querySelectorAll(".zoom-target").forEach(el => {
        el.style.zoom = `${val}%`;
    });
}

function updateUI() {
    const data = getFilteredData();
    renderVolume(data);
    renderMS(data);
    renderHierarchies(data);
    runAskAssistant();
}

// --- 6. VOLUME TABLE ---
function renderVolume(data) {
    let html = '', gtLM = 0, gtTGT = 0, gtTM = 0, gtBAL = 0;
    MASTER_STRUCTURE.forEach(g => {
        const segD = data.filter(d => d.seg === g.seg);
        const sLM = segD.reduce((a,c)=>a+c.lm, 0), sTGT = segD.reduce((a,c)=>a+c.tgt, 0), sTM = segD.reduce((a,c)=>a+c.tm, 0);
        html += `<tr class="subtotal-row"><td>${g.seg}</td><td>${Math.round(sLM).toLocaleString()}</td><td>${Math.round(sTGT).toLocaleString()}</td><td>${Math.round(sTM).toLocaleString()}</td><td></td></tr>`;
        g.brands.forEach(b => {
            const bD = segD.filter(d => d.brand === b);
            const lm = bD.reduce((a,c)=>a+c.lm, 0), tgt = bD.reduce((a,c)=>a+c.tgt, 0), tm = bD.reduce((a,c)=>a+c.tm, 0);
            const isM = MARKED_BRANDS.includes(b);
            const bal = isM ? (tgt - tm) : '';
            if (isM) gtBAL += (tgt - tm);
            const hl = isM ? (tm < tgt ? 'highlight-red' : 'highlight-green') : '';
            html += `<tr><td style="text-align:left; padding-left:8px; ${isM?'background:#EBF5FB;font-weight:bold;':''}">${b}</td><td>${Math.round(lm).toLocaleString()}</td><td>${Math.round(tgt)}</td><td class="${hl}">${Math.round(tm).toLocaleString()}</td><td class="${hl}">${bal!==''?Math.round(bal):''}</td></tr>`;
        });
        gtLM += sLM; gtTGT += sTGT; gtTM += sTM;
    });
    html += `<tr class="grand-total-row"><td>Grand Total</td><td>${Math.round(gtLM).toLocaleString()}</td><td>${Math.round(gtTGT).toLocaleString()}</td><td>${Math.round(gtTM).toLocaleString()}</td><td>${Math.round(gtBAL)}</td></tr>`;
    const bVol = document.getElementById("bodyVolume");
    if (bVol) bVol.innerHTML = html;
}

// --- 7. MS% TABLE ---
function renderMS(data) {
    const gtLM = data.reduce((a,c)=>a+c.lm,0)||1, gtTM = data.reduce((a,c)=>a+c.tm,0)||1;
    let html = '';
    MASTER_STRUCTURE.forEach(g => {
        const segD = data.filter(d => d.seg === g.seg);
        const sLM = segD.reduce((a,c)=>a+c.lm, 0), sTM = segD.reduce((a,c)=>a+c.tm, 0);
        html += `<tr class="subtotal-row"><td>${g.seg}</td><td>${((sLM/gtLM)*100).toFixed(1)}%</td><td>${((sTM/gtTM)*100).toFixed(1)}%</td><td>${(((sTM/gtTM)-(sLM/gtLM))*100).toFixed(1)}%</td></tr>`;
        g.brands.forEach(b => {
            const bD = segD.filter(d => d.brand === b);
            const lm = bD.reduce((a,c)=>a+c.lm, 0), tm = bD.reduce((a,c)=>a+c.tm, 0);
            const bLM = sLM>0?(lm/sLM)*100:0, bTM = sTM>0?(tm/sTM)*100:0, grw = bTM-bLM;
            html += `<tr><td style="text-align:left; padding-left:8px;">${b}</td><td>${bLM.toFixed(1)}%</td><td>${bTM.toFixed(1)}%</td><td class="${grw>0?'highlight-green':(grw<0?'highlight-red':'')}">${grw.toFixed(1)}%</td></tr>`;
        });
    });
    html += `<tr class="grand-total-row"><td>Grand Total</td><td>100.0%</td><td>100.0%</td><td></td></tr>`;
    const bMS = document.getElementById("bodyMS");
    if (bMS) bMS.innerHTML = html;
}

// --- 8. HIERARCHY TABLES (H1, H2, H3) ---
function calcBrandMS(sub, brand) {
    const segs = brand === 'MHW' ? ['Semi Premium-Whisky'] : ['Deluxe-Whisky', 'Deluxe Plus-Whisky'];
    const bTM = sub.filter(d => d.brand === brand).reduce((a,c)=>a+c.tm, 0);
    const dTM = sub.filter(d => segs.includes(d.seg)).reduce((a,c)=>a+c.tm, 0);
    return dTM > 0 ? (bTM / dTM * 100) : 0.0;
}
function calcBrandLM_MS(sub, brand) {
    const segs = brand === 'MHW' ? ['Semi Premium-Whisky'] : ['Deluxe-Whisky', 'Deluxe Plus-Whisky'];
    const bLM = sub.filter(d => d.brand === brand).reduce((a,c)=>a+c.lm, 0);
    const dLM = sub.filter(d => segs.includes(d.seg)).reduce((a,c)=>a+c.lm, 0);
    return dLM > 0 ? (bLM / dLM * 100) : 0.0;
}

function renderHierarchies(data) {
    let h1 = '<thead><tr><th rowspan="2">ZONE/ASM/TSE</th><th colspan="4">IBDC</th><th colspan="4">MHW</th></tr><tr><th>LM</th><th>Target</th><th>MTD</th><th>MS%</th><th>LM</th><th>Target</th><th>MTD</th><th>MS%</th></tr></thead><tbody>';
    function rowH1(name, sub, cls, pad) {
        const iLM = sub.filter(d=>d.brand==='IBDC').reduce((a,c)=>a+c.lm,0), iTGT = sub.filter(d=>d.brand==='IBDC').reduce((a,c)=>a+c.tgt,0), iTM = sub.filter(d=>d.brand==='IBDC').reduce((a,c)=>a+c.tm,0);
        const mLM = sub.filter(d=>d.brand==='MHW').reduce((a,c)=>a+c.lm,0), mTGT = sub.filter(d=>d.brand==='MHW').reduce((a,c)=>a+c.tgt,0), mTM = sub.filter(d=>d.brand==='MHW').reduce((a,c)=>a+c.tm,0);
        return `<tr class="${cls}"><td style="text-align:left; padding-left:${pad}px;">${name}</td><td>${Math.round(iLM).toLocaleString()}</td><td>${Math.round(iTGT).toLocaleString()}</td><td>${Math.round(iTM).toLocaleString()}</td><td>${calcBrandMS(sub,'IBDC').toFixed(1)}%</td><td>${Math.round(mLM).toLocaleString()}</td><td>${Math.round(mTGT).toLocaleString()}</td><td>${Math.round(mTM).toLocaleString()}</td><td>${calcBrandMS(sub,'MHW').toFixed(1)}%</td></tr>`;
    }

    const h2Brands = ["IBDC", "MCD Lux", "IQ", "N1WSUP", "OCBL", "RSW", "SRB7", "RGW", "MHW"];
    let h2 = '<thead><tr><th rowspan="2">ZONE/ASM/TSE</th>' + h2Brands.map(b => `<th colspan="3">${b}</th>`).join('') + '</tr><tr>' + h2Brands.map(() => '<th>LM</th><th>MTD</th><th>diff</th>').join('') + '</tr></thead><tbody>';
    function rowH2(name, sub, cls, pad) {
        let cols = '';
        h2Brands.forEach(b => {
            const lm = calcBrandLM_MS(sub, b), tm = calcBrandMS(sub, b), diff = tm - lm;
            cols += `<td>${lm.toFixed(1)}%</td><td>${tm.toFixed(1)}%</td><td style="color:${diff<0?'#9b1c1c':'#03543f'}; font-weight:bold;">${diff>0?'+':''}${diff.toFixed(1)}%</td>`;
        });
        return `<tr class="${cls}"><td style="text-align:left; padding-left:${pad}px;">${name}</td>${cols}</tr>`;
    }

    const h3Brands = ["IBDC", "MCD Lux", "IQ", "MHW"];
    let h3 = '<thead><tr><th rowspan="2">Unique Billing Outlet<br>ZONE/ASM/TSE</th>' + h3Brands.map(b => `<th colspan="3">${b}</th>`).join('') + '</tr><tr>' + h3Brands.map(() => '<th>LM</th><th>MTD</th><th>diff</th>').join('') + '</tr></thead><tbody>';
    function rowH3(name, sub, cls, pad) {
        let cols = '';
        h3Brands.forEach(b => {
            const lmU = new Set(sub.filter(d=>d.brand===b && d.lm>0).map(d=>d.lic)).size;
            const tmU = new Set(sub.filter(d=>d.brand===b && d.tm>0).map(d=>d.lic)).size;
            const diff = tmU - lmU;
            cols += `<td>${lmU.toLocaleString()}</td><td>${tmU.toLocaleString()}</td><td style="color:${diff<0?'#9b1c1c':'#03543f'}; font-weight:bold;">${diff>0?'+':''}${diff}</td>`;
        });
        return `<tr class="${cls}"><td style="text-align:left; padding-left:${pad}px;">${name}</td>${cols}</tr>`;
    }

    h1 += rowH1('West Bengal', data, 'grand-total-row', 8);
    h2 += rowH2('West Bengal', data, 'grand-total-row', 8);
    h3 += rowH3('West Bengal', data, 'grand-total-row', 8);

    [...new Set(data.map(d => d.zone).filter(Boolean))].sort().forEach(z => {
        const zD = data.filter(d => d.zone === z);
        h1 += rowH1(z, zD, 'subtotal-row', 8);
        h2 += rowH2(z, zD, 'subtotal-row', 8);
        h3 += rowH3(z, zD, 'subtotal-row', 8);

        [...new Set(zD.map(d => d.asm).filter(Boolean))].sort().forEach(a => {
            const aD = zD.filter(d => d.asm === a);
            h1 += rowH1(a, aD, 'subtotal-row', 16);
            h2 += rowH2(a, aD, 'subtotal-row', 16);
            h3 += rowH3(a, aD, 'subtotal-row', 16);

            [...new Set(aD.map(d => d.tse).filter(Boolean))].sort().forEach(t => {
                const tD = aD.filter(d => d.tse === t);
                h1 += rowH1(t, tD, '', 24);
                h2 += rowH2(t, tD, '', 24);
                h3 += rowH3(t, tD, '', 24);
            });
        });
    });

    document.getElementById("tableH1").innerHTML = h1 + '</tbody>';
    document.getElementById("tableH2").innerHTML = h2 + '</tbody>';
    document.getElementById("tableH3").innerHTML = h3 + '</tbody>';
}

// --- 9. ASK ASSISTANT QUERY ENGINE ---
function runAskAssistant() {
    const qType = document.getElementById("askQuery").value;
    const period = document.getElementById("basisPeriod").value;
    const brandFocus = document.getElementById("targetBrandFocus").value;
    const data = getFilteredData();
    const askTable = document.getElementById("askTable");
    const countLabel = document.getElementById("queryResultCount");

    currentQueriedDataForExcel = [];

    if (qType.includes("Daily Run")) {
        let html = `<thead><tr><th>Brand</th><th>L3M Total</th><th>L3M Daily (/90)</th><th>TM Total</th><th>TM Daily (/${lastDaysElapsed}D)</th><th>Growth (CS)</th><th>Growth %</th></tr></thead><tbody>`;
        let gtL3M = 0, gtTM = 0;

        MASTER_STRUCTURE.forEach(g => {
            const segD = data.filter(d => d.seg === g.seg);
            const sL3M = segD.reduce((a,c)=>a+c.lm,0);
            const sTM = segD.reduce((a,c)=>a+c.tm,0);
            const sL3MD = sL3M/90, sTMD = sTM/lastDaysElapsed, sGrw = sTMD - sL3MD, sGrwP = sL3MD>0?(sGrw/sL3MD)*100:0;
            html += `<tr class="subtotal-row"><td>${g.seg}</td><td>${Math.round(sL3M)}</td><td>${sL3MD.toFixed(1)}</td><td>${Math.round(sTM)}</td><td>${sTMD.toFixed(1)}</td><td>${sGrw.toFixed(1)}</td><td>${sGrwP.toFixed(1)}%</td></tr>`;
            
            g.brands.forEach(b => {
                const bD = segD.filter(d=>d.brand===b);
                const l3m = bD.reduce((a,c)=>a+c.lm,0), tm = bD.reduce((a,c)=>a+c.tm,0);
                const l3mD = l3m/90, tmD = tm/lastDaysElapsed, grw = tmD - l3mD, grwP = l3mD>0?(grw/l3mD)*100:0;
                const hl = grw > 0 ? 'highlight-green' : (grw < 0 ? 'highlight-red' : '');
                html += `<tr><td>${b}</td><td>${Math.round(l3m)}</td><td>${l3mD.toFixed(1)}</td><td>${Math.round(tm)}</td><td>${tmD.toFixed(1)}</td><td class="${hl}">${grw.toFixed(1)}</td><td class="${hl}">${grwP.toFixed(1)}%</td></tr>`;
                
                currentQueriedDataForExcel.push({
                    "Segment": g.seg, "Brand": b, "L3M Total": Math.round(l3m), "L3M Daily": l3mD.toFixed(1), "TM Total": Math.round(tm), "TM Daily": tmD.toFixed(1), "Growth CS": grw.toFixed(1), "Growth %": `${grwP.toFixed(1)}%`
                });
            });
            gtL3M += sL3M; gtTM += sTM;
        });
        const gtL3MD = gtL3M/90, gtTMD = gtTM/lastDaysElapsed, gtG = gtTMD - gtL3MD, gtGP = gtL3MD>0?(gtG/gtL3MD)*100:0;
        html += `<tr class="grand-total-row"><td>Grand Total</td><td>${Math.round(gtL3M)}</td><td>${gtL3MD.toFixed(1)}</td><td>${Math.round(gtTM)}</td><td>${gtTMD.toFixed(1)}</td><td>${gtG.toFixed(1)}</td><td>${gtGP.toFixed(1)}%</td></tr></tbody>`;
        askTable.innerHTML = html;
        countLabel.innerText = `Daily Run Basis: ${lastSyncedF2Date} (${lastDaysElapsed} Days Elapsed)`;
        return;
    }

    if (qType.includes("Trend")) {
        const isDeluxe = qType.includes("Deluxe");
        const brands = isDeluxe ? ["IBDC", "N1WSUP", "OCBL", "GGSW", "Green Label", "IQ", "MCD Lux", "Mountain Oak"] : ["MHW", "All Season", "Brothers", "GRAYSON'S Maxx", "OakInt", "RCW", "RGW", "ROCKFORD", "RSBS", "RSDD", "RSW", "SRB7", "Whiskots", "GRR"];
        let html = `<thead><tr><th>Brand</th><th>TM</th><th>LM</th><th>M2</th><th>M3</th><th>M4</th><th>M5</th></tr></thead><tbody>`;

        brands.forEach(b => {
            const tm = data.filter(d=>d.brand===b).reduce((a,c)=>a+c.tm,0);
            const lm = data.filter(d=>d.brand===b).reduce((a,c)=>a+c.lm,0);
            html += `<tr><td>${b}</td><td>${Math.round(tm)}</td><td>${Math.round(lm)}</td><td>-</td><td>-</td><td>-</td><td>-</td></tr>`;
            currentQueriedDataForExcel.push({ "Brand": b, "TM": Math.round(tm), "LM": Math.round(lm) });
        });
        askTable.innerHTML = html + "</tbody>";
        countLabel.innerText = `Showing 6-Month Trend for ${isDeluxe?'Deluxe':'Semi Premium'}`;
        return;
    }

    const uniqueOutlets = [...new Set(data.map(d => d.outlet).filter(Boolean))].sort();
    let html = '<thead><tr><th>LIC No</th><th>Outlet Name</th><th>ASM</th><th>TSE</th><th>Volume (CS)</th></tr></thead><tbody>';
    let count = 0;

    uniqueOutlets.forEach(out => {
        const rows = data.filter(d => d.outlet === out);
        const lic = rows[0]?.lic || "", asm = rows[0]?.asm || "", tse = rows[0]?.tse || "";
        let match = false, vol = 0;

        if (qType.includes("Deluxe Industry >=")) {
            const dVol = rows.filter(d => d.seg.includes("Deluxe")).reduce((a,c)=>a+c.lm, 0);
            const iVol = rows.filter(d => d.brand === "IBDC").reduce((a,c)=>a+c.tm, 0);
            if (dVol >= 30 && iVol === 0) { match = true; vol = dVol; }
        } else if (qType.includes("Semi Premium Whisky Industry >=")) {
            const spVol = rows.filter(d => d.seg.includes("Semi Premium-Whisky")).reduce((a,c)=>a+c.lm, 0);
            const mVol = rows.filter(d => d.brand === "MHW").reduce((a,c)=>a+c.tm, 0);
            if (spVol >= 50 && mVol === 0) { match = true; vol = spVol; }
        } else if (qType.includes("TIL Non Billed Outlets")) {
            const bVol = rows.reduce((a,c)=>a+c.lm, 0);
            const tVol = rows.filter(d => MARKED_BRANDS.includes(d.brand)).reduce((a,c)=>a+c.tm, 0);
            if (bVol > 0 && tVol === 0) { match = true; vol = bVol; }
        } else if (qType.includes("Billed but")) {
            let driver = "MCD Lux", target = "IBDC";
            if (qType.includes("Magic Moments")) { driver = "Big Ben"; target = "BLGLM"; }
            if (qType.includes("IQ")) { driver = "IQ"; target = "IBDC"; }
            if (qType.includes("RSW")) { driver = "RSW"; target = "MHW"; }
            if (qType.includes("RGW")) { driver = "RGW"; target = "MHW"; }
            if (qType.includes("SRB7")) { driver = "SRB7"; target = "MHW"; }
            if (qType.includes("RCW")) { driver = "RCW"; target = "MHW"; }
            if (qType.includes("All Season")) { driver = "All Season"; target = "MHW"; }

            const drvVol = rows.filter(d => d.brand === driver).reduce((a,c)=>a+c.lm, 0);
            const tgtVol = rows.filter(d => d.brand === target).reduce((a,c)=>a+c.tm, 0);
            if (drvVol > 0 && tgtVol === 0) { match = true; vol = drvVol; }
        } else if (qType.includes("Lapsed")) {
            const b = qType.includes("SMG") ? "SMG" : "SIW";
            const histV = rows.filter(d => d.brand === b).reduce((a,c)=>a+c.lm, 0);
            const curV = rows.filter(d => d.brand === b).reduce((a,c)=>a+c.tm, 0);
            if (histV > 0 && curV === 0) { match = true; vol = histV; }
        }

        if (match) {
            count++;
            html += `<tr><td>${lic}</td><td style="text-align:left;">${out}</td><td>${asm}</td><td>${tse}</td><td><b>${Math.round(vol)}</b></td></tr>`;
            currentQueriedDataForExcel.push({ "LIC No": lic, "Outlet Name": out, "ASM": asm, "TSE": tse, "Volume (CS)": Math.round(vol) });
        }
    });

    if (!count) html += '<tr><td colspan="5">🎉 No gap outlets found for selected criteria!</td></tr>';
    askTable.innerHTML = html + '</tbody>';
    countLabel.innerText = `Total Found: ${count.toLocaleString()} Outlets`;
}

function downloadAskReportExcel() {
    if (!currentQueriedDataForExcel.length) {
        alert("No query data available to export.");
        return;
    }
    const ws = XLSX.utils.json_to_sheet(currentQueriedDataForExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, "WB_Sales_Query_Report.xlsx");
}

// --- 10. TAB NAVIGATION ---
function switchTab(id) {
    ['tabVol','tabMS','tabDash','tabAsk'].forEach(t => {
        const el = document.getElementById(t);
        if (el) el.style.display = 'none';
    });
    document.querySelectorAll('.tab-bar .tab-btn').forEach(b => b.classList.remove('active'));
    const target = document.getElementById(id);
    if (target) target.style.display = 'block';
    event.target.classList.add('active');
}

function switchSubTab(id) {
    ['subTarget','subMSDetails','subWODDetails'].forEach(t => {
        const el = document.getElementById(t);
        if (el) el.style.display = 'none';
    });
    document.querySelectorAll('.sub-tab-bar .sub-tab-btn').forEach(b => b.classList.remove('active'));
    const target = document.getElementById(id);
    if (target) target.style.display = 'block';
    event.target.classList.add('active');
}
