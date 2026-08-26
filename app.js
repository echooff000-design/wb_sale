const SHAREPOINT_URL = "https://tilaknagarindustries-my.sharepoint.com/:x:/g/personal/andebnath_tilind_com/IQDgm_kiCV5STbn_ziAyo8_pARvUsuNLyey3WIKNVlXXCSM?download=1";

let rawSalesData = JSON.parse(localStorage.getItem("wb_sales") || "[]");
let usersData = JSON.parse(localStorage.getItem("wb_users") || "[]");

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
    }
};

function handleLogin() {
    const u = document.getElementById("loginUser").value.trim().toLowerCase();
    const p = document.getElementById("loginPass").value.trim();

    const matched = usersData.find(x => String(x.user_id).trim().toLowerCase() === u && String(x.password).trim() === p);
    if (matched || (u === "admin" && p === "admin123")) {
        const session = { name: matched ? matched.Name : "Admin", role: matched ? (matched.role || "User") : "Admin" };
        localStorage.setItem("wb_session", JSON.stringify(session));
        showApp(session);
    } else {
        document.getElementById("loginErr").innerText = "Invalid credentials. Tap 'Sync Cloud' if first launch.";
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
    initFilters();
    updateUI();
    if (!rawSalesData.length) syncDataFromCloud();
}

async function syncDataFromCloud() {
    const indicator = document.getElementById("syncIndicator");
    indicator.innerText = "🔄 Syncing Cloud...";
    try {
        const res = await fetch(SHAREPOINT_URL);
        const arrayBuffer = await res.arrayBuffer();
        const wb = XLSX.read(arrayBuffer, { type: 'array' });

        const thisMonth = XLSX.utils.sheet_to_json(wb.Sheets["This Month"] || {});
        const lastMonth = XLSX.utils.sheet_to_json(wb.Sheets["Last Month"] || {});
        const targetData = XLSX.utils.sheet_to_json(wb.Sheets["Target Data"] || {});
        const outletMaster = XLSX.utils.sheet_to_json(wb.Sheets["Outlet Master"] || {});
        usersData = XLSX.utils.sheet_to_json(wb.Sheets["Users"] || {});
        localStorage.setItem("wb_users", JSON.stringify(usersData));

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

        indicator.innerText = "● Cloud Synced!";
        initFilters();
        updateUI();
    } catch(err) {
        indicator.innerText = "⚠️ Offline Mode (Cache Active)";
    }
}

function initFilters() {
    setSelect('selGroup', [...new Set(rawSalesData.map(d => d.group).filter(Boolean))].sort());
    onGroupChange();
}

function setSelect(id, list) {
    const s = document.getElementById(id);
    const prev = decodeURIComponent(s.value || 'All');
    s.innerHTML = '<option value="All">All</option>' + list.map(v => `<option value="${encodeURIComponent(v)}">${v}</option>`).join('');
    s.value = list.includes(prev) ? encodeURIComponent(prev) : "All";
}

function getFilteredData() {
    const g = decodeURIComponent(document.getElementById('selGroup').value);
    const a = decodeURIComponent(document.getElementById('selASM').value);
    const t = decodeURIComponent(document.getElementById('selTSE').value);
    const l = decodeURIComponent(document.getElementById('selLIC').value);
    const o = decodeURIComponent(document.getElementById('selOutlet').value);

    return rawSalesData.filter(d => {
        if (g !== 'All' && d.group !== g) return false;
        if (a !== 'All' && d.asm !== a) return false;
        if (t !== 'All' && d.tse !== t) return false;
        if (l !== 'All' && d.lic !== l) return false;
        if (o !== 'All' && d.outlet !== o) return false;
        return true;
    });
}

function onGroupChange() {
    const g = decodeURIComponent(document.getElementById('selGroup').value);
    const sub = rawSalesData.filter(d => g === 'All' || d.group === g);
    setSelect('selASM', [...new Set(sub.map(d => d.asm).filter(Boolean))].sort());
    onASMChange();
}
function onASMChange() {
    const a = decodeURIComponent(document.getElementById('selASM').value);
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
    const g = decodeURIComponent(document.getElementById('selGroup').value);
    const a = decodeURIComponent(document.getElementById('selASM').value);
    const t = decodeURIComponent(document.getElementById('selTSE').value);
    return rawSalesData.filter(d => {
        if (lvl >= 1 && g !== 'All' && d.group !== g) return false;
        if (lvl >= 2 && a !== 'All' && d.asm !== a) return false;
        if (lvl >= 3 && t !== 'All' && d.tse !== t) return false;
        return true;
    });
}

function updateUI() {
    const data = getFilteredData();
    renderVolume(data);
    renderMS(data);
    renderHierarchies(data);
    runAskAssistant();
}

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
    document.getElementById("bodyVolume").innerHTML = html;
}

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
    document.getElementById("bodyMS").innerHTML = html;
}

function calcBrandMS(sub, brand) {
    const segs = brand === 'MHW' ? ['Semi Premium-Whisky'] : ['Deluxe-Whisky', 'Deluxe Plus-Whisky'];
    const bTM = sub.filter(d => d.brand === brand).reduce((a,c)=>a+c.tm, 0);
    const dTM = sub.filter(d => segs.includes(d.seg)).reduce((a,c)=>a+c.tm, 0);
    return dTM > 0 ? (bTM / dTM * 100) : 0.0;
}

function renderHierarchies(data) {
    let h1 = '<thead><tr><th rowspan="2">ZONE/ASM/TSE</th><th colspan="4">IBDC</th><th colspan="4">MHW</th></tr><tr><th>LM</th><th>Target</th><th>MTD</th><th>MS%</th><th>LM</th><th>Target</th><th>MTD</th><th>MS%</th></tr></thead><tbody>';
    function row(name, sub, cls, pad) {
        const iLM = sub.filter(d=>d.brand==='IBDC').reduce((a,c)=>a+c.lm,0), iTGT = sub.filter(d=>d.brand==='IBDC').reduce((a,c)=>a+c.tgt,0), iTM = sub.filter(d=>d.brand==='IBDC').reduce((a,c)=>a+c.tm,0);
        const mLM = sub.filter(d=>d.brand==='MHW').reduce((a,c)=>a+c.lm,0), mTGT = sub.filter(d=>d.brand==='MHW').reduce((a,c)=>a+c.tgt,0), mTM = sub.filter(d=>d.brand==='MHW').reduce((a,c)=>a+c.tm,0);
        return `<tr class="${cls}"><td style="text-align:left; padding-left:${pad}px;">${name}</td><td>${Math.round(iLM).toLocaleString()}</td><td>${Math.round(iTGT).toLocaleString()}</td><td>${Math.round(iTM).toLocaleString()}</td><td>${calcBrandMS(sub,'IBDC').toFixed(1)}%</td><td>${Math.round(mLM).toLocaleString()}</td><td>${Math.round(mTGT).toLocaleString()}</td><td>${Math.round(mTM).toLocaleString()}</td><td>${calcBrandMS(sub,'MHW').toFixed(1)}%</td></tr>`;
    }
    h1 += row('West Bengal', data, 'grand-total-row', 8);
    [...new Set(data.map(d => d.zone).filter(Boolean))].sort().forEach(z => {
        const zD = data.filter(d => d.zone === z);
        h1 += row(z, zD, 'subtotal-row', 8);
        [...new Set(zD.map(d => d.asm).filter(Boolean))].sort().forEach(a => {
            const aD = zD.filter(d => d.asm === a);
            h1 += row(a, aD, 'subtotal-row', 16);
            [...new Set(aD.map(d => d.tse).filter(Boolean))].sort().forEach(t => {
                h1 += row(t, aD.filter(d => d.tse === t), '', 24);
            });
        });
    });
    document.getElementById("tableH1").innerHTML = h1 + '</tbody>';
}

function runAskAssistant() {
    const data = getFilteredData();
    const uniqueOutlets = [...new Set(data.map(d => d.outlet).filter(Boolean))].sort();
    let html = '<thead><tr><th>LIC No</th><th>Outlet Name</th><th>ASM</th><th>TSE</th><th>Volume (CS)</th></tr></thead><tbody>';
    let cnt = 0;
    uniqueOutlets.forEach(out => {
        const rows = data.filter(d => d.outlet === out);
        const dVol = rows.filter(d => d.seg && d.seg.includes('Deluxe')).reduce((a,c)=>a + c.tm, 0);
        const iVol = rows.filter(d => d.brand === 'IBDC').reduce((a,c)=>a + c.tm, 0);
        if (dVol >= 30 && iVol === 0) {
            cnt++;
            html += `<tr><td>${rows[0].lic}</td><td style="text-align:left;">${rows[0].outlet}</td><td>${rows[0].asm}</td><td>${rows[0].tse}</td><td><b>${Math.round(dVol)}</b></td></tr>`;
        }
    });
    if (!cnt) html += '<tr><td colspan="5">🎉 No gap outlets found!</td></tr>';
    document.getElementById('askTable').innerHTML = html + '</tbody>';
}

function switchTab(id) {
    ['tabVol','tabMS','tabDash','tabAsk'].forEach(t => document.getElementById(t).style.display = 'none');
    document.querySelectorAll('.tab-bar .tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(id).style.display = 'block';
    event.target.classList.add('active');
}

function switchSubTab(id) {
    ['subTarget','subMSDetails','subWODDetails'].forEach(t => document.getElementById(t).style.display = 'none');
    document.querySelectorAll('.sub-tab-bar .sub-tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(id).style.display = 'block';
    event.target.classList.add('active');
}