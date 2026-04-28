// 7개 일용직 양식 정의 — 회사별 분리 노출
// 휴먼 = KM 계열 청구·지급 + 금전출납(세무용) 휴먼
// 채움 = 채움 청구·외국인·지급 + 금전출납(세무용) 채움
// headerRow = 첫 시트에서 헤더가 있는 행(1-based). 위 행은 병합 타이틀/합계/그룹부제이므로 제외.
const FORMS = [
    { id: 'f1', icon: '📊', title: '업체별 (청구용·KM)', file: '2026.04 업체별(청구용)_KM 휴먼.xlsx',     payCol: 10, co: 'human',  headerRow: 2 },
    { id: 'f2', icon: '📅', title: '일자별 (지급용·KM)', file: '2026.04 일자별(지급용)_KM 휴먼.xlsx',     payCol: 10, co: 'human',  headerRow: 3 },
    { id: 'f7', icon: '💵', title: '금전출납 (세무용)',  file: '2026.04-금전출납(세무용)_휴먼.xlsx',       payCol: 2,  co: 'human',  headerRow: 4 },
    { id: 'f3', icon: '🏢', title: '업체별 (청구용)',    file: '2026.04-업체별(청구용)_채움.xlsx',         payCol: 10, co: 'chaeum', headerRow: 2 },
    { id: 'f4', icon: '🌐', title: '업체별 외국인',      file: '2026.04-업체별(청구용)_외국인 채움.xlsx',  payCol: 10, co: 'chaeum', headerRow: 2 },
    { id: 'f5', icon: '📆', title: '일자별 (지급용)',    file: '2026.04-일자별(지급용)_채움.xlsx',         payCol: 10, co: 'chaeum', headerRow: 3 },
    { id: 'f6', icon: '💰', title: '금전출납 (세무용)',  file: '2026.04-금전출납(세무용)_채움.xlsx',       payCol: 2,  co: 'chaeum', headerRow: 4 }
];

document.addEventListener('DOMContentLoaded', function() {
    buildFormTiles('h');
    buildFormTiles('c');
    checkExistingSession();
    refreshAdminPanel();
    initPurchaseChart();
    startDigitalClock();
});

function startDigitalClock() {
    const dateEl = document.getElementById('clock-date');
    const timeEl = document.getElementById('clock-time');
    if (!dateEl || !timeEl) return;
    const days = ['일','월','화','수','목','금','토'];
    const update = () => {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const day = days[now.getDay()];
        const hh = String(now.getHours()).padStart(2, '0');
        const mi = String(now.getMinutes()).padStart(2, '0');
        const ss = String(now.getSeconds()).padStart(2, '0');
        dateEl.textContent = `${yyyy}.${mm}.${dd} (${day})`;
        timeEl.innerHTML = `${hh}<span class="colon">:</span>${mi}<span class="colon">:</span>${ss}`;

        // 미입금 카드의 시각 갱신 표시 (시계와 연동)
        const tick = document.getElementById('h-unpaid-tick');
        if (tick) tick.textContent = `${hh}:${mi}:${ss}`;
    };
    update();
    setInterval(update, 1000);

    // 미입금 카드 데이터 로드 (페이지 진입 시 + 60초마다 갱신)
    loadHumanUnpaidFromF7();
    setInterval(loadHumanUnpaidFromF7, 60000);

    // 매출액(2026) 카드 — f7 '휴먼' 시트에서 월별 자동 반영
    loadHumanSalesFromF7();
    setInterval(loadHumanSalesFromF7, 60000);

    // 업체별 청구 통계 카드 — f1 요약 시트에서 자동 반영
    loadHumanBillingFromF1();
    setInterval(loadHumanBillingFromF1, 60000);

    // 은행 계좌 현황 카드 — f7 '시재' 시트에서 자동 반영
    loadHumanBankFromF7();
    setInterval(loadHumanBankFromF7, 60000);
}

// 미입금(휴먼) 데이터를 f7 시트에서 추출하여 카드에 표시
async function loadHumanUnpaidFromF7() {
    const card = document.getElementById('h-card-unpaid');
    if (!card) return;

    let sheets = null;
    // 1) saved 우선
    try {
        const saved = JSON.parse(localStorage.getItem('data_human_f7'));
        if (saved && saved.sheets) sheets = saved.sheets;
    } catch (e) {}

    // 2) saved 없으면 원본 fetch
    if (!sheets || !sheets['미입금(휴먼)']) {
        try {
            const fileName = '2026.04-금전출납(세무용)_휴먼.xlsx';
            const res = await fetch(encodeURIComponent(fileName));
            if (res.ok) {
                const buf = await res.arrayBuffer();
                const wb = XLSX.read(new Uint8Array(buf), { type: 'array', dense: true });
                sheets = sheets || {};
                wb.SheetNames.forEach(s => {
                    const ws = wb.Sheets[s];
                    sheets[s] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                });
            }
        } catch (e) {}
    }

    if (!sheets || !sheets['미입금(휴먼)']) return;

    const sheet = sheets['미입금(휴먼)'];
    const unwrap = v => (v && typeof v === 'object' && v.isFormula) ? v.value : v;
    const toNum = v => {
        v = unwrap(v);
        if (v === null || v === undefined || v === '') return 0;
        return parseInt(v.toString().replace(/[^0-9-]/g, '')) || 0;
    };

    // 행2(엑셀 1-based)의 합계 SUM: D2=합계금액, E2=공급가액, F2=세액
    const totalsRow = sheet[1] || [];
    const totalAmount = toNum(totalsRow[3]);
    const totalSupply = toNum(totalsRow[4]);
    const totalTax = toNum(totalsRow[5]);

    // 데이터 행 추출 (행4 헤더 다음, B열이 시리얼 날짜인 행만)
    const items = [];
    for (let i = 4; i < sheet.length; i++) {
        const row = sheet[i] || [];
        let dateVal = unwrap(row[1]);
        if (typeof dateVal === 'string' && dateVal.trim() && !isNaN(dateVal)) dateVal = parseFloat(dateVal);
        if (typeof dateVal !== 'number' || dateVal < 40000 || dateVal > 80000) continue;
        const nameVal = unwrap(row[2]);
        const amountVal = toNum(row[3]);
        const date = new Date((dateVal - 25569) * 86400 * 1000);
        const dateStr = date.toISOString().split('T')[0];
        items.push({ date: dateStr, name: (nameVal || '').toString().trim(), amount: amountVal });
    }
    // 최근 일자가 위로
    items.sort((a, b) => b.date.localeCompare(a.date));

    // DOM 업데이트 — 합계는 마스킹 풀린 경우만 실제 값 표시
    const session = JSON.parse(localStorage.getItem('userSession') || 'null');
    const valEl = document.getElementById('h-val-unpaid');
    if (valEl && session) {
        valEl.classList.add('unmasked');
        valEl.textContent = totalAmount.toLocaleString() + '원';
    }
    const supplyEl = document.getElementById('h-unpaid-supply');
    const taxEl = document.getElementById('h-unpaid-tax');
    const countEl = document.getElementById('h-unpaid-count');
    const itemsEl = document.getElementById('h-unpaid-items');
    const breakdown = document.getElementById('h-unpaid-breakdown');
    if (supplyEl) supplyEl.textContent = totalSupply.toLocaleString() + '원';
    if (taxEl) taxEl.textContent = totalTax.toLocaleString() + '원';
    if (countEl) countEl.textContent = items.length.toLocaleString() + '건';
    if (itemsEl) {
        itemsEl.innerHTML = items.map(it =>
            `<tr><td>${it.date}</td><td>${escapeHtml(it.name)}</td><td>${it.amount.toLocaleString()}원</td></tr>`
        ).join('') || '<tr><td colspan="3" style="text-align:center;color:#999;padding:14px;">데이터 없음</td></tr>';
    }
    if (breakdown && session) breakdown.style.display = 'block';
}

// 휴먼 매출액(2026) — f7 '휴먼' 시트에서 작성일자/매출합계로 월별 합산
async function loadHumanSalesFromF7() {
    const card = document.getElementById('h-card-sales');
    if (!card) return;

    let sheets = null;
    try {
        const saved = JSON.parse(localStorage.getItem('data_human_f7'));
        if (saved && saved.sheets) sheets = saved.sheets;
    } catch (e) {}

    if (!sheets || !sheets['휴먼']) {
        try {
            const fileName = '2026.04-금전출납(세무용)_휴먼.xlsx';
            const res = await fetch(encodeURIComponent(fileName));
            if (res.ok) {
                const buf = await res.arrayBuffer();
                const wb = XLSX.read(new Uint8Array(buf), { type: 'array', dense: true });
                sheets = sheets || {};
                wb.SheetNames.forEach(s => {
                    const ws = wb.Sheets[s];
                    sheets[s] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                });
            }
        } catch (e) {}
    }

    if (!sheets || !sheets['휴먼']) return;

    const sheet = sheets['휴먼'];
    const unwrap = v => (v && typeof v === 'object' && v.isFormula) ? v.value : v;
    const toNum = v => {
        v = unwrap(v);
        if (v === null || v === undefined || v === '') return 0;
        if (typeof v === 'number') return v;
        return parseFloat(v.toString().replace(/[^0-9.\-]/g, '')) || 0;
    };
    const toDate = v => {
        v = unwrap(v);
        if (typeof v === 'string' && v.trim() && !isNaN(v)) v = parseFloat(v);
        if (typeof v !== 'number' || v < 40000 || v > 80000) return null;
        return new Date((v - 25569) * 86400 * 1000);
    };

    // 헤더 행 = row index 2 (1-based 3). 데이터는 index 3부터.
    const monthly = new Array(12).fill(0);
    let total = 0;
    for (let i = 3; i < sheet.length; i++) {
        const row = sheet[i] || [];
        const d = toDate(row[0]);
        if (!d || d.getFullYear() !== 2026) continue;
        const sales = toNum(row[2]); // 매출합계
        monthly[d.getMonth()] += sales;
        total += sales;
    }

    const fmtMM = v => {
        const mm = v / 1000000;
        return mm.toLocaleString('ko-KR', { maximumFractionDigits: 1 });
    };

    // 로그인 상태에서만 실제 값 노출 (다른 카드들과 동일한 마스킹 정책)
    const session = JSON.parse(localStorage.getItem('userSession') || 'null');
    if (!session) return;

    const valEl = document.getElementById('h-val-sales');
    if (valEl) {
        valEl.classList.add('unmasked');
        valEl.textContent = fmtMM(total) + ' 백만원';
    }

    const wrap = document.getElementById('h-sales-monthly-wrap');
    const tbody = document.getElementById('h-sales-monthly-body');
    if (tbody) {
        const rows = [
            `<tr class="total-row"><td>합계</td><td>${fmtMM(total)}</td></tr>`,
            ...monthly.map((v, i) => {
                const cls = v === 0 ? 'month-zero' : '';
                return `<tr class="${cls}"><td>${i + 1}월</td><td>${fmtMM(v)}</td></tr>`;
            })
        ];
        tbody.innerHTML = rows.join('');
    }
    if (wrap) wrap.style.display = 'block';
}

// 휴먼 업체별 청구 통계 — f1 요약 시트(예: '2026년 4월휴먼')에서 KM 업체별 통계 추출
async function loadHumanBillingFromF1() {
    const card = document.getElementById('h-card-purchase');
    if (!card) return;

    let sheets = null;
    try {
        const saved = JSON.parse(localStorage.getItem('data_human_f1'));
        if (saved && saved.sheets) sheets = saved.sheets;
    } catch (e) {}

    if (!sheets) {
        try {
            const fileName = '2026.04 업체별(청구용)_KM 휴먼.xlsx';
            const res = await fetch(encodeURIComponent(fileName));
            if (res.ok) {
                const buf = await res.arrayBuffer();
                const wb = XLSX.read(new Uint8Array(buf), { type: 'array', dense: true });
                sheets = {};
                wb.SheetNames.forEach(s => {
                    const ws = wb.Sheets[s];
                    sheets[s] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                });
            }
        } catch (e) {}
    }
    if (!sheets) return;

    // 요약 시트는 '2026년 N월휴먼' 패턴 — 월이 바뀌어도 자동 매칭
    const summaryName = Object.keys(sheets).find(n => /\d+월휴먼\s*$/.test(n))
                      || Object.keys(sheets).find(n => n.includes('휴먼'));
    if (!summaryName || !sheets[summaryName]) return;
    const sheet = sheets[summaryName];

    const unwrap = v => (v && typeof v === 'object' && v.isFormula) ? v.value : v;
    const toNum = v => {
        v = unwrap(v);
        if (v === null || v === undefined || v === '') return 0;
        if (typeof v === 'number') return v;
        return parseFloat(v.toString().replace(/[^0-9.\-]/g, '')) || 0;
    };
    const toStr = v => {
        v = unwrap(v);
        return (v === null || v === undefined) ? '' : v.toString().trim();
    };

    // 헤더는 row index 1: 주차|업체|지급액|청구액(세금계산서)|차액|결재|청구(기존)|청구액(당월)|미청구|전월이월
    // 데이터는 index 2부터, '계' 행 만나면 종료
    const vendors = [];
    const totals = { current: 0, cumulative: 0, unpaid: 0 };
    for (let i = 2; i < sheet.length; i++) {
        const row = sheet[i] || [];
        const colA = toStr(row[0]);
        const name = toStr(row[1]);
        if (colA === '계' || colA.startsWith('합계')) break;
        if (!name) continue;

        const cumulative = toNum(row[3]); // 청구액(세금계산서)
        const current = toNum(row[7]);     // 청구액(당월)
        const unpaid = toNum(row[9]);      // 전월이월 (= 미입금 carryover)
        const status = toStr(row[5]);      // 결재 (예: 한달(25일))

        // 빈 행/모두 0인 행은 스킵
        if (cumulative === 0 && current === 0 && unpaid === 0) continue;

        vendors.push({ name, current, cumulative, unpaid, status });
        totals.current += current;
        totals.cumulative += cumulative;
        totals.unpaid += unpaid;
    }
    if (vendors.length === 0) return;

    const tickEl = document.getElementById('h-billing-tick');
    if (tickEl) {
        const n = new Date();
        const hh = String(n.getHours()).padStart(2, '0');
        const mi = String(n.getMinutes()).padStart(2, '0');
        const ss = String(n.getSeconds()).padStart(2, '0');
        tickEl.textContent = `${hh}:${mi}:${ss}`;
    }

    const session = JSON.parse(localStorage.getItem('userSession') || 'null');
    const valEl = document.getElementById('h-val-purchase');
    if (valEl && session) {
        valEl.classList.add('unmasked');
        valEl.innerHTML = `<div class="stat-value" style="font-size: 1.8rem;">${totals.current.toLocaleString()}원</div>`;
    }

    if (!session) return;

    const cur = document.getElementById('h-billing-current');
    const cum = document.getElementById('h-billing-cumulative');
    const unp = document.getElementById('h-billing-unpaid');
    const body = document.getElementById('h-billing-body');
    const wrap = document.getElementById('h-billing-breakdown');
    if (cur) cur.textContent = totals.current.toLocaleString() + '원';
    if (cum) cum.textContent = totals.cumulative.toLocaleString() + '원';
    if (unp) unp.textContent = totals.unpaid.toLocaleString() + '원';
    if (body) {
        body.innerHTML = vendors.map(v => `
            <tr>
                <td>${escapeHtml(v.name)}</td>
                <td>${v.current.toLocaleString()}</td>
                <td>${v.cumulative.toLocaleString()}</td>
                <td class="${v.unpaid > 0 ? 'has-unpaid' : ''}">${v.unpaid.toLocaleString()}</td>
                <td>${escapeHtml(v.status)}</td>
            </tr>
        `).join('');
    }
    if (wrap) wrap.style.display = 'block';
}

// 휴먼 은행 계좌 현황 — f7 '시재' 시트 row 0의 (은행명, 잔액) 페어를 추출
async function loadHumanBankFromF7() {
    const card = document.getElementById('h-card-bank');
    if (!card) return;

    let sheets = null;
    try {
        const saved = JSON.parse(localStorage.getItem('data_human_f7'));
        if (saved && saved.sheets) sheets = saved.sheets;
    } catch (e) {}

    if (!sheets || !sheets['시재']) {
        try {
            const fileName = '2026.04-금전출납(세무용)_휴먼.xlsx';
            const res = await fetch(encodeURIComponent(fileName));
            if (res.ok) {
                const buf = await res.arrayBuffer();
                const wb = XLSX.read(new Uint8Array(buf), { type: 'array', dense: true });
                sheets = sheets || {};
                wb.SheetNames.forEach(s => {
                    const ws = wb.Sheets[s];
                    sheets[s] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                });
            }
        } catch (e) {}
    }
    if (!sheets || !sheets['시재']) return;

    const sheet = sheets['시재'];
    const unwrap = v => (v && typeof v === 'object' && v.isFormula) ? v.value : v;
    const toNum = v => {
        v = unwrap(v);
        if (v === null || v === undefined || v === '') return 0;
        if (typeof v === 'number') return v;
        return parseFloat(v.toString().replace(/[^0-9.\-]/g, '')) || 0;
    };
    const toStr = v => {
        v = unwrap(v);
        return (v === null || v === undefined) ? '' : v.toString().trim();
    };

    // Row 0 형식: 합계 | <total> | 기업 | <amt> | 하나 | <amt> | 우리 | <amt> | 신한 | <amt> | 농협 | <amt> | 국민 | <amt>
    const row0 = sheet[0] || [];
    let total = null;
    const banks = [];
    for (let i = 0; i < row0.length - 1; i += 2) {
        const label = toStr(row0[i]);
        if (!label) continue;
        const amount = toNum(row0[i + 1]);
        if (label === '합계') {
            total = amount;
        } else {
            banks.push({ name: label, balance: amount });
        }
    }
    if (banks.length === 0) return;

    // 합계 셀이 비어있으면 은행 잔액 직접 합산
    if (total === null || total === 0) {
        total = banks.reduce((s, b) => s + b.balance, 0);
    }

    const tickEl = document.getElementById('h-bank-tick');
    if (tickEl) {
        const n = new Date();
        const hh = String(n.getHours()).padStart(2, '0');
        const mi = String(n.getMinutes()).padStart(2, '0');
        const ss = String(n.getSeconds()).padStart(2, '0');
        tickEl.textContent = `${hh}:${mi}:${ss}`;
    }

    const session = JSON.parse(localStorage.getItem('userSession') || 'null');
    const valEl = document.getElementById('h-val-bank');
    if (valEl && session) {
        valEl.classList.add('unmasked');
        valEl.textContent = total.toLocaleString() + '원';
    }
    if (!session) return;

    const wrap = document.getElementById('h-bank-wrap');
    const tbody = document.getElementById('h-bank-body');
    if (tbody) {
        tbody.innerHTML = banks.map(b => {
            let cls = '';
            if (b.balance === 0) cls = 'zero-bal';
            else if (b.balance < 0) cls = 'neg-bal';
            return `
                <tr>
                    <td>${escapeHtml(b.name)}</td>
                    <td class="${cls}">${b.balance.toLocaleString()}</td>
                </tr>
            `;
        }).join('');
    }
    if (wrap) wrap.style.display = 'block';
}

function escapeHtml(s) {
    return (s || '').toString()
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function buildFormTiles(prefix) {
    const grid = document.getElementById(prefix + '-form-grid');
    if (!grid) return;
    const company = prefix === 'h' ? 'human' : 'chaeum';
    const forms = FORMS.filter(f => f.co === company);
    grid.innerHTML = forms.map(f => {
        const type = company + '_' + f.id;
        return `
            <div class="form-tile" id="tile-${type}">
                <span class="form-tile-status" id="status-${type}"></span>
                <div class="form-tile-icon">${f.icon}</div>
                <div class="form-tile-title">${f.title}</div>
                <div class="form-tile-meta" id="meta-${type}">데이터 없음</div>
                <div class="form-tile-buttons">
                    <button class="btn-sm btn-open" onclick="openSecureDetail('${type}')">📊 열기</button>
                    <label class="btn-sm btn-upload" for="upload-${type}">📁 업로드</label>
                    <input type="file" id="upload-${type}" accept=".xlsx,.xls" style="display:none" onchange="quickUpload(event,'${type}')">
                </div>
            </div>`;
    }).join('');
    forms.forEach(f => refreshFormTile(company + '_' + f.id));
}

function refreshFormTile(type) {
    const meta = document.getElementById('meta-' + type);
    const status = document.getElementById('status-' + type);
    if (!meta) return;
    try {
        const saved = JSON.parse(localStorage.getItem('data_' + type));
        if (saved && saved.data && saved.data.length > 0) {
            const rows = saved.data.length;
            const fileName = saved.fileName ? saved.fileName.replace(/\.(xlsx|xls)$/i, '') : '저장됨';
            const ts = saved.uploadedAt ? saved.uploadedAt.split(' ')[0] : '';
            meta.innerHTML = `<b>${rows}행</b> · ${ts || fileName}`;
            meta.classList.add('has-data');
            if (status) status.classList.add('uploaded');
        } else {
            meta.textContent = '원본 양식 사용';
            meta.classList.remove('has-data');
            if (status) status.classList.remove('uploaded');
        }
    } catch(e) {
        meta.textContent = '데이터 없음';
    }
}

function quickUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;
    if (!localStorage.getItem('userSession')) {
        alert('업로드는 로그인 후 가능합니다.');
        event.target.value = '';
        return;
    }
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', dense: true });
            const sheets = {};
            const sheetMerges = {};
            wb.SheetNames.forEach(s => {
                const ws = wb.Sheets[s];
                sheets[s] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                sheetMerges[s] = (ws['!merges'] || []).map(m => ({
                    s: { r: m.s.r, c: m.s.c },
                    e: { r: m.e.r, c: m.e.c }
                }));
            });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
            // 양식별 명시 헤더 행 우선
            const fmIdQ = (type.match(/_(f[1-7])$/) || [])[1] || '';
            const formMeta = FORMS.find(x => x.id === fmIdQ);
            let headerIdx = formMeta && formMeta.headerRow ? formMeta.headerRow - 1 : 0;
            // fallback: 명시 헤더 행이 비어있으면 상위 5행 중 4셀 이상 채워진 첫 행
            if (!json[headerIdx] || json[headerIdx].filter(c => c && c.toString().trim()).length < 3) {
                for (let i = 0; i < Math.min(5, json.length); i++) {
                    if (json[i] && json[i].filter(c => c && c.toString().trim()).length >= 4) { headerIdx = i; break; }
                }
            }
            const headers = json[headerIdx] || [];
            // 엑셀 원본 그대로: 빈 행도 보존 (사용자 요청)
            const data = json.slice(headerIdx + 1);
            const payload = {
                formatVersion: 4,
                headers,
                data,
                sheets,
                sheetMerges,
                fileName: file.name,
                uploadedAt: new Date().toLocaleString('ko-KR'),
                uploadedBy: (JSON.parse(localStorage.getItem('userSession')) || {}).name || '관리자'
            };
            localStorage.setItem('data_' + type, JSON.stringify(payload));
            refreshFormTile(type);
            recalcWorkerTotals();
            alert(`✅ "${file.name}" 업로드 완료\n${data.length}행이 저장되었습니다.`);
        } catch (err) {
            alert('파일 읽기 오류: ' + err.message);
        }
        event.target.value = '';
    };
    reader.readAsArrayBuffer(file);
}

function recalcWorkerTotals() {
    ['h', 'c'].forEach(prefix => {
        const company = prefix === 'h' ? 'human' : 'chaeum';
        const workerEl = document.getElementById(prefix + '-val-daily-worker');
        if (!workerEl) return;
        let sum = 0;
        FORMS.filter(f => f.co === company).forEach(f => {
            try {
                const saved = JSON.parse(localStorage.getItem('data_' + company + '_' + f.id));
                if (saved && saved.data) {
                    saved.data.forEach(row => {
                        const v = parseInt((row[f.payCol] || 0).toString().replace(/[^0-9-]/g, '')) || 0;
                        sum += v;
                    });
                }
            } catch(e) {}
        });
        if (sum > 0) workerEl.textContent = sum.toLocaleString() + '원';
    });
}

function initPurchaseChart() {
    // Only Human purchase chart for now as specified in original, 
    // but we could expand this later.
    const ctx = document.getElementById('purchaseBarChart');
    if (!ctx) return;
    // ... rest of logic for purchase chart if it exists in index.html
}

// --- AUTH LOGIC ---

function toggleModal(id) {
    const modal = document.getElementById(id);
    if (!modal) return;
    const isVisible = modal.style.display === 'block';
    
    // Close other modals
    document.querySelectorAll('.auth-modal').forEach(m => m.style.display = 'none');
    
    modal.style.display = isVisible ? 'none' : 'block';
}

function handleSignUp() {
    const name = document.getElementById('reg-name').value;
    const phone = document.getElementById('reg-phone').value;
    const address = document.getElementById('reg-address').value;
    const id = document.getElementById('reg-id').value;
    const pw = document.getElementById('reg-pw').value;
    const grade = document.getElementById('reg-grade').value;

    if (!name || !phone || !id || !pw) {
        alert('모든 필수 정보를 입력해주세요.');
        return;
    }

    let users = JSON.parse(localStorage.getItem('registered_users') || '[]');
    if (users.find(u => u.id === id)) {
        alert('이미 존재하는 아이디입니다.');
        return;
    }

    const newUser = {
        name, phone, address, id, pw, grade,
        approved: false,
        requestDate: new Date().toLocaleString()
    };

    users.push(newUser);
    localStorage.setItem('registered_users', JSON.stringify(users));
    
    alert('회원가입 신청이 완료되었습니다.\n관리자 승인 후 접속이 가능합니다.');
    toggleModal('signup-modal');
    refreshAdminPanel();
}

function handleSecureLogin() {
    const id = document.getElementById('login-id').value.trim();
    const pw = document.getElementById('login-pw').value.trim();

    if (id === 'admin' && pw === 'human7777!') {
        const masterSession = {
            id: 'admin', 
            grade: 'A', 
            name: '총괄 관리자',
            isMaster: true
        };
        localStorage.setItem('userSession', JSON.stringify(masterSession));
        alert('총괄 관리자 권한으로 인증되었습니다.');
        location.reload();
        return;
    }

    let users = JSON.parse(localStorage.getItem('registered_users') || '[]');
    const user = users.find(u => u.id === id && u.pw === pw);

    if (!user) {
        alert('아이디 또는 비밀번호가 일치하지 않습니다.');
        return;
    }

    if (!user.approved) {
        alert('아직 관리자 승인이 완료되지 않은 계정입니다.\n승인 대기 중입니다.');
        return;
    }

    const sessionData = {
        id: user.id, 
        grade: user.grade, 
        name: user.name,
        isMaster: user.grade === 'A'
    };
    localStorage.setItem('userSession', JSON.stringify(sessionData));
    applyPermissions(user.grade, user.name);
    toggleModal('login-modal');
    alert(user.name + '님, 환영합니다. (' + user.grade + '등급' + (user.grade === 'A' ? ' 마스터' : '') + ')');
    location.reload(); // Reload to ensure all components see the new session
}

function checkExistingSession() {
    const session = JSON.parse(localStorage.getItem('userSession'));
    if (session) {
        applyPermissions(session.grade, session.name);
        if (session.grade === 'A') {
            const adminPanel = document.getElementById('admin-panel');
            if (adminPanel) adminPanel.style.display = 'block';
            refreshAdminPanel();
        }
    }
}

function applyPermissions(grade, name) {
    const badge = document.getElementById('user-grade-badge');
    const loginBtn = document.getElementById('login-btn');
    const signupBtn = document.getElementById('signup-btn');
    
    if (badge) {
        badge.style.display = 'inline-block';
        badge.textContent = name + ' (' + grade + ')';
        badge.className = 'status-tag ' + (grade === 'A' ? 'tag-success' : 'tag-warning');
    }
    
    if (loginBtn) {
        loginBtn.textContent = '로그아웃';
        loginBtn.onclick = handleLogout;
    }
    
    if (signupBtn) signupBtn.style.display = 'none';

    // Reveal masked data for both companies
    ['h', 'c'].forEach(prefix => {
        const companyKey = prefix === 'h' ? 'human' : 'chaeum';
        
        // 1. 미입금
        const unpaidEl = document.getElementById(`${prefix}-val-unpaid`);
        if (unpaidEl) {
            unpaidEl.classList.add('unmasked');
            const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_unpaid`));
            if (saved && saved.length > 0) {
                const total = saved.reduce((acc, row) => {
                    const val = parseInt(row[1].replace(/[^0-9]/g, '')) || 0;
                    return acc + val;
                }, 0);
                unpaidEl.textContent = total.toLocaleString() + '원';
            } else {
                unpaidEl.textContent = prefix === 'h' ? '149,816,396원' : '85,420,000원';
            }
        }

        // 2. 매출액 — 휴먼은 f7(금전출납)에서 자동 반영(loadHumanSalesFromF7), 채움은 기존 로직 유지
        const salesEl = document.getElementById(`${prefix}-val-sales`);
        if (salesEl) {
            salesEl.classList.add('unmasked');
            if (prefix === 'h') {
                loadHumanSalesFromF7();
            } else {
                const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_sales`));
                if (saved && saved.data && saved.data.length > 0) {
                    const total = saved.data.reduce((acc, row) => {
                        let amountStr = (row[2] || '0').toString();
                        return acc + (parseInt(amountStr.replace(/[^0-9]/g, '')) || 0);
                    }, 0);
                    salesEl.textContent = total.toLocaleString() + '원';
                } else {
                    salesEl.textContent = '780,000,000원';
                }
            }
        }

        // 3. 매입계산서 — 휴먼은 f1 업체별 청구 통계로 대체(loadHumanBillingFromF1), 채움은 기존 로직 유지
        const purchaseEl = document.getElementById(`${prefix}-val-purchase`);
        if (purchaseEl) {
            purchaseEl.classList.add('unmasked');
            if (prefix === 'h') {
                loadHumanBillingFromF1();
            } else {
                const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_purchase`));
                if (saved && saved.data && saved.data.length > 0) {
                    const total = saved.data.reduce((acc, row) => acc + (parseInt(row[2].toString().replace(/[^0-9]/g, '')) || 0), 0);
                    purchaseEl.innerHTML = `<div class="stat-value" style="font-size: 1.8rem;">${total.toLocaleString()}원</div>`;
                } else {
                    purchaseEl.innerHTML = `<div class="stat-value" style="font-size: 1.8rem;">210,550,000원</div>`;
                }
            }
        }

        // 4. 현재 잔액 장부
        const ledgerEl = document.getElementById(`${prefix}-val-ledger`);
        if (ledgerEl) {
            ledgerEl.classList.add('unmasked');
            const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_ledger`));
            if (saved && saved.data && saved.data.length > 0) {
                const latestRow = saved.data[0];
                const balance = parseInt((latestRow[4] || 0).toString().replace(/[^0-9-]/g, '')) || 0;
                ledgerEl.textContent = balance.toLocaleString() + '원';
            } else {
                ledgerEl.textContent = prefix === 'h' ? '892,440,000원' : '450,220,000원';
            }
        }

        // 은행 계좌 현황 — 휴먼은 f7 '시재'에서 자동 반영(loadHumanBankFromF7), 채움은 기존 로직 유지
        const bankEl = document.getElementById(`${prefix}-val-bank`);
        if (bankEl) {
            bankEl.classList.add('unmasked');
            if (prefix === 'h') {
                loadHumanBankFromF7();
            } else {
                const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_bank`));
                if (saved && saved.data && saved.data.length > 0) {
                    const total = saved.data.reduce((acc, row) => acc + (parseInt(row[1].toString().replace(/[^0-9]/g, '')) || 0), 0);
                    bankEl.textContent = total.toLocaleString() + '원';
                } else {
                    bankEl.textContent = '125,440,000원';
                }
            }
        }

        // 5. 현재 잔액 (2026)
        const bal2026El = document.getElementById(`${prefix}-val-bal-2026`);
        if (bal2026El) {
            bal2026El.classList.add('unmasked');
            bal2026El.textContent = prefix === 'h' ? '312,550,000원' : '150,880,000원';
        }

        // 8. 사용내역
        const usageEl = document.getElementById(`${prefix}-val-usage`);
        if (usageEl) {
            usageEl.classList.add('unmasked');
            usageEl.textContent = prefix === 'h' ? '45,880,000원' : '22,440,000원';
        }

        // 11. 일용직 근무 내역 — 회사별 양식만 합산
        const workerEl = document.getElementById(`${prefix}-val-daily-worker`);
        if (workerEl) {
            workerEl.classList.add('unmasked');
            let sum = 0;
            FORMS.filter(f => f.co === companyKey).forEach(f => {
                try {
                    const saved = JSON.parse(localStorage.getItem('data_' + companyKey + '_' + f.id));
                    if (saved && saved.data) {
                        saved.data.forEach(row => {
                            const v = parseInt((row[f.payCol] || 0).toString().replace(/[^0-9-]/g, '')) || 0;
                            sum += v;
                        });
                    }
                } catch(e) {}
            });
            workerEl.textContent = sum > 0
                ? sum.toLocaleString() + '원'
                : (prefix === 'h' ? '12,500,000원' : '8,400,000원');
        }
    });

    // Update labels
    document.querySelectorAll('.stat-label').forEach(el => {
        if (el.textContent === '로그인 후 확인 가능') {
            el.textContent = '데이터 조회 및 수정 가능';
        }
    });

    // Handle visibility based on grade
    const gradeAOnlyIds = [
        'h-card-ledger', 'h-card-bal-2026', 'h-card-usage', 'h-card-bank', 'h-card-daily-worker',
        'c-card-ledger', 'c-card-bal-2026', 'c-card-usage', 'c-card-bank', 'c-card-daily-worker'
    ];

    if (grade === 'B') {
        gradeAOnlyIds.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = 'none';
        });
    } else {
        gradeAOnlyIds.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = 'block';
        });
    }
}

function handleLogout() {
    localStorage.removeItem('userSession');
    location.reload();
}

// --- ADMIN LOGIC ---

function refreshAdminPanel() {
    const tableBody = document.getElementById('pending-table-body');
    if (!tableBody) return;
    
    let users = JSON.parse(localStorage.getItem('registered_users') || '[]');
    const pending = users.filter(u => !u.approved);
    
    tableBody.innerHTML = '';
    
    if (pending.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:20px;">대기 중인 신청자가 없습니다.</td></tr>';
        return;
    }

    pending.forEach(u => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="padding:10px; border-bottom:1px solid #eee;">${u.name}</td>
            <td style="padding:10px; border-bottom:1px solid #eee;">${u.id}</td>
            <td style="padding:10px; border-bottom:1px solid #eee;">${u.phone}</td>
            <td style="padding:10px; border-bottom:1px solid #eee;">${u.grade}등급</td>
            <td style="padding:10px; border-bottom:1px solid #eee;">${u.requestDate || '-'}</td>
            <td style="padding:10px; border-bottom:1px solid #eee;">
                <button class="btn-sm" style="border-color:var(--success-color); color:var(--success-color);" onclick="approveUser('${u.id}')">승인</button>
                <button class="btn-sm" style="border-color:var(--danger-color); color:var(--danger-color); margin-left:5px;" onclick="rejectUser('${u.id}')">거절</button>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

function approveUser(userId) {
    let users = JSON.parse(localStorage.getItem('registered_users') || '[]');
    const user = users.find(u => u.id === userId);
    if (user) {
        user.approved = true;
        localStorage.setItem('registered_users', JSON.stringify(users));
        alert(userId + ' 계정이 승인되었습니다.');
        refreshAdminPanel();
    }
}

function rejectUser(userId) {
    if (!confirm('정말로 거절하시겠습니까?')) return;
    let users = JSON.parse(localStorage.getItem('registered_users') || '[]');
    users = users.filter(u => u.id !== userId);
    localStorage.setItem('registered_users', JSON.stringify(users));
    refreshAdminPanel();
}

function openSecureDetail(type) {
    if (!localStorage.getItem('userSession')) {
        alert('권한이 없습니다. 로그인이 필요합니다.');
        toggleModal('login-modal');
        return;
    }
    window.open(`editor.html?type=${type}`, '_blank', 'width=1100,height=700');
}

function simulateGlobalSync() {
    const loader = document.getElementById('loading-overlay');
    if (loader) {
        loader.style.display = 'flex';
        setTimeout(() => {
            loader.style.display = 'none';
            alert('동기화 완료');
        }, 1000);
    }
}
