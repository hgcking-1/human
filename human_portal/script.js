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

    // 근무자 통계 카드 — f2 일자별(지급용) 시트에서 자동 반영
    loadHumanWorkerFromF2();
    setInterval(loadHumanWorkerFromF2, 60000);

    // 분기별 부가세 카드 — f7 '휴먼' 시트에서 직접 계산
    loadHumanVATFromF7();
    setInterval(loadHumanVATFromF7, 60000);
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

// 휴먼 분기별 부가세 — f7 '휴먼' 시트의 매출/매입 공급가액을 연·분기로 그룹화하여 직접 계산
let _humanVATData = null; // { byYear: { 2026: { Q1: {...}, ..., total: {...} }, ... } }

async function loadHumanVATFromF7() {
    const card = document.getElementById('h-card-bal-2026');
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

    // 헤더 row 2 (1-based 3). 데이터는 index 3부터.
    // col: 0=작성일자 | 1=상호 | 2=매출합계 | 3=매출공급가액 | 4=매출세액 | 5=매입합계 | 6=매입공급가액 | 7=매입세액
    const byYear = {}; // year -> { Q1, Q2, Q3, Q4, total }
    const ensureYear = y => {
        if (!byYear[y]) {
            byYear[y] = {
                Q1: { sales: 0, purchase: 0, vat: 0, count: 0 },
                Q2: { sales: 0, purchase: 0, vat: 0, count: 0 },
                Q3: { sales: 0, purchase: 0, vat: 0, count: 0 },
                Q4: { sales: 0, purchase: 0, vat: 0, count: 0 },
                total: { sales: 0, purchase: 0, vat: 0, count: 0 }
            };
        }
        return byYear[y];
    };

    for (let i = 3; i < sheet.length; i++) {
        const row = sheet[i] || [];
        let serial = unwrap(row[0]);
        if (typeof serial === 'string' && serial.trim() && !isNaN(serial)) serial = parseFloat(serial);
        if (typeof serial !== 'number' || serial < 30000 || serial > 80000) continue;

        const date = new Date((serial - 25569) * 86400 * 1000);
        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const q = 'Q' + (Math.floor((month - 1) / 3) + 1);

        const salesSupply = toNum(row[3]);
        const purchaseSupply = toNum(row[6]);
        if (salesSupply === 0 && purchaseSupply === 0) continue;

        const yObj = ensureYear(year);
        yObj[q].sales += salesSupply;
        yObj[q].purchase += purchaseSupply;
        yObj[q].count++;
        yObj.total.sales += salesSupply;
        yObj.total.purchase += purchaseSupply;
        yObj.total.count++;
    }

    // 분기/연 부가세 = (sales - purchase) × 10%
    Object.keys(byYear).forEach(y => {
        ['Q1', 'Q2', 'Q3', 'Q4', 'total'].forEach(k => {
            const x = byYear[y][k];
            x.vat = Math.round((x.sales - x.purchase) * 0.1);
        });
    });

    _humanVATData = { byYear };

    // 시계
    const tickEl = document.getElementById('h-vat-tick');
    if (tickEl) {
        const n = new Date();
        tickEl.textContent = `${String(n.getHours()).padStart(2, '0')}:${String(n.getMinutes()).padStart(2, '0')}:${String(n.getSeconds()).padStart(2, '0')}`;
    }

    // 카드 상단 합계 — 최신 연도(가장 큰 연도) 부가세 합계
    const years = Object.keys(byYear).map(Number).sort((a, b) => b - a);
    const latestYear = years[0];
    const session = JSON.parse(localStorage.getItem('userSession') || 'null');
    const valEl = document.getElementById('h-val-bal-2026');
    if (valEl && session && latestYear) {
        valEl.classList.add('unmasked');
        const latestVAT = byYear[latestYear].total.vat;
        const sign = latestVAT < 0 ? '환급 ' : '';
        valEl.textContent = `${latestYear}년 ${sign}${Math.abs(latestVAT).toLocaleString()}원`;
    }
    if (!session) return;

    // 연도 select 채우기 (현재 선택값 유지)
    const sel = document.getElementById('h-vat-year');
    if (sel) {
        const prev = sel.value;
        sel.innerHTML = years.map(y => `<option value="${y}">${y}년</option>`).join('');
        if (prev && years.includes(parseInt(prev))) {
            sel.value = prev;
        } else {
            sel.value = String(latestYear);
        }
    }
    renderVATTable();

    const wrap = document.getElementById('h-vat-wrap');
    if (wrap) wrap.style.display = 'block';
}

function renderVATTable() {
    const body = document.getElementById('h-vat-body');
    const sel = document.getElementById('h-vat-year');
    if (!body || !sel || !_humanVATData) return;
    const year = parseInt(sel.value);
    const yObj = _humanVATData.byYear[year];
    if (!yObj) {
        body.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;padding:14px;">데이터 없음</td></tr>';
        return;
    }
    const fmt = v => v.toLocaleString();
    const rangeMap = { Q1: '1~3월', Q2: '4~6월', Q3: '7~9월', Q4: '10~12월' };
    const rows = ['Q1', 'Q2', 'Q3', 'Q4'].map(q => {
        const x = yObj[q];
        const empty = x.count === 0;
        const cls = empty ? 'empty-q' : '';
        const vatCls = x.vat > 0 ? 'pos' : (x.vat < 0 ? 'neg' : 'zero');
        return `
            <tr class="${cls}">
                <td>${q} (${rangeMap[q]})</td>
                <td>${empty ? '—' : fmt(x.sales)}</td>
                <td>${empty ? '—' : fmt(x.purchase)}</td>
                <td class="${vatCls}">${empty ? '—' : (x.vat >= 0 ? '+' : '') + fmt(x.vat)}</td>
            </tr>
        `;
    });
    const t = yObj.total;
    const tCls = t.vat > 0 ? 'pos' : (t.vat < 0 ? 'neg' : 'zero');
    rows.push(`
        <tr class="year-total">
            <td>${year} 연계</td>
            <td>${fmt(t.sales)}</td>
            <td>${fmt(t.purchase)}</td>
            <td class="${tCls}">${(t.vat >= 0 ? '+' : '') + fmt(t.vat)}</td>
        </tr>
    `);
    body.innerHTML = rows.join('');
}

// 휴먼 근무자 통계 (h-card-ledger) — f2 일자별(지급용)·KM 휴먼에서 추출
let _humanWorkerData = null; // { dailyTotals, workers[], byVendor, byKey }
let _humanWorkerSelectedKey = null;

async function loadHumanWorkerFromF2() {
    const card = document.getElementById('h-card-ledger');
    if (!card) return;

    let sheets = null;
    try {
        const saved = JSON.parse(localStorage.getItem('data_human_f2'));
        if (saved && saved.sheets) sheets = saved.sheets;
    } catch (e) {}

    if (!sheets) {
        try {
            const fileName = '2026.04 일자별(지급용)_KM 휴먼.xlsx';
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
    const serialToDay = serial => {
        if (typeof serial !== 'number' || serial < 40000 || serial > 80000) return null;
        const d = new Date((serial - 25569) * 86400 * 1000);
        return { date: d, day: d.getDate(), month: d.getMonth() + 1, year: d.getFullYear() };
    };

    // 1) 일별 합계 (요약 시트)
    const summaryName = Object.keys(sheets).find(n => /\d+월휴먼\s*$/.test(n));
    const dailyTotals = [];
    if (summaryName && sheets[summaryName]) {
        const ss = sheets[summaryName];
        for (let i = 3; i < ss.length; i++) {
            const r = ss[i] || [];
            const dayLabel = toStr(r[0]); // "1일", "2일" ...
            const m = dayLabel.match(/^(\d+)일/);
            if (!m) continue;
            const day = parseInt(m[1]);
            const count = toNum(r[1]);
            const pay = toNum(r[2]);
            const charge = toNum(r[3]);
            if (count === 0 && pay === 0) continue;
            dailyTotals.push({ day, count, pay, charge });
        }
    }

    // 2) 일별 시트 (1, 2, ...) → 근무자 데이터 누적
    const workerMap = {}; // key → worker
    Object.keys(sheets).forEach(sheetName => {
        if (!/^\d+$/.test(sheetName)) return;
        const sheet = sheets[sheetName];
        if (!sheet || sheet.length < 4) return;

        // row 0의 col 0이 시리얼 날짜
        const serial = toNum(sheet[0] && sheet[0][0]);
        const dInfo = serialToDay(serial);
        if (!dInfo) return;
        const dateStr = `${dInfo.month}/${dInfo.day}`;

        // 데이터 row 4부터
        for (let i = 4; i < sheet.length; i++) {
            const r = sheet[i] || [];
            const name = toStr(r[5]);
            if (!name) continue;
            const phone = toStr(r[6]);
            const jumin = toStr(r[7]);
            const site = toStr(r[2]);
            const workTime = toStr(r[4]);
            const pay = toNum(r[10]);
            const charge = toNum(r[11]);
            const paid = toStr(r[12]) === '지급완료';

            const key = jumin || `${name}|${phone}`;
            if (!workerMap[key]) {
                workerMap[key] = { key, name, phone, jumin, days: [] };
            }
            // 같은 사람이 같은 날 여러 행이면 모두 누적
            workerMap[key].days.push({
                day: dInfo.day, month: dInfo.month, year: dInfo.year,
                dateStr, site, workTime, pay, charge, paid
            });
        }
    });

    // 3) 근무자별 집계
    const workers = Object.values(workerMap).map(w => {
        // 같은 사람이 같은 날 여러 현장 근무한 경우 day 단위 집계
        const dayPayMap = {}; // day → totalPay
        const daySiteSet = {}; // day → Set<site>
        w.days.forEach(d => {
            dayPayMap[d.day] = (dayPayMap[d.day] || 0) + d.pay;
            if (!daySiteSet[d.day]) daySiteSet[d.day] = new Set();
            daySiteSet[d.day].add(d.site);
        });
        const uniqueDays = Object.keys(dayPayMap).map(Number).sort((a, b) => a - b);
        const totalDays = uniqueDays.length;
        const totalPay = w.days.reduce((s, d) => s + d.pay, 0);
        const totalCharge = w.days.reduce((s, d) => s + d.charge, 0);

        // 최대 연속 근무 계산
        let maxLen = 0, maxStart = 0, maxEnd = 0;
        if (uniqueDays.length > 0) {
            let curStart = uniqueDays[0], curEnd = uniqueDays[0], curLen = 1;
            maxLen = 1; maxStart = curStart; maxEnd = curEnd;
            for (let i = 1; i < uniqueDays.length; i++) {
                if (uniqueDays[i] === curEnd + 1) {
                    curEnd = uniqueDays[i]; curLen++;
                } else {
                    curStart = uniqueDays[i]; curEnd = uniqueDays[i]; curLen = 1;
                }
                if (curLen > maxLen) { maxLen = curLen; maxStart = curStart; maxEnd = curEnd; }
            }
        }

        // 현장별 집계
        const siteCount = {}; // site → days count
        const sitePay = {};
        Object.keys(daySiteSet).forEach(day => {
            daySiteSet[day].forEach(site => {
                siteCount[site] = (siteCount[site] || 0) + 1;
            });
        });
        w.days.forEach(d => {
            sitePay[d.site] = (sitePay[d.site] || 0) + d.pay;
        });
        const sites = Object.keys(siteCount).map(s => ({ site: s, days: siteCount[s], pay: sitePay[s] || 0 }))
            .sort((a, b) => b.days - a.days);
        const primarySite = sites[0] ? sites[0].site : '';
        const latestDay = uniqueDays.length ? uniqueDays[uniqueDays.length - 1] : 0;

        return {
            key: w.key, name: w.name, phone: w.phone, jumin: w.jumin,
            days: w.days, uniqueDays, totalDays, totalPay, totalCharge,
            maxConsecutive: { len: maxLen, start: maxStart, end: maxEnd },
            sites, primarySite, latestDay
        };
    });
    workers.sort((a, b) => b.totalDays - a.totalDays || b.totalPay - a.totalPay);

    // 4) 현장별 그룹화
    const byVendor = {}; // siteName → workers (with site-specific stats)
    workers.forEach(w => {
        w.sites.forEach(s => {
            if (!byVendor[s.site]) byVendor[s.site] = [];
            byVendor[s.site].push({ ...w, _vendorDays: s.days, _vendorPay: s.pay });
        });
    });
    Object.keys(byVendor).forEach(v => {
        byVendor[v].sort((a, b) => b._vendorDays - a._vendorDays || b._vendorPay - a._vendorPay);
    });

    const byKey = {};
    workers.forEach(w => { byKey[w.key] = w; });

    _humanWorkerData = { dailyTotals, workers, byVendor, byKey };

    // 시계
    const tickEl = document.getElementById('h-worker-tick');
    if (tickEl) {
        const n = new Date();
        tickEl.textContent = `${String(n.getHours()).padStart(2, '0')}:${String(n.getMinutes()).padStart(2, '0')}:${String(n.getSeconds()).padStart(2, '0')}`;
    }

    // 카드 상단 합계 — 4월 총 지급액
    const totalPay = dailyTotals.reduce((s, d) => s + d.pay, 0);
    const session = JSON.parse(localStorage.getItem('userSession') || 'null');
    const valEl = document.getElementById('h-val-ledger');
    if (valEl && session) {
        valEl.classList.add('unmasked');
        valEl.textContent = `${totalPay.toLocaleString()}원 (월 누계 지급)`;
    }
    if (!session) return;

    renderDailyStrip();
    populateVendorSelect();
    renderWorkerVendorTable();
    renderWorkerPersonList();

    const wrap = document.getElementById('h-worker-wrap');
    if (wrap) wrap.style.display = 'block';
}

function renderDailyStrip() {
    const el = document.getElementById('h-daily-strip');
    if (!el || !_humanWorkerData) return;
    const days = _humanWorkerData.dailyTotals;
    const fmtMan = v => {
        // 만원 단위로 압축 (예: 2178750 → 218만)
        if (v >= 10000) return Math.round(v / 10000).toLocaleString() + '만';
        return v.toLocaleString();
    };
    el.innerHTML = days.map(d => `
        <div class="daily-cell" title="${d.day}일 · ${d.count}명 · 지급 ${d.pay.toLocaleString()}원 · 청구 ${d.charge.toLocaleString()}원">
            <div class="dc-day">${d.day}일</div>
            <div class="dc-pay">${fmtMan(d.pay)}</div>
            <div class="dc-cnt">👥 ${d.count}명</div>
        </div>
    `).join('') || '<div style="padding:8px;color:#999;font-size:0.78rem;">일별 데이터 없음</div>';
}

function populateVendorSelect() {
    const sel = document.getElementById('h-worker-vendor-select');
    if (!sel || !_humanWorkerData) return;
    const vendors = Object.keys(_humanWorkerData.byVendor).sort();
    const totalCount = _humanWorkerData.workers.length;
    sel.innerHTML = `<option value="__ALL__">전체 (${totalCount}명)</option>` +
        vendors.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)} (${_humanWorkerData.byVendor[v].length}명)</option>`).join('');
}

function renderWorkerVendorTable() {
    const body = document.getElementById('h-worker-vendor-body');
    const sumEl = document.getElementById('h-worker-vendor-summary');
    const sel = document.getElementById('h-worker-vendor-select');
    if (!body || !_humanWorkerData) return;
    const vendor = sel ? sel.value : '__ALL__';

    let list, vendorDaysFn, vendorPayFn;
    if (vendor === '__ALL__') {
        list = _humanWorkerData.workers;
        vendorDaysFn = w => w.totalDays;
        vendorPayFn = w => w.totalPay;
    } else {
        list = _humanWorkerData.byVendor[vendor] || [];
        vendorDaysFn = w => w._vendorDays;
        vendorPayFn = w => w._vendorPay;
    }

    const totalPay = list.reduce((s, w) => s + vendorPayFn(w), 0);
    if (sumEl) sumEl.innerHTML = `총 <b>${list.length}</b>명 · 지급 <b>${totalPay.toLocaleString()}원</b>`;

    body.innerHTML = list.map(w => `
        <tr data-key="${escapeHtml(w.key)}" onclick="selectWorker('${escapeHtml(w.key)}')">
            <td>${escapeHtml(w.name)}</td>
            <td class="center">${escapeHtml(w.phone)}</td>
            <td class="center" style="font-size:0.7rem;color:#6b7280;">${escapeHtml(w.jumin)}</td>
            <td class="num">${vendorDaysFn(w)}일</td>
            <td class="num">${w.maxConsecutive.len}일${w.maxConsecutive.len > 1 ? ` (${w.maxConsecutive.start}~${w.maxConsecutive.end})` : ''}</td>
            <td class="num">${vendorPayFn(w).toLocaleString()}</td>
            <td class="center">${w.latestDay}일</td>
        </tr>
    `).join('') || '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px;">데이터 없음</td></tr>';
}

function renderWorkerPersonList() {
    const list = document.getElementById('h-worker-person-list');
    const sumEl = document.getElementById('h-worker-person-summary');
    const search = document.getElementById('h-worker-search');
    if (!list || !_humanWorkerData) return;
    const q = (search ? search.value : '').trim().toLowerCase();

    let filtered = _humanWorkerData.workers;
    if (q) {
        filtered = filtered.filter(w =>
            w.name.toLowerCase().includes(q) ||
            (w.phone || '').toLowerCase().includes(q) ||
            (w.jumin || '').includes(q)
        );
    }

    if (sumEl) sumEl.innerHTML = `검색 결과 <b>${filtered.length}</b>명 / 전체 <b>${_humanWorkerData.workers.length}</b>명`;

    list.innerHTML = filtered.slice(0, 200).map(w => `
        <tr data-key="${escapeHtml(w.key)}" class="${_humanWorkerSelectedKey === w.key ? 'selected' : ''}" onclick="selectWorker('${escapeHtml(w.key)}')">
            <td>${escapeHtml(w.name)}</td>
            <td class="center" style="font-size:0.7rem;color:#6b7280;">${escapeHtml(w.primarySite)}</td>
            <td class="num">${w.totalDays}일</td>
            <td class="num">${w.totalPay.toLocaleString()}</td>
        </tr>
    `).join('') || '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px;">검색 결과 없음</td></tr>';
}

function selectWorker(key) {
    if (!_humanWorkerData) return;
    _humanWorkerSelectedKey = key;
    const w = _humanWorkerData.byKey[key];
    if (!w) return;

    // 개인 탭으로 자동 전환
    switchWorkerTab('person');
    renderWorkerPersonList();

    const detail = document.getElementById('h-worker-detail');
    if (!detail) return;

    const sortedDays = [...w.uniqueDays].sort((a, b) => a - b);
    const inStreak = new Set();
    for (let d = w.maxConsecutive.start; d <= w.maxConsecutive.end; d++) inStreak.add(d);
    const dayPills = sortedDays.map(d =>
        `<span class="worker-detail-day-pill ${inStreak.has(d) ? 'streak' : ''}">${d}일</span>`
    ).join(' ');
    const sitesHtml = w.sites.map(s =>
        `<div>• <b>${escapeHtml(s.site)}</b>: ${s.days}일 · ${s.pay.toLocaleString()}원</div>`
    ).join('');

    detail.innerHTML = `
        <h4>${escapeHtml(w.name)}<small>${escapeHtml(w.phone)} · ${escapeHtml(w.jumin)}</small></h4>
        <div class="worker-detail-stats">
            <div class="worker-detail-stat"><div class="label">총 근무일수</div><div class="value">${w.totalDays}일</div></div>
            <div class="worker-detail-stat highlight"><div class="label">최대 연속 근무</div><div class="value">${w.maxConsecutive.len}일${w.maxConsecutive.len > 1 ? ` (${w.maxConsecutive.start}~${w.maxConsecutive.end})` : ''}</div></div>
            <div class="worker-detail-stat"><div class="label">총 지급액</div><div class="value">${w.totalPay.toLocaleString()}원</div></div>
            <div class="worker-detail-stat"><div class="label">평균 일급</div><div class="value">${w.totalDays > 0 ? Math.round(w.totalPay / w.totalDays).toLocaleString() : 0}원</div></div>
        </div>
        <div class="worker-detail-section-title">📍 근무 현장 (현장별 일수·지급액)</div>
        <div class="worker-detail-sites">${sitesHtml}</div>
        <div class="worker-detail-section-title">📅 근무 일자 (최대 연속 ${w.maxConsecutive.len}일은 노란색)</div>
        <div class="worker-detail-days">${dayPills}</div>
    `;
}

function switchWorkerTab(tab) {
    document.querySelectorAll('.worker-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tab));
    const v = document.getElementById('h-worker-vendor');
    const p = document.getElementById('h-worker-person');
    if (v) v.style.display = tab === 'vendor' ? 'block' : 'none';
    if (p) p.style.display = tab === 'person' ? 'block' : 'none';
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

        // 4. 현재 잔액 장부 — 휴먼은 근무자 통계로 대체(loadHumanWorkerFromF2), 채움은 기존 로직 유지
        const ledgerEl = document.getElementById(`${prefix}-val-ledger`);
        if (ledgerEl) {
            ledgerEl.classList.add('unmasked');
            if (prefix === 'h') {
                loadHumanWorkerFromF2();
            } else {
                const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_ledger`));
                if (saved && saved.data && saved.data.length > 0) {
                    const latestRow = saved.data[0];
                    const balance = parseInt((latestRow[4] || 0).toString().replace(/[^0-9-]/g, '')) || 0;
                    ledgerEl.textContent = balance.toLocaleString() + '원';
                } else {
                    ledgerEl.textContent = '450,220,000원';
                }
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

        // 5. 분기별 부가세 (휴먼) / 현재 잔액 (채움)
        const bal2026El = document.getElementById(`${prefix}-val-bal-2026`);
        if (bal2026El) {
            bal2026El.classList.add('unmasked');
            if (prefix === 'h') {
                loadHumanVATFromF7();
            } else {
                bal2026El.textContent = '150,880,000원';
            }
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
