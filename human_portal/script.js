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

// 업체별 지급 패턴 통계 (f7 휴먼 시트 기반) — 클릭 모달용
let _humanVendorStats = null; // { all: [], byName: {}, byPattern: {} }

function computeHumanVendorStats(humanSheet) {
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

    // 컬럼: 0=작성일자 1=상호 2=매출합계 3=매출공급가 4=매출세액 9=입금일 10=출금일 11=통장
    const txnsByVendor = {};
    for (let i = 3; i < humanSheet.length; i++) {
        const r = humanSheet[i] || [];
        const created = unwrap(r[0]);
        const vendor = toStr(r[1]);
        const sales = toNum(r[2]);
        const ipgum = unwrap(r[9]);
        const memo = toStr(r[8]);
        if (!vendor || sales <= 0) continue;
        if (typeof created !== 'number' || created < 30000 || created > 80000) continue;
        const ipgumNum = (typeof ipgum === 'number' && ipgum > 30000 && ipgum < 80000) ? ipgum : null;
        if (!txnsByVendor[vendor]) txnsByVendor[vendor] = [];
        txnsByVendor[vendor].push({
            created, sales, memo,
            ipgum: ipgumNum,
            gap: ipgumNum !== null ? (ipgumNum - created) : null,
            paid: ipgumNum !== null
        });
    }

    const all = Object.entries(txnsByVendor).map(([name, txns]) => {
        const gaps = txns.filter(t => t.gap !== null && t.gap >= -10 && t.gap < 400).map(t => t.gap);
        const totalAmount = txns.reduce((s, t) => s + t.sales, 0);
        const paidAmount = txns.filter(t => t.paid).reduce((s, t) => s + t.sales, 0);
        const unpaidAmount = totalAmount - paidAmount;
        gaps.sort((a, b) => a - b);
        const median = gaps.length ? gaps[Math.floor(gaps.length / 2)] : null;
        const avg = gaps.length ? gaps.reduce((s, x) => s + x, 0) / gaps.length : null;
        const min = gaps.length ? gaps[0] : null;
        const max = gaps.length ? gaps[gaps.length - 1] : null;

        // 갭 버킷 (당일=0~2 / 단기=3~7 / 10일=8~12 / 보름=13~17 / 장기=18~30 / 매우장기>30)
        const buckets = { sameDay: 0, short: 0, tenDay: 0, biweek: 0, longTerm: 0, veryLong: 0 };
        gaps.forEach(g => {
            if (g <= 2) buckets.sameDay++;
            else if (g <= 7) buckets.short++;
            else if (g <= 12) buckets.tenDay++;
            else if (g <= 17) buckets.biweek++;
            else if (g <= 30) buckets.longTerm++;
            else buckets.veryLong++;
        });

        // 패턴 분류
        let pattern = '데이터부족';
        if (gaps.length >= 3) {
            const n = gaps.length;
            const sameDayPct = buckets.sameDay / n;
            const tenDayPct = buckets.tenDay / n;
            if (sameDayPct >= 0.5) pattern = '당일지급';
            else if (tenDayPct >= 0.5) pattern = '10일지급';
            else if (sameDayPct >= 0.2 && tenDayPct >= 0.2) pattern = '혼합';
            else if (median > 17) pattern = '장기지급';
            else pattern = '기타';
        } else if (gaps.length > 0) {
            if (median <= 2) pattern = '당일지급';
            else if (median >= 8 && median <= 12) pattern = '10일지급';
            else pattern = '기타';
        }

        return {
            name, txns, totalAmount, paidAmount, unpaidAmount,
            count: txns.length, paidCount: txns.filter(t => t.paid).length,
            gaps, median, avg, min, max, buckets, pattern
        };
    }).filter(v => v.count >= 2);

    all.sort((a, b) => b.totalAmount - a.totalAmount);

    const byName = {};
    all.forEach(v => byName[v.name] = v);

    const patternOrder = ['당일지급', '10일지급', '혼합', '장기지급', '기타', '데이터부족'];
    const byPattern = {};
    patternOrder.forEach(p => byPattern[p] = []);
    all.forEach(v => { if (byPattern[v.pattern]) byPattern[v.pattern].push(v); });

    _humanVendorStats = { all, byName, byPattern, patternOrder };
}

// f1 KM 업체명을 f7 매출 거래처명에 매핑 (KM-* → (주)케이엠파크)
function findVendorStats(vendorName) {
    if (!_humanVendorStats) return null;
    const direct = _humanVendorStats.byName[vendorName];
    if (direct) return direct;
    if (vendorName.startsWith('KM-')) return _humanVendorStats.byName['(주)케이엠파크'] || null;
    return null;
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

    // f7 휴먼 시트 로드 → 업체별 지급 패턴 계산
    let f7sheets = null;
    try {
        const saved = JSON.parse(localStorage.getItem('data_human_f7'));
        if (saved && saved.sheets) f7sheets = saved.sheets;
    } catch (e) {}
    if (!f7sheets || !f7sheets['휴먼']) {
        try {
            const fileName = '2026.04-금전출납(세무용)_휴먼.xlsx';
            const res = await fetch(encodeURIComponent(fileName));
            if (res.ok) {
                const buf = await res.arrayBuffer();
                const wb = XLSX.read(new Uint8Array(buf), { type: 'array', dense: true });
                f7sheets = {};
                wb.SheetNames.forEach(s => {
                    f7sheets[s] = XLSX.utils.sheet_to_json(wb.Sheets[s], { header: 1 });
                });
            }
        } catch (e) {}
    }
    if (f7sheets && f7sheets['휴먼']) {
        computeHumanVendorStats(f7sheets['휴먼']);
    }

    if (body) {
        body.innerHTML = vendors.map(v => {
            const stats = findVendorStats(v.name);
            const pattern = stats ? stats.pattern : '데이터없음';
            const cls = patternToClass(pattern);
            const tooltip = stats
                ? `중앙값 ${stats.median ?? '—'}일 / 평균 ${stats.avg ? stats.avg.toFixed(1) : '—'}일 / n=${stats.count}건`
                : '거래 이력 매칭 없음';
            return `
                <tr class="vendor-row" onclick="openVendorModal('${escapeHtml(v.name)}')" title="클릭 시 상세 (입금일 갭·청구내역)">
                    <td>${escapeHtml(v.name)}</td>
                    <td>${v.current.toLocaleString()}</td>
                    <td>${v.cumulative.toLocaleString()}</td>
                    <td class="${v.unpaid > 0 ? 'has-unpaid' : ''}">${v.unpaid.toLocaleString()}</td>
                    <td><span class="pattern-badge ${cls}" title="${tooltip}">${pattern}</span></td>
                </tr>
            `;
        }).join('');
    }
    if (wrap) wrap.style.display = 'block';

    // 분류 보기 버튼 표시 여부
    const catBtn = document.getElementById('h-billing-pattern-btn');
    if (catBtn) catBtn.style.display = _humanVendorStats ? 'inline-block' : 'none';
}

function patternToClass(p) {
    return ({
        '당일지급': 'pat-same',
        '10일지급': 'pat-ten',
        '혼합': 'pat-mix',
        '장기지급': 'pat-long',
        '기타': 'pat-etc',
        '데이터부족': 'pat-na',
        '데이터없음': 'pat-na'
    })[p] || 'pat-na';
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

    // 거래 데이터 파싱 (시재 시트 row 4부터: 계좌·날짜·적요·입금·출금)
    const txns = [];
    let minSerial = Infinity, maxSerial = -Infinity;
    for (let i = 4; i < sheet.length; i++) {
        const r = sheet[i] || [];
        const bankName = toStr(r[0]);
        if (!bankName) continue;
        let serial = unwrap(r[1]);
        if (typeof serial === 'string' && serial.trim() && !isNaN(serial)) serial = parseFloat(serial);
        if (typeof serial !== 'number' || serial < 30000 || serial > 80000) continue;
        const dep = toNum(r[3]);
        const wd = toNum(r[4]);
        if (dep === 0 && wd === 0) continue;
        if (serial < minSerial) minSerial = serial;
        if (serial > maxSerial) maxSerial = serial;
        txns.push({ bank: bankName, serial, desc: toStr(r[2]), dep, wd });
    }

    // 잔액에 있지만 거래엔 없는 은행도 포함하도록 통합 은행 목록 구성
    const bankNames = banks.map(b => b.name);
    txns.forEach(t => { if (!bankNames.includes(t.bank)) bankNames.push(t.bank); });

    // 주차 정의 (가장 큰 거래일이 속한 달의 1~말일을 7일 단위로 분할)
    let weeks = [];
    if (isFinite(maxSerial)) {
        const maxDate = new Date((maxSerial - 25569) * 86400 * 1000);
        const year = maxDate.getFullYear();
        const month = maxDate.getMonth();
        const monthEnd = new Date(year, month + 1, 0).getDate();
        const monthStartSerial = Math.round(new Date(year, month, 1).getTime() / 86400000) + 25569;
        for (let s = 1; s <= monthEnd; s += 7) {
            const e = Math.min(s + 6, monthEnd);
            weeks.push({ label: `W${weeks.length + 1}`, range: `${s}~${e}일`, startSerial: monthStartSerial + (s - 1), endSerial: monthStartSerial + (e - 1) });
        }
    }

    _humanBankData = { banks, bankNames, txns, weeks, total };

    renderHumanBankPills();
}

// 은행별 데이터 캐시 (hover로 펼쳐지는 detail 영역용)
let _humanBankData = null;
let _humanBankActive = null;

function renderHumanBankPills() {
    const pills = document.getElementById('h-bank-pills');
    const wrap = document.getElementById('h-bank-wrap');
    if (!pills || !_humanBankData) return;

    const { banks, total } = _humanBankData;
    const fmtBalCls = v => v === 0 ? 'zero' : (v < 0 ? 'neg' : '');

    const html = [
        `<div class="bank-pill total" data-bank="__total__">
            <span class="bp-name">합계</span>
            <span class="bp-bal">${total.toLocaleString()}원</span>
        </div>`,
        ...banks.map(b => `
            <div class="bank-pill" data-bank="${escapeHtml(b.name)}">
                <span class="bp-name">${escapeHtml(b.name)}</span>
                <span class="bp-bal ${fmtBalCls(b.balance)}">${b.balance.toLocaleString()}원</span>
            </div>
        `)
    ].join('');
    pills.innerHTML = html;

    // hover → 인라인 미리보기 / click → 상세 모달
    pills.querySelectorAll('.bank-pill').forEach(el => {
        const name = el.dataset.bank;
        el.addEventListener('mouseenter', () => activateBank(name));
        el.addEventListener('click', () => openBankModal(name));
    });

    if (wrap) wrap.style.display = 'block';

    // 마지막으로 활성화한 은행 복원 (자동 갱신 시 깜빡임 방지)
    if (_humanBankActive) activateBank(_humanBankActive);
}

function activateBank(name) {
    _humanBankActive = name;
    const pills = document.getElementById('h-bank-pills');
    if (pills) {
        pills.querySelectorAll('.bank-pill').forEach(el => {
            el.classList.toggle('active', el.dataset.bank === name);
        });
    }
    renderHumanBankDetail(name);
}

function renderHumanBankDetail(name) {
    const detail = document.getElementById('h-bank-detail');
    if (!detail || !_humanBankData) return;
    const { bankNames, txns, weeks } = _humanBankData;

    const isTotal = name === '__total__';
    const targetTxns = isTotal ? txns : txns.filter(t => t.bank === name);
    const headerName = isTotal ? '전체 합계' : name;

    // 주간 변동
    const fmt = v => (v >= 0 ? '+' : '') + v.toLocaleString();
    const cls = v => v > 0 ? 'pos' : (v < 0 ? 'neg' : 'zero');
    const weekCells = weeks.map(w => {
        const net = targetTxns
            .filter(t => t.serial >= w.startSerial && t.serial <= w.endSerial)
            .reduce((s, t) => s + t.dep - t.wd, 0);
        return `<div class="bd-trend-cell">
            <div class="bdt-w">${w.label}</div>
            <div class="bdt-r">${w.range}</div>
            <div class="bdt-v ${cls(net)}">${net === 0 ? '—' : fmt(net)}</div>
        </div>`;
    }).join('');
    const monthNet = targetTxns.reduce((s, t) => s + t.dep - t.wd, 0);
    const monthCell = `<div class="bd-trend-cell month-total">
        <div class="bdt-w">월계</div>
        <div class="bdt-r">합산</div>
        <div class="bdt-v ${cls(monthNet)}">${monthNet === 0 ? '—' : fmt(monthNet)}</div>
    </div>`;

    // 주요 매출 기업 (입금 적요별 상위 5)
    const byDesc = {};
    targetTxns.forEach(t => {
        if (t.dep <= 0) return;
        const key = t.desc || '(미기재)';
        if (!byDesc[key]) byDesc[key] = { sum: 0, count: 0 };
        byDesc[key].sum += t.dep;
        byDesc[key].count++;
    });
    const sorted = Object.entries(byDesc).sort((a, b) => b[1].sum - a[1].sum);
    const top = sorted.slice(0, 5);
    const totalDeposits = sorted.reduce((s, [, v]) => s + v.sum, 0);
    const topHtml = top.length === 0
        ? `<div class="bd-empty">해당 기간 입금 내역 없음</div>`
        : `<ol class="bd-top-list">${top.map(([desc, v]) => `
            <li>
                <span class="bdt-name" title="${escapeHtml(desc)}">${escapeHtml(desc)}</span>
                <span class="bdt-amt">${v.sum.toLocaleString()}<span class="bdt-cnt">×${v.count}</span></span>
            </li>
        `).join('')}</ol>`;

    const monthLabel = weeks.length ? (() => {
        const d = new Date((weeks[0].startSerial - 25569) * 86400 * 1000);
        return `${String(d.getFullYear()).slice(2)}-${String(d.getMonth() + 1).padStart(2, '0')}월`;
    })() : '';

    detail.innerHTML = `
        <div class="bank-detail-content">
            <div class="bank-detail-section">
                <h6>
                    <span><span class="bds-bank">${escapeHtml(headerName)}</span> · 📈 주간 변동</span>
                    <span class="bds-meta">${monthLabel} · ${targetTxns.length}건</span>
                </h6>
                <div class="bd-trend">${weekCells}${monthCell}</div>
            </div>
            <div class="bank-detail-section">
                <h6>
                    <span><span class="bds-bank">${escapeHtml(headerName)}</span> · 🏢 주요 매출 기업</span>
                    <span class="bds-meta">∑ ${totalDeposits.toLocaleString()}원 · 클릭 시 상세창</span>
                </h6>
                ${topHtml}
            </div>
        </div>
    `;
}

// 은행 상세 모달 — 클릭 시 표시
function openBankModal(name) {
    if (!_humanBankData) return;
    const modal = document.getElementById('bank-detail-modal');
    if (!modal) return;

    const { banks, txns, weeks, total } = _humanBankData;
    const isTotal = name === '__total__';
    const headerName = isTotal ? '전체 합계' : name;
    const targetTxns = isTotal ? txns : txns.filter(t => t.bank === name);
    const currentBal = isTotal ? total : (banks.find(b => b.name === name) || { balance: 0 }).balance;

    const fmt = v => (v >= 0 ? '+' : '') + v.toLocaleString();
    const cls = v => v > 0 ? 'pos' : (v < 0 ? 'neg' : 'zero');
    const fmtDate = serial => {
        const d = new Date((serial - 25569) * 86400 * 1000);
        return `${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };
    const fmtFullDate = serial => {
        const d = new Date((serial - 25569) * 86400 * 1000);
        const wd = ['일','월','화','수','목','금','토'][d.getDay()];
        return `${String(d.getFullYear()).slice(2)}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} (${wd})`;
    };

    // 제목 + 잔액 태그
    const title = document.getElementById('bank-modal-title');
    if (title) {
        const balCls = currentBal === 0 ? 'zero' : (currentBal < 0 ? 'neg' : '');
        const balColor = balCls === 'neg' ? '#ffcccc' : (balCls === 'zero' ? '#d6eaf8' : 'white');
        title.innerHTML = `🏦 ${escapeHtml(headerName)} 상세 내역 <span class="bm-bal-tag" style="color:${balColor};">${currentBal.toLocaleString()}원</span>`;
    }

    // 1. 요약 카드 4개
    const monthDep = targetTxns.reduce((s, t) => s + t.dep, 0);
    const monthWd = targetTxns.reduce((s, t) => s + t.wd, 0);
    const monthNet = monthDep - monthWd;
    const summaryHtml = `
        <div class="bm-summary">
            <div class="bm-sum-card"><div class="bms-label">현재 잔액</div><div class="bms-value">${currentBal.toLocaleString()}원</div></div>
            <div class="bm-sum-card"><div class="bms-label">월 입금 합계</div><div class="bms-value pos">+${monthDep.toLocaleString()}원</div></div>
            <div class="bm-sum-card"><div class="bms-label">월 출금 합계</div><div class="bms-value neg">−${monthWd.toLocaleString()}원</div></div>
            <div class="bm-sum-card"><div class="bms-label">월 순변동 (${targetTxns.length}건)</div><div class="bms-value ${cls(monthNet)}">${monthNet === 0 ? '—' : fmt(monthNet)}원</div></div>
        </div>
    `;

    // 2. 주간 변동 추이 (확대판)
    const weekCells = weeks.map(w => {
        const net = targetTxns
            .filter(t => t.serial >= w.startSerial && t.serial <= w.endSerial)
            .reduce((s, t) => s + t.dep - t.wd, 0);
        return `<div class="bm-trend-cell">
            <div class="bmt-w">${w.label}</div>
            <div class="bmt-r">${w.range}</div>
            <div class="bmt-v ${cls(net)}">${net === 0 ? '—' : fmt(net)}</div>
        </div>`;
    }).join('');
    const trendMonthCell = `<div class="bm-trend-cell month-total">
        <div class="bmt-w">월계</div>
        <div class="bmt-r">합산</div>
        <div class="bmt-v ${cls(monthNet)}">${monthNet === 0 ? '—' : fmt(monthNet)}</div>
    </div>`;
    const monthLabel = weeks.length ? (() => {
        const d = new Date((weeks[0].startSerial - 25569) * 86400 * 1000);
        return `${String(d.getFullYear()).slice(2)}-${String(d.getMonth() + 1).padStart(2, '0')}월`;
    })() : '';

    // 3. 주요 입금/출금 (적요별 상위 5)
    const topByKind = (kind) => {
        const by = {};
        targetTxns.forEach(t => {
            const amt = kind === 'dep' ? t.dep : t.wd;
            if (amt <= 0) return;
            const key = t.desc || '(미기재)';
            if (!by[key]) by[key] = { sum: 0, count: 0 };
            by[key].sum += amt;
            by[key].count++;
        });
        const sorted = Object.entries(by).sort((a, b) => b[1].sum - a[1].sum);
        return { sorted, top: sorted.slice(0, 5), total: sorted.reduce((s, [, v]) => s + v.sum, 0) };
    };
    const dep = topByKind('dep');
    const wd = topByKind('wd');
    const renderTopCard = (data, kind) => {
        const label = kind === 'dep' ? '주요 입금처' : '주요 출금 항목';
        if (data.top.length === 0) {
            return `<div class="bm-top-card ${kind}">
                <h6><span>${kind === 'dep' ? '💰' : '💸'} ${label}</span><span class="bmt-sum">∑ 0원</span></h6>
                <div class="bmt-empty">해당 기간 ${kind === 'dep' ? '입금' : '출금'} 내역 없음</div>
            </div>`;
        }
        const items = data.top.map(([desc, v]) => `
            <li>
                <span class="bmt-name" title="${escapeHtml(desc)}">${escapeHtml(desc)}</span>
                <span class="bmt-amt">${v.sum.toLocaleString()}<span class="bmt-cnt">×${v.count}</span></span>
            </li>
        `).join('');
        return `<div class="bm-top-card ${kind}">
            <h6><span>${kind === 'dep' ? '💰' : '💸'} ${label} (상위 5)</span><span class="bmt-sum">∑ ${data.total.toLocaleString()}원</span></h6>
            <ol>${items}</ol>
        </div>`;
    };

    // 4. 최근 10일 거래 내역 (가장 최근 거래일 기준 10일 전부터)
    const allSerials = txns.map(t => t.serial);
    const refSerial = allSerials.length ? Math.max(...allSerials) : 0;
    const cutoffSerial = refSerial - 9; // 최근 10일 (오늘 포함)

    // 잔액 흐름 계산: 현재잔액 = 기초잔액 + sum(전체 거래의 net) → 기초잔액 역산
    const sortedAll = [...targetTxns].sort((a, b) => a.serial - b.serial || 0);
    const totalNet = sortedAll.reduce((s, t) => s + t.dep - t.wd, 0);
    const openingBal = currentBal - totalNet;
    let running = openingBal;
    const txnsWithRunning = sortedAll.map(t => {
        running += t.dep - t.wd;
        return { ...t, running };
    });
    // 최근 10일만 필터, 최신 → 과거 순으로 표시
    const recent = txnsWithRunning.filter(t => t.serial >= cutoffSerial).reverse();

    let ledgerHtml;
    if (recent.length === 0) {
        ledgerHtml = `<div class="bm-ledger-empty">최근 10일 내 거래 내역이 없습니다.</div>`;
    } else {
        // 일자별로 묶어 day-divider 행 추가
        const rowsHtml = [];
        let prevSerial = null;
        // recent는 최신부터 → 일자 그룹핑도 그대로
        const dayGroups = {};
        recent.forEach(t => {
            if (!dayGroups[t.serial]) dayGroups[t.serial] = [];
            dayGroups[t.serial].push(t);
        });
        const sortedDays = Object.keys(dayGroups).map(Number).sort((a, b) => b - a);
        sortedDays.forEach(serial => {
            const items = dayGroups[serial];
            const dayDep = items.reduce((s, t) => s + t.dep, 0);
            const dayWd = items.reduce((s, t) => s + t.wd, 0);
            const dayNet = dayDep - dayWd;
            const colspan = isTotal ? 5 : 4;
            rowsHtml.push(`<tr class="day-divider"><td colspan="${colspan}">${fmtFullDate(serial)} <span class="day-net ${cls(dayNet)}">일 순변동: ${dayNet === 0 ? '0' : fmt(dayNet)}원</span></td></tr>`);
            items.forEach(t => {
                const depCell = t.dep > 0 ? `<td class="dep">+${t.dep.toLocaleString()}</td>` : `<td class="dep zero-cell">—</td>`;
                const wdCell = t.wd > 0 ? `<td class="wd">−${t.wd.toLocaleString()}</td>` : `<td class="wd zero-cell">—</td>`;
                const bankCell = isTotal ? `<td class="bank">${escapeHtml(t.bank)}</td>` : '';
                rowsHtml.push(`<tr>
                    <td class="date">${fmtDate(t.serial)}</td>
                    ${bankCell}
                    <td class="desc">${escapeHtml(t.desc || '(미기재)')}</td>
                    ${depCell}
                    ${wdCell}
                    <td class="run">${t.running.toLocaleString()}</td>
                </tr>`);
            });
        });
        const headBank = isTotal ? '<th>은행</th>' : '';
        ledgerHtml = `<div class="bm-ledger-wrap">
            <table class="bm-ledger-table">
                <thead><tr><th>날짜</th>${headBank}<th>적요</th><th>입금</th><th>출금</th><th>잔액</th></tr></thead>
                <tbody>${rowsHtml.join('')}</tbody>
            </table>
        </div>`;
    }

    const refLabel = refSerial ? `${fmtDate(cutoffSerial)} ~ ${fmtDate(refSerial)}` : '데이터 없음';

    document.getElementById('bank-modal-body').innerHTML = `
        ${summaryHtml}
        <div class="bm-section">
            <div class="bm-section-title"><span>📈 주간 변동 추이</span><span class="bms-meta">${monthLabel}</span></div>
            <div class="bm-trend">${weekCells}${trendMonthCell}</div>
        </div>
        <div class="bm-section">
            <div class="bm-section-title"><span>💱 주요 입출금 현황</span><span class="bms-meta">월 ${targetTxns.length}건 기준</span></div>
            <div class="bm-top-grid">${renderTopCard(dep, 'dep')}${renderTopCard(wd, 'wd')}</div>
        </div>
        <div class="bm-section">
            <div class="bm-section-title"><span>📋 최근 10일 입출금 내역</span><span class="bms-meta">${refLabel} · ${recent.length}건</span></div>
            ${ledgerHtml}
        </div>
    `;

    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

function closeBankModal() {
    const modal = document.getElementById('bank-detail-modal');
    if (modal) modal.style.display = 'none';
    document.body.style.overflow = '';
}

// ESC 키로 모달 닫기
document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    const vpm = document.getElementById('vendor-pattern-modal');
    if (vpm && vpm.style.display === 'block') { closeVendorPatternModal(); return; }
    const vm = document.getElementById('vendor-detail-modal');
    if (vm && vm.style.display === 'block') { closeVendorModal(); return; }
    const wm = document.getElementById('worker-card-modal');
    if (wm && wm.style.display === 'block') { closeWorkerModal(); return; }
    const bm = document.getElementById('bank-detail-modal');
    if (bm && bm.style.display === 'block') closeBankModal();
});

// 업체 상세 모달 — 청구→입금 갭 분석
function openVendorModal(vendorName) {
    const modal = document.getElementById('vendor-detail-modal');
    const body = document.getElementById('vendor-modal-body');
    const title = document.getElementById('vendor-modal-title');
    if (!modal || !body || !title) return;

    const stats = findVendorStats(vendorName);
    const patternLabel = stats ? stats.pattern : '데이터없음';
    const patternCls = patternToClass(patternLabel);
    title.innerHTML = `📑 ${escapeHtml(vendorName)} <span class="pattern-badge ${patternCls}">${patternLabel}</span>`;

    if (!stats || stats.count === 0) {
        body.innerHTML = `
            <div style="padding: 40px; text-align: center; color: #9ca3af;">
                <div style="font-size: 2.5rem; margin-bottom: 12px;">📭</div>
                <div style="font-size: 0.95rem; font-weight: 600; color: #6b7280;">매출 거래 이력 없음</div>
                <div style="font-size: 0.78rem; margin-top: 8px;">금전출납(세무용·휴먼) 시트의 매출 데이터에서 매칭되는 업체를 찾지 못했습니다.<br>업체명 표기가 일치하는지 확인하세요.</div>
            </div>
        `;
        modal.style.display = 'block';
        document.body.style.overflow = 'hidden';
        return;
    }

    const fmtSerial = s => {
        if (typeof s !== 'number') return '—';
        const d = new Date((s - 25569) * 86400 * 1000);
        return `${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };
    const fmtFullSerial = s => {
        if (typeof s !== 'number') return '—';
        const d = new Date((s - 25569) * 86400 * 1000);
        return `${String(d.getFullYear()).slice(2)}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };

    // 1. 요약 4카드
    const unpaidCount = stats.count - stats.paidCount;
    const summaryHtml = `
        <div class="vm-summary">
            <div class="vm-sum-card"><div class="vms-label">총 청구건수</div><div class="vms-val">${stats.count}건</div></div>
            <div class="vm-sum-card ok"><div class="vms-label">입금완료</div><div class="vms-val">${stats.paidCount}건 / ${stats.paidAmount.toLocaleString()}원</div></div>
            <div class="vm-sum-card alert"><div class="vms-label">미입금</div><div class="vms-val">${unpaidCount}건 / ${stats.unpaidAmount.toLocaleString()}원</div></div>
            <div class="vm-sum-card"><div class="vms-label">평균/중앙값 갭</div><div class="vms-val">${stats.avg ? stats.avg.toFixed(1) : '—'}일 / ${stats.median ?? '—'}일</div></div>
        </div>
    `;

    // 2. 갭 히스토그램
    const total = stats.gaps.length;
    const histDef = [
        { key: 'sameDay', label: '당일', range: '0~2일', cls: 'same' },
        { key: 'short', label: '단기', range: '3~7일', cls: 'short' },
        { key: 'tenDay', label: '10일', range: '8~12일', cls: 'ten' },
        { key: 'biweek', label: '2주', range: '13~17일', cls: 'biweek' },
        { key: 'longTerm', label: '장기', range: '18~30일', cls: 'long' },
        { key: 'veryLong', label: '매우장기', range: '30일+', cls: 'very' },
    ];
    const histHtml = `
        <div class="vm-section">
            <h4>📊 청구→입금 갭 분포 <span class="vms-meta">유효 ${total}건 / 최소 ${stats.min ?? '—'}일 · 최대 ${stats.max ?? '—'}일</span></h4>
            <div class="vm-hist">
                ${histDef.map(h => {
                    const c = stats.buckets[h.key] || 0;
                    const pct = total > 0 ? (c / total * 100).toFixed(0) : 0;
                    return `<div class="vm-hist-bar ${h.cls}">
                        <div class="vmh-label">${h.label}</div>
                        <div class="vmh-range">${h.range}</div>
                        <div class="vmh-count">${c}</div>
                        <div class="vmh-pct">${pct}%</div>
                    </div>`;
                }).join('')}
            </div>
        </div>
    `;

    // 3. 다음 예상 입금일 안내 (최근 미입금 청구 + 평균 갭)
    let forecastHtml = '';
    const unpaidTxns = stats.txns.filter(t => !t.paid).sort((a, b) => b.created - a.created);
    if (unpaidTxns.length > 0 && stats.median !== null) {
        const recent = unpaidTxns[0];
        const expected = recent.created + (stats.median || 10);
        const today = Math.round(new Date().getTime() / 86400000) + 25569;
        const daysUntil = expected - today;
        const status = daysUntil > 0 ? `${daysUntil}일 후 예정` : (daysUntil === 0 ? '오늘 예정' : `${-daysUntil}일 지연`);
        const cls = daysUntil >= 0 ? 'ok' : 'alert';
        forecastHtml = `
            <div class="vm-forecast" style="${daysUntil < 0 ? 'background:linear-gradient(135deg,#fef2f2,#fecaca);border-color:#f87171;' : ''}">
                <div class="vmf-icon">${daysUntil >= 0 ? '📅' : '⚠️'}</div>
                <div class="vmf-text">
                    <div class="vmf-title">📌 다음 예상 입금일: ${fmtFullSerial(expected)} (${status})</div>
                    <div class="vmf-detail">최근 청구 ${fmtFullSerial(recent.created)} · ${recent.sales.toLocaleString()}원 · 패턴 중앙값 ${stats.median}일 적용</div>
                </div>
            </div>
        `;
    }

    // 4. 청구 내역 테이블 (최신순, 최근 30건)
    const sortedTxns = [...stats.txns].sort((a, b) => b.created - a.created).slice(0, 30);
    const ledgerRows = sortedTxns.map(t => {
        const gapCls = t.gap === null ? 'unpaid' : (t.gap <= 2 ? 'same-day' : (t.gap <= 12 ? 'ten-day' : 'long-day'));
        const gapText = t.gap === null ? '미입금' : `${t.gap}일`;
        return `
            <tr>
                <td class="date">${fmtSerial(t.created)}</td>
                <td class="amt">${t.sales.toLocaleString()}</td>
                <td class="date">${t.ipgum ? fmtSerial(t.ipgum) : '—'}</td>
                <td class="gap ${gapCls}">${gapText}</td>
                <td class="memo">${escapeHtml(t.memo || '')}</td>
            </tr>
        `;
    }).join('');
    const ledgerHtml = `
        <div class="vm-section">
            <h4>📋 청구 내역 (최근 30건) <span class="vms-meta">전체 ${stats.count}건 중</span></h4>
            <div class="vm-ledger-wrap">
                <table class="vm-ledger-table">
                    <thead><tr><th>청구일</th><th>금액</th><th>입금일</th><th>갭</th><th>비고</th></tr></thead>
                    <tbody>${ledgerRows || '<tr><td colspan="5" style="text-align:center;color:#9ca3af;padding:14px;">데이터 없음</td></tr>'}</tbody>
                </table>
            </div>
        </div>
    `;

    body.innerHTML = summaryHtml + histHtml + forecastHtml + ledgerHtml;
    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

function closeVendorModal() {
    const modal = document.getElementById('vendor-detail-modal');
    if (modal) modal.style.display = 'none';
    if (!isAnyModalOpen()) document.body.style.overflow = '';
}

function isAnyModalOpen() {
    return ['vendor-detail-modal', 'vendor-pattern-modal', 'worker-card-modal', 'bank-detail-modal']
        .some(id => {
            const el = document.getElementById(id);
            return el && el.style.display === 'block';
        });
}

// 업체 지급패턴 분류 모달
function openVendorPatternModal() {
    if (!_humanVendorStats) return;
    const modal = document.getElementById('vendor-pattern-modal');
    const body = document.getElementById('vendor-pattern-body');
    if (!modal || !body) return;

    const { byPattern, patternOrder, all } = _humanVendorStats;
    const helpHtml = `
        <div class="vp-help">
            💡 <b>지급패턴 분류 기준</b> — 매출 거래의 청구일(작성일자) → 입금일 갭을 분석하여 자동 분류합니다.<br>
            • <b>당일지급</b>: 갭 0~2일이 50% 이상 · 즉시 정산 업체 (현금성·내부거래 위주)<br>
            • <b>10일지급</b>: 갭 8~12일이 50% 이상 · 정기 결제일 업체 (가장 일반적)<br>
            • <b>혼합 (당일+10일)</b>: 두 패턴이 모두 20% 이상 — 일정 관리 시 <b style="color:#92400e;">혼선 주의</b><br>
            • <b>장기지급</b>: 중앙값 18일 초과 · 분기/세금계산서 정산 업체<br>
            업체명을 클릭하면 상세 분석을 볼 수 있습니다.
        </div>
    `;

    const patternMeta = {
        '당일지급': { icon: '⚡', desc: '청구 즉시(0~2일 내) 입금되는 업체. 현금성 거래·내부 정산 위주.', cls: 'pat-same' },
        '10일지급': { icon: '📅', desc: '청구 후 약 10일(8~12일) 후 입금되는 업체. 가장 일반적인 정기 결제 패턴.', cls: 'pat-ten' },
        '혼합': { icon: '⚠️', desc: '당일지급과 10일지급 패턴이 혼재. 청구별로 입금 시점이 달라 일정 관리 주의 필요.', cls: 'pat-mix' },
        '장기지급': { icon: '🐢', desc: '중앙값 18일 초과로 장기간 후 입금. 자금 흐름 예측 시 별도 고려.', cls: 'pat-long' },
        '기타': { icon: '📊', desc: '특정 패턴에 부합하지 않는 업체.', cls: 'pat-etc' },
        '데이터부족': { icon: '❓', desc: '거래 건수가 부족해 패턴을 추정할 수 없음.', cls: 'pat-na' }
    };

    const tabs = patternOrder.filter(p => byPattern[p].length > 0);
    const tabsHtml = tabs.map((p, i) => `
        <button class="vp-tab ${i === 0 ? 'active' : ''}" data-pattern="${p}" onclick="switchVendorPatternTab('${p}')">
            ${patternMeta[p]?.icon || '•'} ${p}
            <span class="vp-tab-count">${byPattern[p].length}</span>
        </button>
    `).join('');

    const panesHtml = tabs.map((p, i) => {
        const list = byPattern[p];
        const meta = patternMeta[p] || {};
        const items = list.map(v => `
            <div class="vp-vendor-item" data-vendor="${escapeHtml(v.name)}">
                <span class="vpv-name" title="${escapeHtml(v.name)}">${escapeHtml(v.name)}</span>
                <span class="vpv-stat"><span class="vpv-l">중앙갭/평균</span>${v.median ?? '—'}일 / ${v.avg ? v.avg.toFixed(1) : '—'}일</span>
                <span class="vpv-amt"><span class="vpv-l" style="color:#9ca3af;">총 청구</span>${v.totalAmount.toLocaleString()}원</span>
                <span class="vpv-paid"><span class="vpv-l" style="color:#9ca3af;">입금완료</span>${v.paidCount}/${v.count}건</span>
                <span class="vpv-arrow">→</span>
            </div>
        `).join('') || '<div class="vp-empty">해당 패턴의 업체 없음</div>';
        return `
            <div class="vp-pane ${i === 0 ? 'active' : ''}" data-pattern="${p}">
                <div class="vp-pane-desc">${meta.icon || ''} ${meta.desc || ''}</div>
                <div class="vp-vendor-list">${items}</div>
            </div>
        `;
    }).join('');

    body.innerHTML = helpHtml +
        `<div class="vp-tabs">${tabsHtml}</div>` +
        panesHtml +
        `<p style="margin-top:14px; font-size:0.7rem; color:#95a5a6; text-align:right;">📌 분석 대상: 매출 거래 ${all.length}개 업체 (2건 이상)</p>`;

    // 업체 클릭 → 패턴 모달 닫고 상세 모달 열기 (위임 핸들러)
    body.querySelectorAll('.vp-vendor-item').forEach(el => {
        el.addEventListener('click', () => {
            const name = el.dataset.vendor;
            closeVendorPatternModal();
            openVendorModal(name);
        });
    });

    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

function switchVendorPatternTab(pattern) {
    document.querySelectorAll('.vp-tab').forEach(t => t.classList.toggle('active', t.dataset.pattern === pattern));
    document.querySelectorAll('.vp-pane').forEach(p => p.classList.toggle('active', p.dataset.pattern === pattern));
}

function closeVendorPatternModal() {
    const modal = document.getElementById('vendor-pattern-modal');
    if (modal) modal.style.display = 'none';
    if (!isAnyModalOpen()) document.body.style.overflow = '';
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

        // 입사일/퇴사일 자동 계산 — 데이터의 첫·마지막 근무일
        const firstDay = uniqueDays.length ? uniqueDays[0] : 0;
        const lastDay = uniqueDays.length ? uniqueDays[uniqueDays.length - 1] : 0;
        const spanDays = firstDay && lastDay ? (lastDay - firstDay + 1) : 0;
        const attendance = spanDays > 0 ? (totalDays / spanDays) : 0;
        const firstDayInfo = w.days.find(d => d.day === firstDay) || w.days[0] || {};
        const lastDayInfo = w.days.slice().reverse().find(d => d.day === lastDay) || w.days[0] || {};
        const yyyy = firstDayInfo.year || (new Date()).getFullYear();
        const mm = firstDayInfo.month || (new Date()).getMonth() + 1;

        return {
            key: w.key, name: w.name, phone: w.phone, jumin: w.jumin,
            days: w.days, uniqueDays, totalDays, totalPay, totalCharge,
            maxConsecutive: { len: maxLen, start: maxStart, end: maxEnd },
            sites, primarySite, latestDay,
            firstDay, lastDay, spanDays, attendance,
            firstDate: { year: firstDayInfo.year || yyyy, month: firstDayInfo.month || mm, day: firstDay },
            lastDate: { year: lastDayInfo.year || yyyy, month: lastDayInfo.month || mm, day: lastDay }
        };
    });
    workers.sort((a, b) => b.totalDays - a.totalDays || b.totalPay - a.totalPay);

    // 데이터셋 전체의 최신 근무일 → 재직 상태 추정 기준
    const datasetMaxDay = workers.reduce((max, w) => Math.max(max, w.lastDay), 0);
    const datasetMinDay = workers.reduce((min, w) => w.firstDay > 0 ? Math.min(min, w.firstDay) : min, datasetMaxDay || 31);
    const datasetMonthLen = (() => {
        const w = workers.find(x => x.firstDate && x.firstDate.year);
        if (!w) return 31;
        return new Date(w.firstDate.year, w.firstDate.month, 0).getDate();
    })();
    workers.forEach(w => {
        const gap = datasetMaxDay - w.lastDay;
        // 마지막 근무일이 데이터 최신일과 7일 이내면 재직중, 아니면 추정 퇴사
        w.isActive = gap <= 7;
        w.exitGap = gap;
    });

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

    _humanWorkerData = { dailyTotals, workers, byVendor, byKey, datasetMaxDay, datasetMinDay, datasetMonthLen };

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

    // 직원카드 모달 표시
    openWorkerModal(key);

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

// 직원카드 모달
function openWorkerModal(key) {
    if (!_humanWorkerData) return;
    const w = _humanWorkerData.byKey[key];
    if (!w) return;
    const modal = document.getElementById('worker-card-modal');
    const body = document.getElementById('worker-modal-body');
    if (!modal || !body) return;

    const { datasetMaxDay, datasetMonthLen } = _humanWorkerData;
    const fmtDate = (info) => {
        if (!info || !info.day) return '—';
        return `${info.year}-${String(info.month).padStart(2, '0')}-${String(info.day).padStart(2, '0')}`;
    };
    const fmtMD = (info) => {
        if (!info || !info.day) return '—';
        return `${String(info.month).padStart(2, '0')}/${String(info.day).padStart(2, '0')}`;
    };

    // 마스킹된 주민번호 (앞 6자리만 노출)
    const juminMasked = w.jumin
        ? w.jumin.replace(/^(\d{6})[-]?(\d?).*/, '$1-$2******').replace(/-$/, '-*******')
        : '—';

    const initial = (w.name || '?').slice(0, 1);
    const empNo = (w.jumin && w.jumin.length >= 6) ? w.jumin.slice(0, 6) : (w.phone ? w.phone.replace(/-/g, '').slice(-4) : 'N/A');

    // 입사일 / 마지막 근무일 / 활동기간 / 퇴사 추정
    const exitText = w.isActive ? '재직중' : `추정 ${fmtMD(w.lastDate)}`;
    const exitSub = w.isActive
        ? `최근 근무: ${w.lastDay}일 (${w.exitGap}일 전)`
        : `${w.exitGap}일간 미출근`;

    // 1. ID 카드 헤더
    const idCardHtml = `
        <div class="wc-card">
            <div class="wc-id-band">
                <span class="wc-id-co">HUMANMS · 일용근로자 카드</span>
                <span class="wc-id-no">EMP-${empNo}</span>
            </div>
            <div class="wc-id-main">
                <div class="wc-avatar"><span class="wc-avatar-text">${escapeHtml(initial)}</span></div>
                <div class="wc-info">
                    <div class="wc-name">
                        ${escapeHtml(w.name)}
                        <span class="wc-status-badge ${w.isActive ? 'active' : 'inactive'}">${w.isActive ? '● 재직중' : '◐ 추정 퇴사'}</span>
                    </div>
                    <div class="wc-meta">
                        <div class="wc-meta-item"><span class="wc-meta-label">📞 연락처</span><span class="wc-meta-val">${escapeHtml(w.phone || '—')}</span></div>
                        <div class="wc-meta-item"><span class="wc-meta-label">🆔 주민번호</span><span class="wc-meta-val">${escapeHtml(juminMasked)}</span></div>
                        <div class="wc-meta-item"><span class="wc-meta-label">📍 주현장</span><span class="wc-meta-val">${escapeHtml(w.primarySite || '—')}</span></div>
                        <div class="wc-meta-item"><span class="wc-meta-label">🏢 현장수</span><span class="wc-meta-val">${w.sites.length}개소</span></div>
                    </div>
                </div>
            </div>
            <div class="wc-dates">
                <div class="wc-date-card entry">
                    <div class="wcd-label">📥 입사일 (자동)</div>
                    <div class="wcd-val">${fmtDate(w.firstDate)}</div>
                    <div class="wcd-sub">데이터 첫 근무일</div>
                </div>
                <div class="wc-date-card exit ${w.isActive ? 'active' : ''}">
                    <div class="wcd-label">📤 퇴사일 (추정)</div>
                    <div class="wcd-val">${exitText}</div>
                    <div class="wcd-sub">${exitSub}</div>
                </div>
                <div class="wc-date-card">
                    <div class="wcd-label">📆 근무일수</div>
                    <div class="wcd-val">${w.totalDays}일</div>
                    <div class="wcd-sub">고유 근무일 수</div>
                </div>
                <div class="wc-date-card">
                    <div class="wcd-label">📊 활동 기간</div>
                    <div class="wcd-val">${w.spanDays}일</div>
                    <div class="wcd-sub">출근율 ${(w.attendance * 100).toFixed(0)}%</div>
                </div>
            </div>
        </div>
    `;

    // 2. 통계 카드 4개
    const avgPay = w.totalDays > 0 ? Math.round(w.totalPay / w.totalDays) : 0;
    const statsHtml = `
        <div class="wc-section">
            <h4>💼 근무 통계 <span class="wcs-meta">${w.firstDate.year}년 ${w.firstDate.month}월 기준</span></h4>
            <div class="wc-stats">
                <div class="wc-stat-card"><div class="wcs-label">총 지급액</div><div class="wcs-val">${w.totalPay.toLocaleString()}원</div><div class="wcs-sub">${w.totalDays}일 합산</div></div>
                <div class="wc-stat-card"><div class="wcs-label">평균 일급</div><div class="wcs-val">${avgPay.toLocaleString()}원</div><div class="wcs-sub">총지급 ÷ 근무일</div></div>
                <div class="wc-stat-card"><div class="wcs-label">최대 연속</div><div class="wcs-val">${w.maxConsecutive.len}일</div><div class="wcs-sub">${w.maxConsecutive.len > 1 ? `${w.maxConsecutive.start}~${w.maxConsecutive.end}일` : '—'}</div></div>
                <div class="wc-stat-card"><div class="wcs-label">총 청구액</div><div class="wcs-val" style="color:#1f618d;">${w.totalCharge.toLocaleString()}원</div><div class="wcs-sub">현장 청구 합산</div></div>
            </div>
        </div>
    `;

    // 3. 근무 캘린더 (월 전체, 7열 그리드)
    const monthLen = datasetMonthLen || 31;
    const workedSet = new Set(w.uniqueDays);
    const streakSet = new Set();
    if (w.maxConsecutive.len > 1) {
        for (let d = w.maxConsecutive.start; d <= w.maxConsecutive.end; d++) streakSet.add(d);
    }
    // 요일 정렬을 위해 1일이 무슨 요일인지 확인
    const firstWd = new Date(w.firstDate.year, w.firstDate.month - 1, 1).getDay();
    const calCells = [];
    for (let i = 0; i < firstWd; i++) calCells.push(`<div class="wc-cal-cell outside"></div>`);
    for (let d = 1; d <= monthLen; d++) {
        const isWorked = workedSet.has(d);
        const isStreak = streakSet.has(d);
        const isEntry = d === w.firstDay;
        const isExit = !w.isActive && d === w.lastDay;
        const cls = ['wc-cal-cell'];
        if (isStreak) cls.push('streak');
        else if (isWorked) cls.push('worked');
        if (isEntry) cls.push('entry-day');
        if (isExit) cls.push('exit-day');
        // 그날 지급액
        const dayPay = w.days.filter(x => x.day === d).reduce((s, x) => s + x.pay, 0);
        const payText = dayPay > 0 ? (dayPay >= 10000 ? Math.round(dayPay / 10000) + '만' : '') : '';
        const tip = isWorked
            ? `${d}일 · 지급 ${dayPay.toLocaleString()}원${isEntry ? ' · 📥입사' : ''}${isExit ? ' · 📤퇴사' : ''}`
            : `${d}일 · 미근무`;
        calCells.push(`<div class="${cls.join(' ')}" title="${tip}">${d}<span class="wcc-pay">${payText}</span></div>`);
    }
    const calHtml = `
        <div class="wc-section">
            <h4>📅 근무 캘린더 <span class="wcs-meta">${w.firstDate.year}-${String(w.firstDate.month).padStart(2, '0')} · 데이터 ${datasetMaxDay}일까지</span></h4>
            <div class="wc-cal">
                <div class="wc-cal-cell outside" style="background:#fff5f5;color:#e74c3c;border:none;">일</div>
                <div class="wc-cal-cell outside" style="background:transparent;color:#7f8c8d;border:none;">월</div>
                <div class="wc-cal-cell outside" style="background:transparent;color:#7f8c8d;border:none;">화</div>
                <div class="wc-cal-cell outside" style="background:transparent;color:#7f8c8d;border:none;">수</div>
                <div class="wc-cal-cell outside" style="background:transparent;color:#7f8c8d;border:none;">목</div>
                <div class="wc-cal-cell outside" style="background:transparent;color:#7f8c8d;border:none;">금</div>
                <div class="wc-cal-cell outside" style="background:#f0f8ff;color:#2980b9;border:none;">토</div>
                ${calCells.join('')}
            </div>
            <div class="wc-cal-legend">
                <span><i class="lg-box" style="background:#d1fae5;border:1px solid #6ee7b7;"></i>근무</span>
                <span><i class="lg-box" style="background:#fde68a;border:1px solid #f59e0b;"></i>최대 연속 (${w.maxConsecutive.len}일)</span>
                <span><i class="lg-box" style="background:white;border:2px solid #047857;"></i>입사일</span>
                <span><i class="lg-box" style="background:white;border:2px solid #c0392b;"></i>퇴사일(추정)</span>
            </div>
        </div>
    `;

    // 4. 근무 현장
    const sitesHtml = `
        <div class="wc-section">
            <h4>📍 근무 현장 <span class="wcs-meta">총 ${w.sites.length}개소</span></h4>
            <div class="wc-sites">
                ${w.sites.map(s => `
                    <div class="wc-site-item">
                        <span class="wcsi-name">${escapeHtml(s.site)}</span>
                        <span class="wcsi-days">${s.days}일</span>
                        <span class="wcsi-pay">${s.pay.toLocaleString()}원</span>
                    </div>
                `).join('') || '<div style="color:#9ca3af;font-size:0.78rem;">근무 현장 없음</div>'}
            </div>
        </div>
    `;

    // 5. 일자별 근무 내역 (최신순)
    const sortedDayEntries = [...w.days].sort((a, b) => b.day - a.day);
    const ledgerRows = sortedDayEntries.map(d => `
        <tr>
            <td class="date">${String(d.month).padStart(2, '0')}/${String(d.day).padStart(2, '0')}</td>
            <td class="site">${escapeHtml(d.site || '—')}</td>
            <td class="time">${escapeHtml(d.workTime || '—')}</td>
            <td class="pay">${d.pay.toLocaleString()}</td>
            <td class="charge">${d.charge.toLocaleString()}</td>
            <td class="status"><span class="${d.paid ? 'badge-paid' : 'badge-unpaid'}">${d.paid ? '지급완료' : '미지급'}</span></td>
        </tr>
    `).join('');
    const ledgerHtml = `
        <div class="wc-section">
            <h4>📋 일자별 근무 내역 <span class="wcs-meta">총 ${w.days.length}건</span></h4>
            <div class="wc-ledger-wrap">
                <table class="wc-ledger-table">
                    <thead><tr><th>날짜</th><th>현장</th><th>시간</th><th>지급</th><th>청구</th><th>상태</th></tr></thead>
                    <tbody>${ledgerRows || '<tr><td colspan="6" style="text-align:center;color:#9ca3af;padding:14px;">근무 내역 없음</td></tr>'}</tbody>
                </table>
            </div>
        </div>
    `;

    body.innerHTML = idCardHtml + statsHtml + calHtml + sitesHtml + ledgerHtml;
    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
}

function closeWorkerModal() {
    const modal = document.getElementById('worker-card-modal');
    if (modal) modal.style.display = 'none';
    document.body.style.overflow = '';
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
