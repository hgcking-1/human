document.addEventListener('DOMContentLoaded', function() {
    checkExistingSession();
    refreshAdminPanel();
    initPurchaseChart(); 
});

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

        // 2. 매출액
        const salesEl = document.getElementById(`${prefix}-val-sales`);
        if (salesEl) {
            salesEl.classList.add('unmasked');
            const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_sales`));
            if (saved && saved.data && saved.data.length > 0) {
                const total = saved.data.reduce((acc, row) => {
                    let amountStr = (row[2] || '0').toString();
                    return acc + (parseInt(amountStr.replace(/[^0-9]/g, '')) || 0);
                }, 0);
                salesEl.textContent = total.toLocaleString() + '원';
            } else {
                salesEl.textContent = prefix === 'h' ? '1,250,000,000원' : '780,000,000원';
            }
        }

        // 3. 매입계산서
        const purchaseEl = document.getElementById(`${prefix}-val-purchase`);
        if (purchaseEl) {
            purchaseEl.classList.add('unmasked');
            const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_purchase`));
            if (saved && saved.data && saved.data.length > 0) {
                const total = saved.data.reduce((acc, row) => acc + (parseInt(row[2].toString().replace(/[^0-9]/g, '')) || 0), 0);
                purchaseEl.innerHTML = `<div class="stat-value" style="font-size: 1.8rem;">${total.toLocaleString()}원</div>`;
            } else {
                purchaseEl.innerHTML = `<div class="stat-value" style="font-size: 1.8rem;">${prefix === 'h' ? '425,110,200원' : '210,550,000원'}</div>`;
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

        // 은행 계좌 현황
        const bankEl = document.getElementById(`${prefix}-val-bank`);
        if (bankEl) {
            bankEl.classList.add('unmasked');
            const saved = JSON.parse(localStorage.getItem(`data_${companyKey}_bank`));
            if (saved && saved.data && saved.data.length > 0) {
                const total = saved.data.reduce((acc, row) => acc + (parseInt(row[1].toString().replace(/[^0-9]/g, '')) || 0), 0);
                bankEl.textContent = total.toLocaleString() + '원';
            } else {
                bankEl.textContent = prefix === 'h' ? '542,110,000원' : '125,440,000원';
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

        // 11. 일용직 근무 내역
        const workerEl = document.getElementById(`${prefix}-val-daily-worker`);
        if (workerEl) {
            workerEl.classList.add('unmasked');
            workerEl.textContent = prefix === 'h' ? '12,500,000원' : '8,400,000원';
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
