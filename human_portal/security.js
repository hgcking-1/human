// 보안 강화 — 클라이언트 측 다운로드/스크린샷 억제 + DevTools 억제
// 주의: 클라이언트 측 차단은 결정적이지 않으며, 결정적 차단은 백엔드 인증 필요.
(function () {
    'use strict';

    // 1) 우클릭 차단 (기본 컨텍스트 메뉴 → "이미지 저장", "다른 이름으로 저장" 차단)
    document.addEventListener('contextmenu', function (e) {
        e.preventDefault();
        return false;
    }, { capture: true });

    // 2) 키보드 단축키 차단 (저장/인쇄/소스보기/개발자도구)
    document.addEventListener('keydown', function (e) {
        const key = (e.key || '').toLowerCase();

        // F12 (개발자 도구)
        if (key === 'f12') { e.preventDefault(); return false; }

        // PrintScreen 시도 시 클립보드 비우기 + 화면 일시 블러
        if (key === 'printscreen' || e.code === 'PrintScreen') {
            try {
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    navigator.clipboard.writeText('⛔ 이 사이트의 자료는 캡처가 금지되어 있습니다.');
                }
            } catch (_) {}
            flashSecurityWarning('🚫 화면 캡처는 금지되어 있습니다');
            return;
        }

        if (e.ctrlKey || e.metaKey) {
            // Ctrl+S (저장), Ctrl+P (인쇄), Ctrl+U (소스보기) — 항상 차단
            if (['s', 'p', 'u'].includes(key)) {
                e.preventDefault();
                flashSecurityWarning('🚫 이 사이트는 자료 다운로드/인쇄를 허용하지 않습니다');
                return false;
            }
            // Ctrl+Shift+I/J/C (개발자도구), Ctrl+Shift+S (스크린샷)
            if (e.shiftKey && ['i', 'j', 'c', 's'].includes(key)) {
                e.preventDefault();
                return false;
            }
            // Ctrl+C / Ctrl+A 는 편집 가능 영역에서만 허용
            if (key === 'c' || key === 'a') {
                const t = e.target;
                const editable = t && (t.isContentEditable || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || (t.closest && t.closest('[contenteditable="true"]')));
                if (!editable) {
                    e.preventDefault();
                    return false;
                }
            }
        }
    }, { capture: true });

    // 3) 드래그 차단 (이미지/링크 드래그-드롭 저장 방지)
    document.addEventListener('dragstart', function (e) {
        e.preventDefault();
        return false;
    }, { capture: true });

    // 4) 텍스트 선택 차단 (입력 필드/contenteditable 제외)
    document.addEventListener('selectstart', function (e) {
        const t = e.target;
        const editable = t && (t.isContentEditable || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || (t.closest && t.closest('[contenteditable="true"]')));
        if (!editable) {
            e.preventDefault();
            return false;
        }
    }, { capture: true });

    // 5) copy 이벤트 차단 (편집 영역 외)
    document.addEventListener('copy', function (e) {
        const t = e.target;
        const editable = t && (t.isContentEditable || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || (t.closest && t.closest('[contenteditable="true"]')));
        if (!editable) {
            e.preventDefault();
            try { e.clipboardData && e.clipboardData.setData('text/plain', '⛔ 복사 금지'); } catch (_) {}
            return false;
        }
    }, { capture: true });

    // 6) 인쇄 차단 (window.print 호출 + 인쇄 미디어 쿼리는 CSS에서 처리)
    window.addEventListener('beforeprint', function (e) {
        flashSecurityWarning('🚫 인쇄가 차단되어 있습니다');
    });

    // 7) 포커스 손실 시 화면 블러 (스크린샷 도구 활성화 시 흔히 포커스 잃음)
    let blurOverlay = null;
    function ensureBlurOverlay() {
        if (blurOverlay) return blurOverlay;
        blurOverlay = document.createElement('div');
        blurOverlay.id = '__sec_blur_overlay';
        blurOverlay.style.cssText = [
            'position:fixed', 'inset:0', 'background:rgba(20,20,30,0.94)',
            'z-index:2147483646', 'display:none',
            'flex-direction:column', 'align-items:center', 'justify-content:center',
            'color:white', 'font-family:sans-serif', 'text-align:center', 'padding:40px',
            'pointer-events:auto'
        ].join(';');
        blurOverlay.innerHTML = '<div style="font-size:3rem;margin-bottom:16px;">🔒</div>' +
            '<div style="font-size:1.4rem;font-weight:700;margin-bottom:8px;">보안 보호 모드</div>' +
            '<div style="font-size:0.95rem;opacity:0.85;line-height:1.6;max-width:520px;">' +
            '창 포커스가 떠나면 자동으로 화면을 가립니다.<br>' +
            '데이터는 화면 캡처/다른 창에서 노출되지 않습니다.<br>' +
            '<small style="opacity:0.7;">창을 다시 활성화하면 자동 복귀합니다.</small></div>';
        document.body.appendChild(blurOverlay);
        return blurOverlay;
    }
    function showBlur() { ensureBlurOverlay().style.display = 'flex'; }
    function hideBlur() { if (blurOverlay) blurOverlay.style.display = 'none'; }

    // 페이지가 처음 마운트되기 전에는 hide
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initBlurGuard);
    } else {
        initBlurGuard();
    }
    function initBlurGuard() {
        window.addEventListener('blur', showBlur);
        window.addEventListener('focus', hideBlur);
        document.addEventListener('visibilitychange', function () {
            if (document.hidden) showBlur(); else hideBlur();
        });
    }

    // 8) 일시적 보안 경고 토스트
    let toastTimer = null;
    function flashSecurityWarning(msg) {
        let toast = document.getElementById('__sec_toast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = '__sec_toast';
            toast.style.cssText = [
                'position:fixed', 'top:20px', 'left:50%', 'transform:translateX(-50%)',
                'background:#c0392b', 'color:white', 'padding:12px 24px', 'border-radius:8px',
                'font-weight:700', 'font-size:0.95rem', 'box-shadow:0 8px 24px rgba(192,57,43,0.4)',
                'z-index:2147483647', 'display:none', 'pointer-events:none'
            ].join(';');
            document.body.appendChild(toast);
        }
        toast.textContent = msg;
        toast.style.display = 'block';
        if (toastTimer) clearTimeout(toastTimer);
        toastTimer = setTimeout(function () { toast.style.display = 'none'; }, 2200);
    }

    // 9) DevTools 감지 (window 크기 차이로 추정 — 100% 정확하진 않음)
    let devtoolsDetected = false;
    function detectDevTools() {
        const threshold = 200;
        const widthDiff = window.outerWidth - window.innerWidth;
        const heightDiff = window.outerHeight - window.innerHeight;
        const isOpen = widthDiff > threshold || heightDiff > threshold;
        if (isOpen && !devtoolsDetected) {
            devtoolsDetected = true;
            flashSecurityWarning('⚠️ 개발자 도구 사용이 감지되었습니다. 무단 사용 시 법적 책임이 따릅니다.');
        } else if (!isOpen) {
            devtoolsDetected = false;
        }
    }
    setInterval(detectDevTools, 1500);

    // 10) 워터마크 오버레이 — 로그인된 사용자 정보 + 시각 표시 (스크린샷 추적)
    function ensureWatermark() {
        let wm = document.getElementById('__sec_watermark');
        if (!wm) {
            wm = document.createElement('div');
            wm.id = '__sec_watermark';
            wm.style.cssText = [
                'position:fixed', 'inset:0', 'pointer-events:none',
                'z-index:9998', 'opacity:0.06', 'overflow:hidden',
                'font-family:sans-serif', 'color:#000'
            ].join(';');
            document.body.appendChild(wm);
        }
        let session = null;
        try { session = JSON.parse(localStorage.getItem('userSession') || 'null'); } catch (_) {}
        const who = session && (session.email || session.user || session.name) ? (session.email || session.user || session.name) : '익명';
        const now = new Date().toISOString().replace('T', ' ').slice(0, 19);
        const tile = `<div style="display:inline-block;width:340px;height:200px;text-align:center;line-height:200px;font-size:14px;font-weight:700;transform:rotate(-25deg);">${who} · ${now}</div>`;
        wm.innerHTML = Array(40).fill(tile).join('');
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', ensureWatermark);
    } else {
        ensureWatermark();
    }
    setInterval(ensureWatermark, 30000);

    // 11) 페이지 떠날 때 민감 데이터 캐시 무효화 트리거(선택)
    window.addEventListener('pagehide', function () { /* placeholder */ });
})();
