(function() {
    var ADMIN_PWD = 'Aa@0981737608';
    var DEFAULT_PWD = 'GAS2026';
    var STORAGE_KEY = 'gas_course_pwd';
    var AUTH_KEY = 'gas_course_auth';

    // 已驗證過就跳過
    if (sessionStorage.getItem(AUTH_KEY) === '1') return;

    // 取得目前密碼
    function getPwd() { return localStorage.getItem(STORAGE_KEY) || DEFAULT_PWD; }

    // 隱藏頁面內容
    document.documentElement.style.overflow = 'hidden';
    var hideStyle = document.createElement('style');
    hideStyle.id = 'gate-hide';
    hideStyle.textContent = 'body > *:not(#pw-gate) { display: none !important; }';
    document.head.appendChild(hideStyle);

    // 建立密碼閘門 UI
    var gate = document.createElement('div');
    gate.id = 'pw-gate';
    gate.innerHTML = [
        '<div style="position:fixed;inset:0;background:linear-gradient(135deg,#1b5e20 0%,#004d40 100%);z-index:99999;display:flex;align-items:center;justify-content:center;font-family:Microsoft JhengHei,Segoe UI,Arial,sans-serif;">',
        '  <div id="pw-card" style="background:#fff;border-radius:24px;padding:50px 40px;width:380px;max-width:90vw;box-shadow:0 20px 60px rgba(0,0,0,0.4);text-align:center;">',
        '    <div style="width:70px;height:70px;background:linear-gradient(135deg,#1b5e20,#4caf50);border-radius:50%;margin:0 auto 20px;display:flex;align-items:center;justify-content:center;font-size:32px;">🔒</div>',
        '    <h2 id="pw-title" style="color:#1b5e20;font-size:24px;margin-bottom:8px;">課程簡報驗證</h2>',
        '    <p id="pw-desc" style="color:#888;font-size:14px;margin-bottom:28px;">請輸入課程密碼以進入簡報</p>',
        '    <div id="pw-main">',
        '      <input id="pw-input" type="password" placeholder="請輸入密碼" style="width:100%;padding:14px 18px;border:2px solid #e0e0e0;border-radius:12px;font-size:18px;outline:none;transition:border-color 0.3s;text-align:center;letter-spacing:2px;">',
        '      <p id="pw-err" style="color:#e53935;font-size:13px;margin-top:10px;min-height:20px;"></p>',
        '      <button id="pw-btn" style="width:100%;padding:14px;background:linear-gradient(135deg,#1b5e20,#4caf50);color:#fff;border:none;border-radius:12px;font-size:18px;font-weight:700;cursor:pointer;transition:transform 0.2s,box-shadow 0.2s;box-shadow:0 4px 15px rgba(27,94,32,0.3);margin-top:5px;">進入簡報</button>',
        '    </div>',
        '    <div id="pw-admin" style="display:none;">',
        '      <div style="background:#e8f5e9;border-radius:12px;padding:16px;margin-bottom:20px;text-align:left;">',
        '        <p style="color:#1b5e20;font-size:13px;font-weight:700;margin-bottom:6px;">目前學員密碼：</p>',
        '        <p id="pw-current" style="color:#2e7d32;font-size:20px;font-weight:700;letter-spacing:2px;"></p>',
        '      </div>',
        '      <input id="pw-new" type="text" placeholder="輸入新密碼" style="width:100%;padding:14px 18px;border:2px solid #e0e0e0;border-radius:12px;font-size:18px;outline:none;text-align:center;letter-spacing:2px;margin-bottom:12px;">',
        '      <div style="display:flex;gap:10px;">',
        '        <button id="pw-save" style="flex:1;padding:14px;background:#1b5e20;color:#fff;border:none;border-radius:12px;font-size:16px;font-weight:700;cursor:pointer;">儲存新密碼</button>',
        '        <button id="pw-enter" style="flex:1;padding:14px;background:#4caf50;color:#fff;border:none;border-radius:12px;font-size:16px;font-weight:700;cursor:pointer;">直接進入</button>',
        '      </div>',
        '      <p id="pw-save-msg" style="color:#4caf50;font-size:13px;margin-top:10px;min-height:20px;"></p>',
        '    </div>',
        '    <p style="color:#bbb;font-size:11px;margin-top:20px;">GAS 自動化 × Line Bot 實戰教學 © 2026 阿亮老師</p>',
        '  </div>',
        '</div>'
    ].join('\n');

    // 插入 DOM
    function insertGate() {
        document.body.insertBefore(gate, document.body.firstChild);
        var input = document.getElementById('pw-input');
        var btn = document.getElementById('pw-btn');
        var err = document.getElementById('pw-err');

        // 聚焦
        setTimeout(function() { input.focus(); }, 100);

        // 輸入框效果
        input.addEventListener('focus', function() { this.style.borderColor = '#4caf50'; });
        input.addEventListener('blur', function() { this.style.borderColor = '#e0e0e0'; });

        // 按鈕效果
        btn.addEventListener('mouseenter', function() { this.style.transform = 'translateY(-2px)'; this.style.boxShadow = '0 6px 20px rgba(27,94,32,0.4)'; });
        btn.addEventListener('mouseleave', function() { this.style.transform = ''; this.style.boxShadow = '0 4px 15px rgba(27,94,32,0.3)'; });

        // Enter 鍵
        input.addEventListener('keydown', function(e) { if (e.key === 'Enter') submit(); });

        // 送出
        btn.addEventListener('click', submit);

        function submit() {
            var val = input.value.trim();
            if (!val) { err.textContent = '請輸入密碼'; shake(); return; }

            // 管理員
            if (val === ADMIN_PWD) {
                showAdmin();
                return;
            }

            // 學員密碼
            if (val === getPwd()) {
                unlock();
                return;
            }

            err.textContent = '密碼錯誤，請重新輸入';
            input.value = '';
            shake();
        }

        function shake() {
            var card = document.getElementById('pw-card');
            card.style.animation = 'none';
            card.offsetHeight; // reflow
            card.style.animation = 'pw-shake 0.4s ease';
        }

        // 管理員面板
        function showAdmin() {
            document.getElementById('pw-title').textContent = '管理員模式';
            document.getElementById('pw-desc').textContent = '您可以修改學員密碼或直接進入簡報';
            document.getElementById('pw-main').style.display = 'none';
            document.getElementById('pw-admin').style.display = 'block';
            document.getElementById('pw-current').textContent = getPwd();
            var lockIcon = document.querySelector('#pw-card > div:first-child');
            if (lockIcon) lockIcon.innerHTML = '👑';

            document.getElementById('pw-save').addEventListener('click', function() {
                var newPwd = document.getElementById('pw-new').value.trim();
                if (!newPwd) { document.getElementById('pw-save-msg').style.color = '#e53935'; document.getElementById('pw-save-msg').textContent = '請輸入新密碼'; return; }
                if (newPwd.length < 3) { document.getElementById('pw-save-msg').style.color = '#e53935'; document.getElementById('pw-save-msg').textContent = '密碼至少 3 個字元'; return; }
                localStorage.setItem(STORAGE_KEY, newPwd);
                document.getElementById('pw-current').textContent = newPwd;
                document.getElementById('pw-new').value = '';
                document.getElementById('pw-save-msg').style.color = '#4caf50';
                document.getElementById('pw-save-msg').textContent = '密碼已更新為：' + newPwd;
            });

            document.getElementById('pw-enter').addEventListener('click', unlock);
        }

        function unlock() {
            sessionStorage.setItem(AUTH_KEY, '1');
            gate.querySelector('div').style.opacity = '0';
            gate.querySelector('div').style.transition = 'opacity 0.4s ease';
            setTimeout(function() {
                gate.remove();
                var h = document.getElementById('gate-hide');
                if (h) h.remove();
                document.documentElement.style.overflow = '';
            }, 400);
        }
    }

    // 加入動畫 CSS
    var animStyle = document.createElement('style');
    animStyle.textContent = '@keyframes pw-shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-10px)}40%,80%{transform:translateX(10px)}}';
    document.head.appendChild(animStyle);

    // 等 DOM 載入
    if (document.body) insertGate();
    else document.addEventListener('DOMContentLoaded', insertGate);
})();
