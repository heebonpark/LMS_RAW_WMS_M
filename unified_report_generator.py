# -*- coding: utf-8 -*-
"""
@file: unified_report_generator.py
@description: 물품재고 현황 리포트 + 해지고객 장비 반납 현황 리포트를
              하나의 HTML 파일, 하나의 공유 암호로 묶는 통합 리포트 생성기.

각 리포트는 inventory_html_generator.AdvancedHTMLGeneratorApp /
demolition_return_html_generator.DemolitionReturnApp의 _build_report_html()이
만드는 완전한 단일 HTML 문서를 그대로 사용합니다. 두 문서는 각자 고유한 CSS
클래스/DOM id/전역 함수(예: togglePanel, initCharts, checkAccess)를 갖고
있어서 한 페이지에 그대로 합치면 id/함수 충돌이 납니다. 그래서 각 리포트를
<iframe srcdoc="..."> 안에 통째로 담아 완전히 격리된 문서로 렌더링하고,
바깥 셸(shell) 페이지에서 한 번만 인증하면 두 iframe에 같은 비밀번호를
자동으로 입력해 잠금을 풀어주는 방식을 씁니다.
"""

import os
import html
import json
import secrets
import string
import datetime

from inventory_html_generator import ADMIN_PASSWORD


def esc(val):
    if val is None:
        return ""
    return html.escape(str(val))


def build_unified_report(inv_app, dem_app, download_dir):
    """inv_app/dem_app(이미 만들어진 두 앱 인스턴스)에서 각각 리포트 HTML을 만들어
    하나의 통합 HTML 파일로 저장합니다.

    admin 인증은 이 함수를 호출하는 쪽(자동화된 "한번에 실행" 흐름)에서 이미
    monkeypatch로 처리되므로 여기서는 추가로 묻지 않습니다 (run_sync()에 별도
    암호 확인이 없는 것과 동일한 설계).

    성공 시 (html_path, shared_user_pwd) 튜플을, 실패 시 None을 반환합니다.
    실패 원인은 각 _build_report_html() 내부에서 messagebox로 이미 안내됩니다.
    """
    if not os.path.exists(download_dir):
        download_dir = os.getcwd()

    alphabet = string.ascii_letters + string.digits
    shared_user_pwd = ''.join(secrets.choice(alphabet) for _ in range(8))

    inv_result = inv_app._build_report_html(forced_user_pwd=shared_user_pwd)
    if inv_result is None:
        return None
    inv_html, _, _ = inv_result

    dem_result = dem_app._build_report_html(forced_user_pwd=shared_user_pwd)
    if dem_result is None:
        return None
    dem_html, _, _ = dem_result

    expiry_str = getattr(inv_app, "expiry_str", None) or getattr(dem_app, "expiry_str", None)

    shell_html = _build_shell_html(inv_html, dem_html, shared_user_pwd, expiry_str)

    html_path = os.path.join(
        download_dir, f"unified_report_{datetime.date.today().strftime('%Y%m%d')}.html"
    )
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(shell_html)

    return html_path, shared_user_pwd


def _build_shell_html(inv_html, dem_html, shared_user_pwd, expiry_str):
    js_user_pwd = json.dumps(shared_user_pwd)
    js_admin_pwd = json.dumps(ADMIN_PASSWORD)
    js_expiry = json.dumps(expiry_str)
    # srcdoc 속성값으로 안전하게 넣기 위한 이스케이프 (quote=True로 " ' < > & 모두 처리)
    inv_srcdoc = html.escape(inv_html, quote=True)
    dem_srcdoc = html.escape(dem_html, quote=True)

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>'26년도 강북 / 강원본부 통합 리포트</title>
    <link rel="preconnect" href="https://cdn.jsdelivr.net">
    <link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css" rel="stylesheet">
    <style>
        :root {{
            --navy: #0b1b4d; --blue: #1d4ed8; --blue-light: #3b82f6;
            --slate: #64748b; --bg: #eef1f8; --card-border: #e6e9f2;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Pretendard', 'Malgun Gothic', 'Segoe UI', sans-serif;
            background: radial-gradient(1200px 600px at 10% -10%, #dbe4ff 0%, var(--bg) 45%), var(--bg);
            color: #1e2233; margin: 0; padding: 24px 20px; letter-spacing: -0.01em;
        }}
        #auth-overlay {{
            position: fixed; inset: 0; background: radial-gradient(900px 500px at 50% 0%, #16255e 0%, #0b1230 65%);
            display: flex; justify-content: center; align-items: center; z-index: 9999;
        }}
        .auth-card {{
            background: #ffffff; padding: 44px 40px; border-radius: 20px;
            box-shadow: 0 30px 60px -12px rgba(0,0,0,0.45); width: 400px; text-align: center;
        }}
        .auth-card h2 {{ color: var(--navy); margin-top: 0; font-size: 18pt; font-weight: 800; letter-spacing: -0.02em; }}
        .auth-card input {{
            width: 100%; padding: 13px; margin: 20px 0; border: 1.5px solid #dbe0ea; border-radius: 10px;
            box-sizing: border-box; font-size: 12pt; text-align: center; outline: none; transition: border-color .15s, box-shadow .15s;
        }}
        .auth-card input:focus {{ border-color: var(--blue-light); box-shadow: 0 0 0 4px rgba(59,130,246,0.15); }}
        .auth-card button {{
            width: 100%; padding: 13px; background: linear-gradient(135deg, var(--blue), var(--blue-light));
            color: white; border: none; border-radius: 10px; font-weight: 700; font-size: 11pt; cursor: pointer;
            box-shadow: 0 10px 20px -6px rgba(29,78,216,0.5);
        }}
        .error-msg {{ color: #ef4444; font-size: 9.5pt; margin-top: 12px; display: none; font-weight: 600; }}

        #report-content {{ display: none; max-width: 1560px; margin: 0 auto; }}
        .security-banner {{
            background: linear-gradient(90deg, #0b1230, #16255e); color: #7dd3fc; padding: 13px 22px;
            border-radius: 12px; margin-bottom: 18px; display: flex; justify-content: space-between;
            font-size: 9.5pt; font-weight: 600; align-items: center; flex-wrap: wrap; gap: 6px;
        }}
        .unified-tabs {{ display: flex; gap: 8px; margin-bottom: 16px; }}
        .unified-tab-btn {{
            flex: 1; border: 1px solid var(--card-border); background: white; color: var(--navy);
            font-size: 10.5pt; font-weight: 700; padding: 14px 16px; border-radius: 12px;
            cursor: pointer; transition: all .15s ease; box-shadow: 0 1px 2px rgba(15,23,42,0.04);
        }}
        .unified-tab-btn:hover {{ border-color: var(--blue-light); }}
        .unified-tab-btn.active {{ background: linear-gradient(135deg, var(--navy) 0%, var(--blue) 100%); border-color: var(--blue); color: white; }}

        .report-frame-wrap {{ display: none; }}
        .report-frame-wrap.active {{ display: block; }}
        .report-frame-wrap iframe {{
            width: 100%; height: calc(100vh - 200px); min-height: 640px; border: 0;
            border-radius: 16px; background: white; box-shadow: 0 20px 40px -18px rgba(15,23,42,0.15);
        }}

        @media (max-width: 880px) {{
            body {{ padding: 14px 10px; }}
            .unified-tabs {{ flex-direction: column; }}
        }}
    </style>
    <script>
        const USER_PWD = {js_user_pwd};
        const ADMIN_PWD = {js_admin_pwd};
        const EXPIRY_DATE = {js_expiry};

        function checkAccess() {{
            const inputVal = document.getElementById("password-input").value.trim();
            const now = new Date();
            const today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
            const errorElem = document.getElementById("error-message");
            if (inputVal !== ADMIN_PWD && EXPIRY_DATE && today > EXPIRY_DATE) {{
                errorElem.innerText = "조회 만료 기한(" + EXPIRY_DATE + ")이 경과하여 일반 접근이 불가합니다.";
                errorElem.style.display = "block";
                return;
            }}
            if (inputVal === USER_PWD || inputVal === ADMIN_PWD) {{
                document.getElementById("auth-overlay").style.display = "none";
                document.getElementById("report-content").style.display = "block";
            }} else {{
                errorElem.innerText = "비밀번호가 일치하지 않습니다.";
                errorElem.style.display = "block";
            }}
        }}

        function selectUnifiedTab(name, btn) {{
            document.querySelectorAll('.unified-tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.querySelectorAll('.report-frame-wrap').forEach(w => w.classList.remove('active'));
            document.getElementById('frame-wrap-' + name).classList.add('active');
        }}

        // 각 iframe은 자기 자신의 로그인 화면을 그대로 갖고 있습니다. 바깥에서 이미
        // 인증했으므로, iframe이 로드되면 같은 공유 비밀번호를 대신 입력하고
        // 그 안의 checkAccess()를 호출해 다시 묻지 않고 자동으로 잠금을 풀어줍니다.
        function autoUnlockFrame(iframe) {{
            try {{
                const doc = iframe.contentDocument || iframe.contentWindow.document;
                const input = doc.getElementById('password-input');
                const win = iframe.contentWindow;
                if (input && win && typeof win.checkAccess === 'function') {{
                    input.value = USER_PWD;
                    win.checkAccess();
                }}
            }} catch (e) {{
                // srcdoc iframe은 부모와 동일 출처라 정상적으로는 여기 오지 않습니다.
            }}
        }}
    </script>
</head>
<body>
    <div id="auth-overlay">
        <div class="auth-card">
            <h2>🔒 통합 보안 리포트 인증</h2>
            <p style="font-size:9.5pt; color:#64748b; margin-bottom:15px;">접근 암호를 입력하세요. (물품재고 · 해지고객 장비반납 공통)</p>
            <input type="password" id="password-input" placeholder="비밀번호 입력" onkeydown="if(event.key==='Enter') checkAccess();">
            <button onclick="checkAccess()">보고서 조회</button>
            <div id="error-message" class="error-msg"></div>
        </div>
    </div>

    <div id="report-content">
        <div class="security-banner">
            <span>🔒 통합 보안 리포트 (공유 암호 1개로 두 리포트 모두 조회)</span>
            <span>⏳ 조회 만료 기한: {esc(expiry_str)}까지</span>
        </div>
        <div class="unified-tabs">
            <button type="button" class="unified-tab-btn active" onclick="selectUnifiedTab('inv', this)">📦 물품재고 현황 리포트</button>
            <button type="button" class="unified-tab-btn" onclick="selectUnifiedTab('dem', this)">🔧 해지고객 장비 반납 현황 리포트</button>
        </div>
        <div class="report-frame-wrap active" id="frame-wrap-inv">
            <iframe srcdoc="{inv_srcdoc}" onload="autoUnlockFrame(this)"></iframe>
        </div>
        <div class="report-frame-wrap" id="frame-wrap-dem">
            <iframe srcdoc="{dem_srcdoc}" onload="autoUnlockFrame(this)"></iframe>
        </div>
    </div>
</body>
</html>
"""
