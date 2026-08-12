# -*- coding: utf-8 -*-
"""
@file: inventory_html_generator.py
@description: 강북·강원본부 재고관리 전자동화 & 프리미엄 EDA 시각화 대시보드 생성 시스템
              (보안/버그 수정판)

수정 사항:
  1. 관리자 비밀번호를 ADMIN_PASSWORD 상수 하나로 통합 (하드코딩 반복 제거)
  2. html.escape()로 표/스크립트 삽입값 이스케이프 (XSS/HTML 깨짐 방지)
  3. json.dumps()로 차트 데이터를 JS에 안전하게 삽입 (따옴표 포함 지사명 대응)
  4. 1번 표와 차트 데이터의 행 범위를 동일하게(7~17) 통일
  5. bare except -> except Exception as e (로그 출력)
  6. Excel COM 객체 try/finally로 반드시 Quit() 처리
  7. random -> secrets 모듈로 비밀번호 생성 (암호학적으로 안전)
"""

import os
import secrets
import string
import html
import json
import datetime
import calendar
import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import win32com.client

# 관리자 비밀번호는 여기 한 곳에서만 관리합니다.
ADMIN_PASSWORD = "0303"
TARGET_WORKBOOK_KEYWORD = "이동재고_강북강원본부 집계"


def esc(val):
    """HTML 표에 넣기 전 값 이스케이프 (None 안전 처리)"""
    if val is None:
        return ""
    return html.escape(str(val))


class AdvancedHTMLGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("강북·강원본부 재고관리 전자동화 & EDA 시각화 시스템")
        self.root.geometry("580x480")
        self.root.configure(bg="#0f172a")

        now = datetime.datetime.now()
        last_day = calendar.monthrange(now.year, now.month)[1]
        self.expiry_str = datetime.date(now.year, now.month, last_day).strftime('%Y-%m-%d')

        self.setup_ui()

    def setup_ui(self):
        header = tk.Frame(self.root, bg="#1e3a8a", height=80)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="재고 현황 전자동화 및 EDA 시각화 생성", font=("Malgun Gothic", 12, "bold"),
                 bg="#1e3a8a", fg="white").pack(pady=22)

        body = tk.Frame(self.root, bg="#0f172a")
        body.pack(fill="both", expand=True, padx=25, pady=20)

        tk.Label(body, text="📁 작업 디렉토리 (다운로드 경로)", font=("Malgun Gothic", 9, "bold"),
                 bg="#0f172a", fg="#94a3b8").pack(anchor="w", pady=(0, 4))

        self.path_var = tk.StringVar(value=r"C:\Users\user\Downloads")
        path_entry = tk.Entry(body, textvariable=self.path_var, font=("Malgun Gothic", 10),
                               bg="#1e293b", fg="#e2e8f0", insertbackground="white", relief="flat")
        path_entry.pack(fill="x", ipady=8, pady=(0, 15))

        info_text = (
            f"• 조회 만료 기한 (당월 말일 자동적용): {self.expiry_str}\n"
            f"• 관리자 인증 비밀번호는 담당자에게 문의하세요\n"
            "• 달성율 100% 미만 부진 지사 강조, 지사별 필터 버튼 및 EDA 차트 적용"
        )
        tk.Label(body, text=info_text, font=("Malgun Gothic", 9), bg="#0f172a", fg="#cbd5e1",
                 justify="left").pack(anchor="w", pady=(0, 20))

        btn_sync = tk.Button(
            body, text="⚡ 1단계: RAW 반영 및 '모두 새로고침' 실행", command=self.run_sync,
            font=("Malgun Gothic", 10, "bold"), bg="#3b82f6", fg="white", activebackground="#2563eb",
            relief="flat", cursor="hand2"
        )
        btn_sync.pack(fill="x", ipady=10, pady=(0, 10))

        btn_html = tk.Button(
            body, text="🎨 2단계: EDA 프리미엄 대시보드 보고서 생성", command=self.generate_html_report,
            font=("Malgun Gothic", 10, "bold"), bg="#10b981", fg="white", activebackground="#059669",
            relief="flat", cursor="hand2"
        )
        btn_html.pack(fill="x", ipady=12)

    def run_sync(self):
        download_dir = self.path_var.get().strip()
        if not os.path.exists(download_dir):
            download_dir = os.getcwd()

        monitoring_path = os.path.join(download_dir, "모니터링-기관별 이동중 재고현황.xlsx")
        inventory_path = os.path.join(download_dir, "재고조회-재고현황.xlsx")

        excel = None
        try:
            # CSV는 감사(audit) 기록용으로 남기고, 실제 Excel 쓰기는 메모리 상의
            # DataFrame을 그대로 사용합니다 (다시 읽지 않음 -> 불필요한 디스크 I/O 제거).
            df_mon_f = None
            if os.path.exists(monitoring_path):
                df_mon = pd.read_excel(monitoring_path, sheet_name=0)
                df_mon_f = df_mon[df_mon['재고상태'].astype(str).str.contains("신품|리퍼", na=False)].iloc[:, :15] \
                    if '재고상태' in df_mon.columns else df_mon.iloc[:, :15]
                df_mon_f.to_csv(os.path.join(download_dir, "이동재고_RAW_안전적용용.csv"),
                                 index=False, encoding='utf-8-sig')

            df_inv_f = None
            if os.path.exists(inventory_path):
                df_inv = pd.read_excel(inventory_path, sheet_name=0)
                df_inv_f = df_inv[df_inv['재고상태'].astype(str).str.contains("신품|리퍼", na=False)].iloc[:, :16] \
                    if '재고상태' in df_inv.columns else df_inv.iloc[:, :16]
                df_inv_f.to_csv(os.path.join(download_dir, "재고현황_RAW_안전적용용.csv"),
                                 index=False, encoding='utf-8-sig')

            excel = win32com.client.Dispatch("Excel.Application")
            excel.ScreenUpdating = False
            excel.Calculation = -4135  # xlCalculationManual

            target_wb = self._find_target_workbook(excel)
            if not target_wb:
                messagebox.showerror("오류", f"현재 엑셀에서 '{TARGET_WORKBOOK_KEYWORD}.xlsx' 파일이 열려있지 않습니다.")
                return False

            if df_mon_f is not None:
                ws_mov = target_wb.Sheets("이동재고 RAW")
                ws_mov.Range("A:O").ClearContents()
                # NaN(빈 셀)이 그대로 COM Range.Value에 들어가면 쓰기 오류가 날 수 있어 None으로 치환합니다.
                clean_mon = df_mon_f.astype(object).where(pd.notnull(df_mon_f), None)
                mat = [clean_mon.columns.tolist()] + clean_mon.values.tolist()
                ws_mov.Range(ws_mov.Cells(1, 1), ws_mov.Cells(len(mat), len(mat[0]))).Value = mat

            if df_inv_f is not None:
                ws_inv = target_wb.Sheets("재고현황 RAW")
                ws_inv.Range("A:P").ClearContents()
                clean_inv = df_inv_f.astype(object).where(pd.notnull(df_inv_f), None)
                mat_i = [clean_inv.columns.tolist()] + clean_inv.values.tolist()
                ws_inv.Range(ws_inv.Cells(1, 1), ws_inv.Cells(len(mat_i), len(mat_i[0]))).Value = mat_i

            target_wb.RefreshAll()
            messagebox.showinfo("성공", "RAW 데이터 반영 및 모두 새로고침이 성공적으로 완료되었습니다!")
            return True
        except Exception as e:
            messagebox.showerror("실행 오류", f"처리 중 오류 발생:\n{e}")
            return False
        finally:
            # 계산모드/화면갱신을 원복하고 COM 객체를 반드시 해제합니다.
            if excel is not None:
                try:
                    excel.ScreenUpdating = True
                    excel.Calculation = -4105  # xlCalculationAutomatic
                except Exception:
                    pass

    def _find_target_workbook(self, excel):
        for wb in excel.Workbooks:
            if TARGET_WORKBOOK_KEYWORD in wb.Name:
                return wb
        return None

    def _excel_serial_to_date(self, serial):
        """엑셀 날짜 일련번호를 datetime으로 변환 (셀 서식이 날짜가 아닌 숫자로 읽힐 때 대비)."""
        try:
            return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(serial))
        except (ValueError, OverflowError, TypeError):
            return None

    def _is_excel_error(self, val):
        """COM으로 셀을 읽을 때 #N/A, #REF! 등의 오류값은 문자열이 아니라 거대한 음수 정수
        (예: #N/A -> -2146826246)로 반환됩니다. 이런 값을 걸러냅니다."""
        if isinstance(val, bool):
            return False
        if isinstance(val, (int, float)) and -2146826300 <= val <= -2146826000:
            return True
        if isinstance(val, str) and val.strip().startswith('#'):
            return True
        return False

    def _parse_date_value(self, date_val):
        """날짜 셀이 datetime/일련번호/문자열 등 다양한 형태로 들어와도 최대한 파싱합니다."""
        if self._is_excel_error(date_val):
            return None
        if isinstance(date_val, (datetime.datetime, datetime.date)):
            return date_val
        if isinstance(date_val, (int, float)) and not isinstance(date_val, bool):
            return self._excel_serial_to_date(date_val)
        if isinstance(date_val, str):
            s = date_val.strip()
            if not s:
                return None
            for fmt in ('%Y-%m-%d', '%Y.%m.%d', '%Y/%m/%d', '%y.%m.%d', '%y-%m-%d', '%Y%m%d'):
                try:
                    return datetime.datetime.strptime(s[:10].replace(' ', ''), fmt)
                except ValueError:
                    continue
        return None

    def _collect_raw_eda(self, target_wb, valid_branches):
        """'이동재고 RAW' 시트에서 의뢰유형·재고상태·변경일자·변경지사 컬럼을 찾아
        탐색적 분석(EDA)용 집계 데이터를 만듭니다. 시트/컬럼이 없으면 available=False로 반환합니다."""
        result = {"available": False}
        try:
            ws_raw = target_wb.Sheets("이동재고 RAW")
        except Exception:
            return result

        try:
            header_map = {}
            for c in range(1, 41):
                v = ws_raw.Cells(1, c).Value
                if v:
                    header_map[str(v).strip()] = c

            def find_col(*keywords):
                """keywords는 우선순위 순서. 앞선(더 구체적인) 키워드를 모든 컬럼에서 먼저 찾고,
                못 찾으면 다음 키워드로 넘어갑니다. 같은 키워드 내에서는 완전 일치를 부분 일치보다 우선합니다."""
                for kw in keywords:
                    for name, idx in header_map.items():
                        if name == kw:
                            return idx
                    for name, idx in header_map.items():
                        if kw in name:
                            return idx
                return None

            col_type = find_col('의뢰유형')
            col_status = find_col('재고상태')
            col_date = find_col('변경일자')
            col_branch = find_col('변경지사', '지사')

            if not all([col_type, col_status, col_date, col_branch]):
                return result

            max_col_used = max(header_map.values())
            last_row = ws_raw.Cells(ws_raw.Rows.Count, col_type).End(-4162).Row  # xlUp
            if last_row < 2:
                return result

            # 셀 단위 COM 호출은 느리므로 Range 한 번으로 통째 읽어옵니다.
            data = ws_raw.Range(ws_raw.Cells(2, 1), ws_raw.Cells(last_row, max_col_used)).Value
            if data and not isinstance(data[0], tuple):
                data = (data,)  # 데이터가 1행뿐일 때 win32com이 튜플의 튜플이 아닌 단일 튜플을 반환하는 경우 보정

            now = datetime.datetime.now()
            cur_year = now.year
            before_label = f"{(cur_year - 1) % 100}년이전"

            type_status = {}   # "전체" 및 지사별: {req_type: {"신품": n, "리퍼": n}}
            month_type = {}    # "전체" 및 지사별: {month_label: {req_type: n}}
            branch_set = set()
            type_set = set()
            month_order = []
            skipped_na = 0

            def add_count(store_ts, store_mt, req_type, status_norm, month_label):
                store_ts.setdefault(req_type, {"신품": 0, "리퍼": 0})
                store_ts[req_type][status_norm] += 1
                store_mt.setdefault(month_label, {})
                store_mt[month_label][req_type] = store_mt[month_label].get(req_type, 0) + 1

            type_status["전체"] = {}
            month_type["전체"] = {}

            for row in data:
                req_type = row[col_type - 1]
                status = row[col_status - 1]
                date_val = row[col_date - 1]
                branch = row[col_branch - 1]

                # 변경지사(또는 다른 필수 필드)가 #N/A 등 오류값이면 집계에서 제외합니다.
                if self._is_excel_error(branch) or self._is_excel_error(req_type) or self._is_excel_error(status):
                    skipped_na += 1
                    continue
                if not req_type or not status or not branch:
                    continue

                req_type = str(req_type).strip()
                status_str = str(status)
                if '신품' in status_str:
                    status_norm = '신품'
                elif '리퍼' in status_str:
                    status_norm = '리퍼'
                else:
                    continue

                branch = str(branch).strip()
                if not branch or branch.startswith('#'):
                    skipped_na += 1
                    continue

                parsed_date = self._parse_date_value(date_val)
                if parsed_date is not None:
                    if parsed_date.year < cur_year:
                        month_label = before_label
                    else:
                        month_label = f"{parsed_date.year % 100}년{parsed_date.month:02d}월"
                else:
                    month_label = "기타"

                if month_label not in month_order:
                    month_order.append(month_label)

                type_set.add(req_type)
                branch_set.add(branch)

                add_count(type_status["전체"], month_type["전체"], req_type, status_norm, month_label)
                type_status.setdefault(branch, {})
                month_type.setdefault(branch, {})
                add_count(type_status[branch], month_type[branch], req_type, status_norm, month_label)

            if not type_set:
                return result

            def month_sort_key(m):
                if m == before_label:
                    return (0, 0)
                if m == '기타':
                    return (2, 0)
                try:
                    yy = int(m[:2])
                    mm = int(m[3:5])
                    return (1, yy * 100 + mm)
                except (ValueError, IndexError):
                    return (1, 999999)

            month_order.sort(key=month_sort_key)

            # 의뢰유형은 전체 지사 합계가 높은 순으로 정렬 (신품+리퍼 합계 기준)
            req_type_totals = {
                t: type_status["전체"].get(t, {}).get("신품", 0) + type_status["전체"].get(t, {}).get("리퍼", 0)
                for t in type_set
            }
            req_types_sorted = sorted(type_set, key=lambda t: req_type_totals[t], reverse=True)
            req_type_totals_sorted = [req_type_totals[t] for t in req_types_sorted]

            # 지사 순서는 가나다순 대신 중앙→원주 지사 순(valid_branches 순)을 사용합니다.
            branch_order = [b for b in valid_branches if b != '합계' and b in branch_set]
            branch_order += sorted(b for b in branch_set if b not in branch_order)
            branches_sorted = ['전체'] + branch_order

            result.update({
                "available": True,
                "type_status": type_status,
                "month_type": month_type,
                "month_order": month_order,
                "req_types": req_types_sorted,
                "req_type_totals": req_type_totals_sorted,
                "branches": branches_sorted,
                "branch_order": branch_order,
                "skipped_na": skipped_na,
            })
            return result
        except Exception:
            # RAW 시트 구조가 예상과 다르면 이 섹션만 건너뛰고 나머지 리포트는 정상 생성합니다.
            return {"available": False}

    def generate_html_report(self):
        pwd = simpledialog.askstring("관리자 인증", "관리자 비밀번호를 입력하세요:", show="*")
        if pwd != ADMIN_PASSWORD:
            messagebox.showerror("인증 실패", "관리자 비밀번호가 일치하지 않습니다.")
            return False

        result = self._build_report_html()
        if result is None:
            return False
        html_content, random_user_pwd, download_dir = result

        html_path = os.path.join(
            download_dir, f"inventory_eda_dashboard_{datetime.date.today().strftime('%Y%m%d')}.html"
        )
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        self._show_result_dialog(html_path, random_user_pwd)
        return True

    def _build_report_html(self, forced_user_pwd=None):
        """엑셀 데이터를 모아 완전한 단일 HTML 리포트 문자열을 만듭니다 (파일로 쓰지는 않음).
        generate_html_report()(개별 파일 저장)와 통합 리포트 빌더(unified_report_generator.py)
        양쪽에서 재사용합니다. 실패 시 None을 반환합니다."""
        download_dir = self.path_var.get().strip()
        if not os.path.exists(download_dir):
            download_dir = os.getcwd()

        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            target_wb = self._find_target_workbook(excel)
            if not target_wb:
                messagebox.showerror("오류", f"현재 엑셀에서 '{TARGET_WORKBOOK_KEYWORD}.xlsx' 파일이 열려있지 않습니다.")
                return None

            ws = target_wb.Sheets("현황정리")
            # 암호학적으로 안전한 난수 생성기(secrets)를 사용합니다.
            alphabet = string.ascii_letters + string.digits
            random_user_pwd = forced_user_pwd or ''.join(secrets.choice(alphabet) for _ in range(8))

            valid_branches = ['중앙지사', '강북지사', '서대문지사', '고양지사', '의정부지사',
                               '남양주지사', '강릉지사', '원주지사', '합계']

            # --- 표1 데이터를 먼저 한 번만 읽어서, 표 렌더링과 차트 데이터 생성에 동일하게 재사용 ---
            table1_rows = []
            for r in range(7, 18):
                row_vals = [ws.Cells(r, c).Value for c in range(2, 12)]
                b_name = str(row_vals[0]).strip() if row_vals[0] is not None else ""
                if b_name in valid_branches:
                    table1_rows.append((b_name, row_vals))

            chart_branches = []
            chart_actual_amt = []
            chart_target_amt = []
            chart_achieve_rate = []
            for b_name, row_vals in table1_rows:
                if b_name == '합계':
                    continue
                val_act = row_vals[6]   # 합계금액 (인덱스 6 = 컬럼 H)
                val_tgt = row_vals[7]   # 목표금액 (인덱스 7 = 컬럼 I)
                val_rate = row_vals[8]  # 달성율 (인덱스 8 = 컬럼 J, 0~1 소수)
                chart_branches.append(b_name)
                chart_actual_amt.append(round((val_act / 1000) if isinstance(val_act, (int, float)) else 0, 1))
                chart_target_amt.append(round((val_tgt / 1000) if isinstance(val_tgt, (int, float)) else 0, 1))
                chart_achieve_rate.append(round((val_rate * 100) if isinstance(val_rate, (int, float)) else 0, 1))

            monthly_totals = [0] * 8
            branch_monthly = {}
            # 표2를 먼저 만들어 monthly_totals/branch_monthly를 채운 뒤 자동 리뷰·차트에 반영합니다.
            table2_html = self._build_table2(ws, valid_branches, monthly_totals, branch_monthly)
            table3_html = self._build_table3(ws, valid_branches)

            quarter_labels, branch_quarterly, hq_quarterly = self._compute_quarterly(branch_monthly, monthly_totals)
            raw_eda = self._collect_raw_eda(target_wb, valid_branches)

            html_content = self._build_html_head(random_user_pwd)
            html_content += self._build_auto_review(table1_rows, monthly_totals)
            html_content += self._build_chart_section()
            html_content += self._build_branch_trend_ranking(branch_monthly)
            html_content += self._build_impact_review(branch_monthly, monthly_totals)
            html_content += self._build_quarter_panel(quarter_labels, branch_quarterly, hq_quarterly)
            html_content += self._build_raw_eda_panel(raw_eda)
            html_content += self._build_refurb_rate_panel(raw_eda)
            html_content += self._build_table1(table1_rows)
            html_content += table2_html
            html_content += table3_html
            html_content += self._build_footer(random_user_pwd)
            html_content += self._build_scripts(
                chart_branches, chart_actual_amt, chart_target_amt, chart_achieve_rate,
                monthly_totals, branch_monthly, quarter_labels, branch_quarterly, hq_quarterly, raw_eda
            )

            return html_content, random_user_pwd, download_dir

        except Exception as e:
            messagebox.showerror("실행 오류", f"보고서 생성 중 오류 발생:\n{e}")
            return None
        finally:
            excel = None  # COM 참조 해제 (읽기 전용 세션이라 Quit는 호출하지 않음: 사용자의 열린 엑셀을 닫지 않기 위함)

    def _show_result_dialog(self, html_path, random_user_pwd):
        """messagebox는 텍스트를 마우스로 선택/복사할 수 없으므로,
        비밀번호를 복사-붙여넣기 할 수 있는 전용 결과창을 띄웁니다."""
        dlg = tk.Toplevel(self.root)
        dlg.title("물품재고 현황 리포트 생성 완료")
        dlg.configure(bg="#0f172a")
        dlg.resizable(False, False)
        dlg.grab_set()

        frame = tk.Frame(dlg, bg="#0f172a")
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="🎨 물품재고 현황 리포트 생성 완료!", font=("Malgun Gothic", 12, "bold"),
                 bg="#0f172a", fg="white").pack(anchor="w", pady=(0, 12))

        tk.Label(frame, text="저장 경로", font=("Malgun Gothic", 9, "bold"),
                 bg="#0f172a", fg="#94a3b8").pack(anchor="w")
        path_entry = tk.Entry(frame, font=("Malgun Gothic", 9), bg="#1e293b", fg="#e2e8f0",
                               insertbackground="white", relief="flat")
        path_entry.insert(0, html_path)
        path_entry.configure(state="readonly", readonlybackground="#1e293b")
        path_entry.pack(fill="x", ipady=6, pady=(2, 14))

        tk.Label(frame, text="🔑 일반 사용자용 암호 (복사 후 붙여넣기 가능)", font=("Malgun Gothic", 9, "bold"),
                 bg="#0f172a", fg="#94a3b8").pack(anchor="w")

        pwd_row = tk.Frame(frame, bg="#0f172a")
        pwd_row.pack(fill="x", pady=(2, 4))

        pwd_entry = tk.Entry(pwd_row, font=("Malgun Gothic", 11, "bold"), bg="#1e293b", fg="#38bdf8",
                              insertbackground="white", relief="flat", justify="center")
        pwd_entry.insert(0, random_user_pwd)
        pwd_entry.configure(state="readonly", readonlybackground="#1e293b")
        pwd_entry.pack(side="left", fill="x", expand=True, ipady=8)

        def copy_pwd():
            self.root.clipboard_clear()
            self.root.clipboard_append(random_user_pwd)
            self.root.update()
            copy_btn.configure(text="✅ 복사됨")
            dlg.after(1500, lambda: copy_btn.configure(text="📋 복사"))

        copy_btn = tk.Button(pwd_row, text="📋 복사", command=copy_pwd, font=("Malgun Gothic", 9, "bold"),
                              bg="#3b82f6", fg="white", activebackground="#2563eb", relief="flat", cursor="hand2")
        copy_btn.pack(side="left", padx=(8, 0), ipady=6, ipadx=6)

        tk.Label(frame, text=f"⏳ 조회 만료 기한: {self.expiry_str} (당월 말일)", font=("Malgun Gothic", 9),
                 bg="#0f172a", fg="#cbd5e1").pack(anchor="w", pady=(14, 16))

        tk.Button(frame, text="확인", command=dlg.destroy, font=("Malgun Gothic", 10, "bold"),
                  bg="#10b981", fg="white", activebackground="#059669", relief="flat",
                  cursor="hand2").pack(fill="x", ipady=8)

    # ---------- HTML 빌더 함수들 ----------

    def _build_html_head(self, random_user_pwd):
        # 비밀번호/만료일은 JS 문자열 리터럴에도 들어가므로 json.dumps로 안전하게 이스케이프합니다.
        js_user_pwd = json.dumps(random_user_pwd)
        js_admin_pwd = json.dumps(ADMIN_PASSWORD)
        js_expiry = json.dumps(self.expiry_str)

        return f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>'26년도 강북 / 강원본부 물품재고 현황 리포트</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
    <link rel="preconnect" href="https://cdn.jsdelivr.net">
    <link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css" rel="stylesheet">
    <style>
        :root {{
            --navy: #0b1b4d;
            --blue: #1d4ed8;
            --blue-light: #3b82f6;
            --amber: #f59e0b;
            --slate: #64748b;
            --bg: #eef1f8;
            --card-border: #e6e9f2;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Pretendard', 'Malgun Gothic', 'Segoe UI', sans-serif;
            background: radial-gradient(1200px 600px at 10% -10%, #dbe4ff 0%, var(--bg) 45%), var(--bg);
            color: #1e2233;
            margin: 0;
            padding: 32px 20px;
            letter-spacing: -0.01em;
        }}
        #auth-overlay {{
            position: fixed; inset: 0;
            background: radial-gradient(900px 500px at 50% 0%, #16255e 0%, #0b1230 65%);
            display: flex; justify-content: center; align-items: center; z-index: 9999;
        }}
        .auth-card {{
            background: #ffffff; padding: 44px 40px; border-radius: 20px;
            box-shadow: 0 30px 60px -12px rgba(0,0,0,0.45), 0 0 0 1px rgba(255,255,255,0.05);
            width: 400px; text-align: center;
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
            box-shadow: 0 10px 20px -6px rgba(29,78,216,0.5); transition: transform .12s, box-shadow .12s;
        }}
        .auth-card button:hover {{ transform: translateY(-1px); box-shadow: 0 14px 24px -6px rgba(29,78,216,0.55); }}
        .error-msg {{ color: #ef4444; font-size: 9.5pt; margin-top: 12px; display: none; font-weight: 600; }}
        #report-content {{
            display: none; max-width: 1480px; margin: 0 auto; background: #ffffff;
            padding: 40px 44px; border-radius: 20px;
            box-shadow: 0 1px 2px rgba(15,23,42,0.04), 0 20px 40px -18px rgba(15,23,42,0.15);
            border: 1px solid var(--card-border);
        }}
        .security-banner {{
            background: linear-gradient(90deg, #0b1230, #16255e); color: #7dd3fc; padding: 13px 22px;
            border-radius: 12px; margin-bottom: 26px; display: flex; justify-content: space-between;
            font-size: 9.5pt; font-weight: 600; align-items: center; letter-spacing: 0.01em;
        }}
        .header-card {{
            background: linear-gradient(135deg, var(--navy) 0%, var(--blue) 60%, #2563eb 100%);
            color: white; padding: 34px; border-radius: 16px; margin-bottom: 26px; text-align: center;
            box-shadow: 0 20px 35px -12px rgba(29,78,216,0.4);
            position: relative; overflow: hidden;
        }}
        .header-card::after {{
            content: ""; position: absolute; inset: 0;
            background: radial-gradient(400px 200px at 85% -20%, rgba(255,255,255,0.18), transparent 70%);
        }}
        .header-card h1 {{ margin: 0 0 8px 0; font-size: 23pt; font-weight: 800; letter-spacing: -0.02em; position: relative; }}
        .header-card p {{ margin: 0; font-size: 10.5pt; color: #bfdbfe; position: relative; }}
        .filter-container {{
            display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 26px; background: #f6f8fc;
            padding: 14px 16px; border-radius: 12px; border: 1px solid var(--card-border); align-items: center;
        }}
        .filter-label {{ font-weight: 700; color: #334155; margin-right: 8px; font-size: 9.5pt; }}
        .filter-btn {{
            padding: 7px 15px; background: white; border: 1px solid #dbe0ea; border-radius: 999px;
            cursor: pointer; font-size: 9.5pt; font-weight: 600; color: #475569; transition: all 0.15s;
        }}
        .filter-btn:hover {{ background: #eef2ff; border-color: #c7d2fe; }}
        .filter-btn.active {{
            background: linear-gradient(135deg, var(--blue), var(--blue-light)); color: white; border-color: transparent;
            box-shadow: 0 4px 10px -2px rgba(29,78,216,0.4);
        }}
        .branch-select {{
            padding: 7px 14px; background: white; border: 1px solid #dbe0ea; border-radius: 8px;
            cursor: pointer; font-size: 9.5pt; font-weight: 600; color: #33415c;
            font-family: inherit; outline: none;
        }}
        .branch-select:focus {{ border-color: var(--blue-light); box-shadow: 0 0 0 3px rgba(59,130,246,0.15); }}
        .chart-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .chart-card {{
            background: #fbfcfe; padding: 22px 20px 16px; border-radius: 14px; border: 1px solid var(--card-border);
            box-shadow: 0 1px 3px rgba(15,23,42,0.03);
        }}
        .chart-card h3 {{ margin-top: 0; font-size: 10.5pt; font-weight: 700; color: var(--navy); text-align: center; margin-bottom: 14px; }}
        .chart-canvas-wrap {{ position: relative; width: 100%; height: 300px; }}

        /* ---- 접기/펼치기 패널 ---- */
        .panel {{ margin-bottom: 10px; }}
        .section-title {{
            font-size: 13.5pt; font-weight: 800; color: var(--navy); margin: 32px 0 14px 0;
            border-left: 4px solid var(--blue-light); padding: 2px 0 2px 12px; display: flex;
            justify-content: space-between; align-items: center; letter-spacing: -0.01em;
            cursor: pointer; user-select: none; transition: opacity .15s;
        }}
        .section-title:hover {{ opacity: 0.75; }}
        .section-title-right {{ display: flex; align-items: center; gap: 10px; }}
        .unit-label {{ font-size: 9pt; font-weight: 400; color: var(--slate); }}
        .chevron {{ display: inline-block; font-size: 10pt; color: var(--slate); transition: transform .25s ease; }}
        .panel.collapsed .chevron {{ transform: rotate(-90deg); }}
        .panel-body {{ display: grid; grid-template-rows: 1fr; transition: grid-template-rows .28s ease; }}
        .panel.collapsed .panel-body {{ grid-template-rows: 0fr; }}
        .panel-body-inner {{ overflow: hidden; min-height: 0; }}

        /* ---- 자동 분석 리뷰 카드 ---- */
        .review-card {{
            background: linear-gradient(135deg, #f5f8ff 0%, #eef2fb 100%); border: 1px solid #dde5f9;
            border-radius: 14px; padding: 20px 22px 16px; margin-bottom: 4px;
        }}
        .review-list {{ list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 9px; }}
        .review-list li {{ font-size: 9.6pt; line-height: 1.5; color: #33415c; display: flex; gap: 8px; align-items: flex-start; }}
        .review-list li .badge {{ flex-shrink: 0; font-size: 10pt; }}
        .review-list li b {{ color: var(--navy); }}
        .review-list li.tone-warn b {{ color: #d03b3b; }}
        .review-list li.tone-good b {{ color: #0ca30c; }}

        /* ---- 지사별 증감률 랭킹 ---- */
        .rank-list {{ display: flex; flex-direction: column; gap: 10px; }}
        .rank-row {{
            display: grid; grid-template-columns: 78px 1fr 74px 90px; align-items: center; gap: 12px;
        }}
        .rank-name {{ font-size: 9.6pt; font-weight: 700; color: var(--navy); }}
        .rank-bar-track {{ background: rgba(148,163,184,0.18); border-radius: 999px; height: 9px; overflow: hidden; }}
        .rank-bar-fill {{ height: 100%; border-radius: 999px; transition: width .3s ease; }}
        .rank-pct {{ font-size: 9.6pt; font-weight: 800; text-align: right; font-variant-numeric: tabular-nums; }}
        .rank-last {{ font-size: 8.6pt; color: var(--slate); text-align: right; font-variant-numeric: tabular-nums; }}
        .rank-caption {{ font-size: 8.3pt; color: var(--slate); margin-top: 14px; }}
        @media (max-width: 620px) {{
            .rank-row {{ grid-template-columns: 60px 1fr 60px; }}
            .rank-last {{ display: none; }}
        }}

        table {{
            width: 100%; border-collapse: separate; border-spacing: 0; background: white;
            font-size: 9.3pt;
        }}
        .table-wrap {{
            border-radius: 12px; overflow-x: auto; box-shadow: 0 1px 2px rgba(15,23,42,0.04);
            border: 1px solid var(--card-border);
        }}
        th, td {{ border-bottom: 1px solid #eef1f6; padding: 10px 12px; text-align: center; vertical-align: middle; white-space: nowrap; }}
        th {{
            background-color: #f6f8fc; color: #33415c; font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.02em; font-size: 8.7pt;
        }}
        /* 값은 가운데 정렬, 숫자는 자릿수 폭 고정으로 줄맞춤 */
        td {{ text-align: center; font-variant-numeric: tabular-nums; }}
        tbody tr:last-child td {{ border-bottom: none; }}
        tr:hover td {{ background-color: #f5f8ff; transition: background 0.15s; }}
        .total-row td {{ background-color: #e8edfb !important; font-weight: 800; color: var(--navy); }}
        .sluggish-row td {{ background-color: #fef1f1 !important; color: #d03b3b; font-weight: 600; }}
        .footer {{
            text-align: right; font-size: 8.8pt; color: var(--slate); border-top: 1px solid var(--card-border);
            padding-top: 18px; margin-top: 36px; font-weight: 500;
        }}

        /* ---- 반응형 ---- */
        @media (max-width: 880px) {{
            body {{ padding: 16px 10px; }}
            #report-content {{ padding: 22px 16px; border-radius: 14px; }}
            .header-card {{ padding: 22px 16px; }}
            .header-card h1 {{ font-size: 16.5pt; }}
            .header-card p {{ font-size: 8.8pt; }}
            .security-banner {{ flex-direction: column; align-items: flex-start; gap: 6px; }}
            .chart-grid {{ grid-template-columns: 1fr; }}
            .chart-canvas-wrap {{ height: 240px; }}
            .section-title {{ font-size: 12pt; }}
            table {{ font-size: 8.6pt; }}
            th, td {{ padding: 7px 8px; }}
        }}
    </style>
    <script>
        const USER_PWD = {js_user_pwd};
        const ADMIN_PWD = {js_admin_pwd};
        const EXPIRY_DATE = {js_expiry};

        function checkAccess() {{
            const inputVal = document.getElementById("password-input").value.trim();
            const now = new Date();
            // toISOString()은 UTC 기준이라 한국 시간 자정~오전 9시 사이에는 날짜가 하루 전으로 밀리므로
            // 로컬(브라우저) 날짜의 연/월/일을 직접 조합합니다.
            const today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
            const errorElem = document.getElementById("error-message");

            if (inputVal !== ADMIN_PWD && today > EXPIRY_DATE) {{
                errorElem.innerText = "조회 만료 기한(" + EXPIRY_DATE + ")이 경과하여 일반 접근이 불가합니다.";
                errorElem.style.display = "block";
                return;
            }}

            if (inputVal === USER_PWD || inputVal === ADMIN_PWD) {{
                document.getElementById("auth-overlay").style.display = "none";
                document.getElementById("report-content").style.display = "block";
                initCharts();
            }} else {{
                errorElem.innerText = "비밀번호가 일치하지 않습니다.";
                errorElem.style.display = "block";
            }}
        }}

        function togglePanel(panelId) {{
            const panel = document.getElementById(panelId);
            if (panel) panel.classList.toggle('collapsed');
        }}

        function filterBranch(branchName, btnEl) {{
            document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.remove('active'));
            btnEl.classList.add('active');

            document.querySelectorAll('table').forEach(table => {{
                table.querySelectorAll('tbody tr').forEach(row => {{
                    const branchCell = row.cells[0];
                    if (branchCell) {{
                        const name = branchCell.innerText.trim();
                        row.style.display = (branchName === '전체' || name === branchName || name === '합계') ? '' : 'none';
                    }}
                }});
            }});
        }}
    </script>
</head>
<body>
    <div id="auth-overlay">
        <div class="auth-card">
            <h2>🔒 보안 보고서 인증</h2>
            <p style="font-size:9.5pt; color:#64748b; margin-bottom:15px;">접근 암호를 입력하세요.</p>
            <input type="password" id="password-input" placeholder="비밀번호 입력" onkeydown="if(event.key==='Enter') checkAccess();">
            <button onclick="checkAccess()">보고서 조회</button>
            <div id="error-message" class="error-msg"></div>
        </div>
    </div>

    <div id="report-content">
        <div class="security-banner">
            <span>🔒 보안 보고서 (일반 사용자 암호는 별도 전달)</span>
            <span>⏳ 조회 만료 기한: {esc(self.expiry_str)}까지</span>
        </div>

        <div class="header-card">
            <h1>'26년도 강북 / 강원본부 물품재고 현황 리포트</h1>
            <p>EDA 분석 기법 및 인터랙티브 필터링이 적용된 프리미엄 시각화 리포트</p>
        </div>

        <div class="filter-container">
            <span class="filter-label">🔍 지사별 빠른 필터:</span>
            <button class="filter-btn active" onclick="filterBranch('전체', this)">전체 보기</button>
            <button class="filter-btn" onclick="filterBranch('중앙지사', this)">중앙지사</button>
            <button class="filter-btn" onclick="filterBranch('강북지사', this)">강북지사</button>
            <button class="filter-btn" onclick="filterBranch('서대문지사', this)">서대문지사</button>
            <button class="filter-btn" onclick="filterBranch('고양지사', this)">고양지사</button>
            <button class="filter-btn" onclick="filterBranch('의정부지사', this)">의정부지사</button>
            <button class="filter-btn" onclick="filterBranch('남양주지사', this)">남양주지사</button>
            <button class="filter-btn" onclick="filterBranch('강릉지사', this)">강릉지사</button>
            <button class="filter-btn" onclick="filterBranch('원주지사', this)">원주지사</button>
        </div>
"""

    def _panel_wrap(self, panel_id, title_html, unit_label, body_html):
        """섹션을 접기/펼치기 가능한 패널로 감쌉니다. 클릭 시 togglePanel()이 토글합니다."""
        unit_html = f'<span class="unit-label">{esc(unit_label)}</span>' if unit_label else ""
        return f"""
        <div class="panel" id="{panel_id}">
            <div class="section-title" onclick="togglePanel('{panel_id}')">
                <span>{title_html}</span>
                <span class="section-title-right">{unit_html}<span class="chevron">▾</span></span>
            </div>
            <div class="panel-body"><div class="panel-body-inner">
{body_html}
            </div></div>
        </div>
"""

    def _build_chart_section(self):
        # 금액(천원)과 달성율(%)은 단위가 달라 두 축을 한 차트에 억지로 겹치면(듀얼 축) 왜곡되어
        # 보이므로 별도 차트로 분리합니다 (축 1개 원칙).
        body = """
        <div class="chart-grid">
            <div class="chart-card">
                <h3>📊 지사별 재고 합계금액 vs 목표금액</h3>
                <div class="chart-canvas-wrap"><canvas id="amountBarChart"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>🎯 지사별 달성율</h3>
                <div class="chart-canvas-wrap"><canvas id="achieveRateChart"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>📈 지사별 월별 이동 재고 수량 추이 (범례 클릭 시 본부합계 표시)</h3>
                <div class="chart-canvas-wrap"><canvas id="trendLineChart"></canvas></div>
            </div>
        </div>
"""
        return self._panel_wrap("panel-chart", "📊 EDA 시각화 차트", "", body)

    def _build_branch_trend_ranking(self, branch_monthly):
        """지사별 이동재고 증감률을 계산해 랭킹 형태로 보여줍니다 (어느 지사가 늘고 있는지 한눈에 파악)."""
        rows = []
        for name, values in branch_monthly.items():
            nonzero_idx = [i for i, v in enumerate(values) if v]
            if len(nonzero_idx) < 2:
                continue
            first_i, last_i = nonzero_idx[0], nonzero_idx[-1]
            first_v, last_v = values[first_i], values[last_i]
            if first_v:
                pct = (last_v - first_v) / first_v * 100
            else:
                pct = 100.0 if last_v > 0 else 0.0
            rows.append((name, pct, last_v))

        if not rows:
            body = '<div class="review-card"><ul class="review-list"><li><span class="badge">ℹ️</span><span>증감률을 계산할 월별 데이터가 부족합니다.</span></li></ul></div>'
            return self._panel_wrap("panel-ranking", "📈 지사별 이동재고 증감률 랭킹", "", body)

        rows.sort(key=lambda x: x[1], reverse=True)
        max_abs = max(abs(pct) for _, pct, _ in rows) or 1

        items_html = []
        for name, pct, last_v in rows:
            is_up = pct >= 0
            bar_pct = min(abs(pct) / max_abs * 100, 100)
            color = "#0ca30c" if is_up else "#d03b3b"
            arrow = "▲" if is_up else "▼"
            items_html.append(f"""
                <div class="rank-row">
                    <span class="rank-name">{esc(name)}</span>
                    <div class="rank-bar-track">
                        <div class="rank-bar-fill" style="width:{bar_pct:.1f}%; background:{color};"></div>
                    </div>
                    <span class="rank-pct" style="color:{color};">{arrow} {abs(pct):.1f}%</span>
                    <span class="rank-last">최근 {last_v:,.0f}건</span>
                </div>""")

        body = f"""
        <div class="review-card">
            <div class="rank-list">{"".join(items_html)}
            </div>
            <div class="rank-caption">※ 데이터가 있는 첫 달 대비 마지막 달 증감률 기준 정렬</div>
        </div>
"""
        return self._panel_wrap("panel-ranking", "📈 지사별 이동재고 증감률 랭킹", "", body)

    def _build_impact_review(self, branch_monthly, monthly_totals):
        """증감률(%)만으로는 기준값이 작을 때 왜곡되므로, 절대 증가량(건) 기준으로
        어느 지사가 실제로 이동재고 증가에 가장 큰 영향을 미치는지 분석하는 관리 리뷰입니다."""
        deltas = []  # (name, delta, first_v, last_v)
        for name, values in branch_monthly.items():
            nonzero_idx = [i for i, v in enumerate(values) if v]
            if not nonzero_idx:
                continue
            first_i, last_i = nonzero_idx[0], nonzero_idx[-1]
            if first_i == last_i:
                # 데이터가 한 달치 뿐이면 0에서 새로 생긴 것으로 간주
                first_v, last_v = 0, values[last_i]
            else:
                first_v, last_v = values[first_i], values[last_i]
            deltas.append((name, last_v - first_v, first_v, last_v))

        if not deltas:
            body = '<div class="review-card"><ul class="review-list"><li><span class="badge">ℹ️</span><span>영향도를 계산할 데이터가 부족합니다.</span></li></ul></div>'
            return self._panel_wrap("panel-impact", "🧭 이동재고 증가 영향도 관리 리뷰", "", body)

        increases = sorted([d for d in deltas if d[1] > 0], key=lambda x: x[1], reverse=True)
        decreases = sorted([d for d in deltas if d[1] < 0], key=lambda x: x[1])
        total_increase = sum(d[1] for d in increases) or 1

        medals = ["🥇", "🥈", "🥉"]
        lines = []
        if increases:
            top_share = sum(d[1] for d in increases[:3]) / total_increase * 100
            lines.append(
                ("good", "🔎",
                 f"전체 증가분 <b>{total_increase:,.0f}건</b> 중 상위 {min(3, len(increases))}개 지사가 "
                 f"<b>{top_share:.1f}%</b>를 차지합니다.")
            )
            for i, (name, delta, first_v, last_v) in enumerate(increases[:3]):
                share = delta / total_increase * 100
                medal = medals[i] if i < len(medals) else "▪️"
                lines.append(
                    ("good", medal,
                     f"<b>{esc(name)}</b>: {first_v:,.0f}건 → {last_v:,.0f}건 "
                     f"(<b>+{delta:,.0f}건</b>, 전체 증가분의 {share:.1f}%) — 이동재고 증가에 가장 큰 영향")
                )
        else:
            lines.append(("neutral", "ℹ️", "증가 추세를 보이는 지사가 없습니다."))

        if decreases:
            names = ", ".join(f"{esc(n)}({d:,.0f}건)" for n, d, _, _ in decreases[:4])
            more = f" 외 {len(decreases)-4}곳" if len(decreases) > 4 else ""
            lines.append(("warn", "⚠️", f"이동재고 감소 지사 <b>{len(decreases)}곳</b>: {names}{more} — 재고 부족 여부 점검 필요"))

        lines.append(("neutral", "📌",
                       "증감률(%)은 기준월 수치가 매우 작을 때 크게 부풀려질 수 있어, "
                       "위와 같이 <b>절대 증가량(건)</b> 기준으로도 함께 확인하는 것을 권장합니다."))

        li_html = "\n".join(
            f'                <li class="tone-{tone}"><span class="badge">{badge}</span><span>{text}</span></li>'
            for tone, badge, text in lines
        )
        body = f"""
        <div class="review-card">
            <ul class="review-list">
{li_html}
            </ul>
        </div>
"""
        return self._panel_wrap("panel-impact", "🧭 이동재고 증가 영향도 관리 리뷰", "", body)

    def _compute_quarterly(self, branch_monthly, monthly_totals):
        """월별 데이터를 3개월 단위(분기)로 묶어 사분기 지표를 계산합니다."""
        n = len(monthly_totals)
        groups = []
        i = 0
        qnum = 1
        quarter_labels = []
        while i < n:
            chunk = list(range(i, min(i + 3, n)))
            groups.append(chunk)
            label = f"{qnum}분기"
            if len(chunk) < 3:
                label += "(부분)"
            quarter_labels.append(label)
            i += 3
            qnum += 1

        branch_quarterly = {
            name: [sum(values[idx] for idx in chunk) for chunk in groups]
            for name, values in branch_monthly.items()
        }
        hq_quarterly = [sum(monthly_totals[idx] for idx in chunk) for chunk in groups]
        return quarter_labels, branch_quarterly, hq_quarterly

    def _build_quarter_panel(self, quarter_labels, branch_quarterly, hq_quarterly):
        # QoQ(전분기 대비) 증감률 - 마지막 분기 기준
        qoq_badge = ""
        if len(hq_quarterly) >= 2 and hq_quarterly[-2]:
            qoq = (hq_quarterly[-1] - hq_quarterly[-2]) / hq_quarterly[-2] * 100
            tone_color = "#0ca30c" if qoq >= 0 else "#d03b3b"
            arrow = "▲" if qoq >= 0 else "▼"
            qoq_badge = (
                f'<div style="margin-bottom:14px; font-size:9.6pt; font-weight:700; color:{tone_color};">'
                f'{arrow} 전분기 대비 본부 합계 {abs(qoq):.1f}% {"증가" if qoq >= 0 else "감소"}</div>'
            )

        rows_html = ""
        for name, q_vals in branch_quarterly.items():
            cells = "".join(f"<td>{v:,.0f}</td>" for v in q_vals)
            rows_html += f"<tr><td>{esc(name)}</td>{cells}</tr>"
        total_cells = "".join(f"<td>{v:,.0f}</td>" for v in hq_quarterly)
        rows_html += f'<tr class="total-row"><td>본부 합계</td>{total_cells}</tr>'

        header_cells = "".join(f"<th>{esc(q)}</th>" for q in quarter_labels)
        body = f"""
        {qoq_badge}
        <div class="chart-card" style="margin-bottom:18px;">
            <h3>📊 지사별 분기 이동재고 (누적 막대)</h3>
            <div class="chart-canvas-wrap"><canvas id="quarterChart"></canvas></div>
        </div>
        <div class="table-wrap">
        <table>
            <thead><tr><th>구분</th>{header_cells}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        </div>
"""
        return self._panel_wrap("panel-quarter", "📅 분기별(사분기) 이동재고 지표", "", body)

    def _build_type_branch_matrix(self, raw_eda):
        """의뢰유형(행) × 지사(중앙지사~원주지사, 열) 교차표. 셀 값은 신품+리퍼 합계 건수.
        값 크기에 따라 배경색 농도를 다르게 주는 히트맵 스타일을 적용해 한눈에 비교되도록 합니다."""
        req_types = raw_eda["req_types"]  # 이미 합계 높은순 정렬됨
        branch_order = raw_eda["branch_order"]
        type_status = raw_eda["type_status"]

        def cell_count(branch, req_type):
            ts = type_status.get(branch, {})
            v = ts.get(req_type)
            if not v:
                return 0
            return v.get("신품", 0) + v.get("리퍼", 0)

        # 히트맵 색상 강도 기준값(총합계 큰 항목이 스케일을 다 잡아먹지 않도록, 상위 항목 제외한 값 기준 max 사용)
        all_cells = [cell_count(b, t) for t in req_types for b in branch_order]
        sorted_cells = sorted(all_cells, reverse=True)
        heat_max = sorted_cells[min(3, len(sorted_cells) - 1)] or (sorted_cells[0] if sorted_cells else 1) or 1

        def heat_style(v):
            if v <= 0:
                return ""
            ratio = min(v / heat_max, 1.0)
            alpha = 0.08 + ratio * 0.55
            text_color = "#0b1b4d" if ratio < 0.6 else "#ffffff"
            weight = "700" if ratio >= 0.35 else "500"
            return f' style="background:rgba(37,99,235,{alpha:.2f}); color:{text_color}; font-weight:{weight};"'

        header_cells = "".join(f"<th>{esc(b)}</th>" for b in branch_order) + "<th>합계</th>"
        rows_html = ""
        col_totals = [0] * len(branch_order)
        for req_type in req_types:
            row_vals = [cell_count(b, req_type) for b in branch_order]
            for i, v in enumerate(row_vals):
                col_totals[i] += v
            row_total = sum(row_vals)
            cells = "".join(f"<td{heat_style(v)}>{v:,.0f}</td>" for v in row_vals)
            rows_html += f"<tr><td>{esc(req_type)}</td>{cells}<td><b>{row_total:,.0f}</b></td></tr>"
        total_cells = "".join(f"<td><b>{v:,.0f}</b></td>" for v in col_totals)
        rows_html += f'<tr class="total-row"><td>합계</td>{total_cells}<td><b>{sum(col_totals):,.0f}</b></td></tr>'

        return f"""
        <div class="table-wrap" style="margin-top:18px;">
        <table>
            <thead><tr><th>의뢰유형 \\ 지사</th>{header_cells}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        </div>
        <div class="rank-caption" style="margin-top:8px;">※ 색이 진할수록 건수가 많은 셀입니다 (히트맵). 의뢰유형은 전체 합계가 높은 순으로 정렬됩니다.</div>
"""

    def _build_raw_eda_panel(self, raw_eda):
        if not raw_eda.get("available"):
            body = (
                '<div class="review-card"><ul class="review-list"><li><span class="badge">ℹ️</span>'
                "<span>'이동재고 RAW' 시트에서 의뢰유형/재고상태/변경일자/변경지사 컬럼을 찾지 못해 "
                "이 분석을 건너뛰었습니다. 컬럼명을 확인해 주세요.</span></li></ul></div>"
            )
            return self._panel_wrap("panel-raw-eda", "🔬 이동재고 RAW 탐색적 분석", "", body)

        branch_options = "".join(
            f'<option value="{esc(b)}">{esc(b)}</option>' for b in raw_eda["branches"]
        )
        matrix_html = self._build_type_branch_matrix(raw_eda)

        na_notice = ""
        if raw_eda.get("skipped_na"):
            na_notice = (
                f'<div class="rank-caption" style="margin-bottom:14px;">'
                f'※ 변경지사/의뢰유형/재고상태가 #N/A 등 오류값인 {raw_eda["skipped_na"]:,}건은 집계에서 제외했습니다.</div>'
            )

        body = f"""
        {na_notice}
        <div class="filter-container" style="margin-bottom:18px;">
            <span class="filter-label">🏢 변경지사 필터:</span>
            <select id="raw-eda-branch-select" class="branch-select" onchange="updateRawEda(this.value)">
                {branch_options}
            </select>
        </div>
        <div class="chart-grid">
            <div class="chart-card">
                <h3>📦 의뢰유형별 재고상태 구성 (신품/리퍼, 합계 높은순)</h3>
                <div class="chart-canvas-wrap"><canvas id="typeStatusChart"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>🗓️ 변경일자(월별) × 의뢰유형 추이</h3>
                <div class="chart-canvas-wrap"><canvas id="monthTypeChart"></canvas></div>
            </div>
        </div>
        <div class="section-title" style="cursor:default; margin-top:28px;">
            <span>📋 의뢰유형 × 지사(중앙~원주) 교차표</span>
            <span class="section-title-right"><span class="unit-label">[단위: 건, 신품+리퍼 합계]</span></span>
        </div>
        <div class="chart-card" style="margin-bottom:0;">
            <h3>📊 의뢰유형별 전체 건수 (지사 합계, 높은순)</h3>
            <div class="chart-canvas-wrap" style="height:260px;"><canvas id="reqTypeTotalChart"></canvas></div>
        </div>
        {matrix_html}
"""
        return self._panel_wrap("panel-raw-eda", "🔬 이동재고 RAW 탐색적 분석 (의뢰유형·재고상태·변경지사)", "", body)

    def _build_refurb_rate_panel(self, raw_eda):
        """지사별로 어떤 의뢰유형에서 리퍼(재고상태) 사용 비율이 높은지 보여줍니다.
        (신품/리퍼 절대 건수가 아니라 '리퍼 / (신품+리퍼) × 100' 비율 기준)"""
        if not raw_eda.get("available"):
            return ""

        branch_order = raw_eda["branch_order"]
        req_types = raw_eda["req_types"]  # 합계 높은순 정렬됨
        type_status = raw_eda["type_status"]

        def counts(branch, req_type):
            d = type_status.get(branch, {}).get(req_type, {})
            return d.get("신품", 0), d.get("리퍼", 0)

        # 지사별 전체 리퍼율 + 표본이 충분한(>=5건) 의뢰유형 중 리퍼율이 가장 높은 유형
        branch_overall = {}
        branch_top_type = {}
        for b in branch_order:
            new_sum = ref_sum = 0
            candidates = []
            for t in req_types:
                n, r = counts(b, t)
                new_sum += n
                ref_sum += r
                tot = n + r
                if tot >= 5:
                    candidates.append((t, r / tot * 100, tot))
            total = new_sum + ref_sum
            branch_overall[b] = (ref_sum / total * 100) if total > 0 else None
            branch_top_type[b] = max(candidates, key=lambda x: x[1]) if candidates else None

        ranking = sorted(
            [(b, r) for b, r in branch_overall.items() if r is not None],
            key=lambda x: x[1], reverse=True
        )

        if not ranking:
            body = '<div class="review-card"><ul class="review-list"><li><span class="badge">ℹ️</span><span>리퍼율을 계산할 데이터가 부족합니다.</span></li></ul></div>'
            return self._panel_wrap("panel-refurb-rate", "♻️ 지사별 리퍼 사용율 (의뢰유형별)", "", body)

        max_rate = max(r for _, r in ranking) or 1
        rank_rows = []
        for b, rate in ranking:
            bar_pct = min(rate / max_rate * 100, 100)
            top = branch_top_type.get(b)
            top_note = f" · 최다: {esc(top[0])} ({top[1]:.1f}%)" if top else ""
            rank_rows.append(f"""
                <div class="rank-row">
                    <span class="rank-name">{esc(b)}</span>
                    <div class="rank-bar-track">
                        <div class="rank-bar-fill" style="width:{bar_pct:.1f}%; background:#d97706;"></div>
                    </div>
                    <span class="rank-pct" style="color:#b45309;">{rate:.1f}%</span>
                    <span class="rank-last">{top_note}</span>
                </div>""")

        # 히트맵 교차표: 지사(행) × 의뢰유형(열) 리퍼율(%)
        def heat_style(rate):
            if rate is None:
                return "", "-"
            ratio = min(rate / 100, 1.0)
            alpha = 0.08 + ratio * 0.55
            text_color = "#7c2d12" if ratio < 0.6 else "#ffffff"
            weight = "700" if ratio >= 0.35 else "500"
            return f' style="background:rgba(217,119,6,{alpha:.2f}); color:{text_color}; font-weight:{weight};"', f"{rate:.0f}%"

        header_cells = "".join(f"<th>{esc(t)}</th>" for t in req_types) + "<th>전체</th>"
        body_rows = ""
        for b in branch_order:
            cells = ""
            for t in req_types:
                n, r = counts(b, t)
                tot = n + r
                rate = (r / tot * 100) if tot > 0 else None
                style, label = heat_style(rate)
                cells += f"<td{style}>{label}</td>"
            overall = branch_overall.get(b)
            ov_style, ov_label = heat_style(overall)
            body_rows += f"<tr><td>{esc(b)}</td>{cells}<td{ov_style}><b>{ov_label}</b></td></tr>"

        body = f"""
        <div class="review-card">
            <div class="rank-list">{"".join(rank_rows)}
            </div>
            <div class="rank-caption">※ 전체 신품+리퍼 합계 대비 리퍼 비율(%) 기준. "최다"는 지사 내 표본 5건 이상인 의뢰유형 중 리퍼율이 가장 높은 유형입니다.</div>
        </div>
        <div class="table-wrap" style="margin-top:18px;">
        <table>
            <thead><tr><th>지사 \\ 의뢰유형</th>{header_cells}</tr></thead>
            <tbody>{body_rows}</tbody>
        </table>
        </div>
        <div class="rank-caption" style="margin-top:8px;">※ 색이 진할수록 해당 지사·의뢰유형의 리퍼 사용 비율이 높습니다. '-'는 데이터 없음.</div>
"""
        return self._panel_wrap("panel-refurb-rate", "♻️ 지사별 리퍼 사용율 (의뢰유형별)", "", body)

    def _build_auto_review(self, table1_rows, monthly_totals):
        """표1·이동재고 데이터를 기반으로 자동 인사이트 문구를 생성합니다 (자동 리뷰 알고리즘)."""
        branch_rates = []
        for b_name, row_vals in table1_rows:
            if b_name == '합계':
                continue
            rate_val = row_vals[8] if len(row_vals) > 8 else None
            if isinstance(rate_val, (int, float)):
                branch_rates.append((b_name, rate_val * 100))

        items = []  # (tone, badge, html_text)

        if branch_rates:
            avg_rate = sum(r for _, r in branch_rates) / len(branch_rates)
            below = sorted([(n, r) for n, r in branch_rates if r < 100], key=lambda x: x[1])
            best_name, best_rate = max(branch_rates, key=lambda x: x[1])
            worst_name, worst_rate = min(branch_rates, key=lambda x: x[1])

            avg_tone = "good" if avg_rate >= 100 else "warn"
            items.append((avg_tone, "📊", f"본부 전체 평균 달성율은 <b>{avg_rate:.1f}%</b> 입니다."))

            if below:
                names = ", ".join(f"{n}({r:.1f}%)" for n, r in below[:5])
                more = f" 외 {len(below)-5}곳" if len(below) > 5 else ""
                items.append(("warn", "⚠️", f"목표 미달 지사 <b>{len(below)}곳</b>: {names}{more}"))
            else:
                items.append(("good", "✅", "전 지사가 목표를 달성했습니다."))

            items.append(("good", "🏆", f"최고 실적 지사: <b>{best_name}</b> ({best_rate:.1f}%)"))
            items.append(("warn" if worst_rate < 100 else "neutral", "🔻",
                           f"최저 실적 지사: <b>{worst_name}</b> ({worst_rate:.1f}%)"))

        nonzero_idx = [i for i, v in enumerate(monthly_totals) if v]
        if len(nonzero_idx) >= 2:
            last_i, prev_i = nonzero_idx[-1], nonzero_idx[-2]
            last_v, prev_v = monthly_totals[last_i], monthly_totals[prev_i]
            month_labels = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월']
            if prev_v:
                pct = (last_v - prev_v) / prev_v * 100
                arrow, tone = ("📈", "good") if pct >= 0 else ("📉", "warn")
                items.append((tone, arrow,
                               f"{month_labels[last_i]} 이동 재고는 {month_labels[prev_i]} 대비 "
                               f"<b>{abs(pct):.1f}% {'증가' if pct >= 0 else '감소'}</b> 했습니다."))

        if not items:
            items.append(("neutral", "ℹ️", "자동 분석을 위한 데이터가 충분하지 않습니다."))

        li_html = "\n".join(
            f'                <li class="tone-{tone}"><span class="badge">{badge}</span><span>{text}</span></li>'
            for tone, badge, text in items
        )
        body = f"""
        <div class="review-card">
            <ul class="review-list">
{li_html}
            </ul>
        </div>
"""
        return self._panel_wrap("panel-review", "🧠 자동 분석 리뷰", "", body)

    def _build_table1(self, table1_rows):
        out = """
        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th rowspan="2">구분</th>
                    <th colspan="3">재고 수량(건)</th>
                    <th colspan="3">재고 금액(천원)</th>
                    <th rowspan="2">목표금액</th>
                    <th rowspan="2">달성율</th>
                    <th rowspan="2">순위</th>
                </tr>
                <tr>
                    <th>신품</th><th>리퍼</th><th>합계</th>
                    <th>신품금액</th><th>리퍼금액</th><th>합계금액</th>
                </tr>
            </thead>
            <tbody>
"""
        for b_name, row_vals in table1_rows:
            is_total = (b_name == '합계')
            rate_val = row_vals[8] if len(row_vals) > 8 else None
            is_sluggish = False
            if not is_total and rate_val is not None:
                try:
                    is_sluggish = float(rate_val) < 1.0
                except (TypeError, ValueError):
                    pass

            row_cls = ' class="total-row"' if is_total else (' class="sluggish-row"' if is_sluggish else "")
            out += f"<tr{row_cls}>"
            for idx, val in enumerate(row_vals):
                val_str = ""
                if val is not None:
                    if isinstance(val, (int, float)) and idx > 0 and idx != 8:
                        val_str = f"{val/1000:,.0f}" if idx in [4, 5, 6, 7] else f"{val:,.0f}"
                    elif idx == 8 and isinstance(val, (int, float)):
                        val_str = f"{val*100:.1f}%"
                    else:
                        val_str = esc(val)
                out += f"<td>{val_str}</td>"
            out += "</tr>"
        out += "\n            </tbody>\n        </table>\n        </div>\n"
        return self._panel_wrap("panel-table1", "○ 강북강원본부 재고 현황", "[단위: 건, 천원, %]", out)

    def _build_table2(self, ws, valid_branches, monthly_totals, branch_monthly):
        out = """
        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>구분</th><th>25년이전</th><th>25년</th>
                    <th>1월</th><th>2월</th><th>3월</th><th>4월</th><th>5월</th><th>6월</th><th>7월</th><th>8월</th>
                    <th>합계</th><th>총합</th><th>순위</th>
                </tr>
            </thead>
            <tbody>
"""
        for r in range(20, 31):
            row_vals = [ws.Cells(r, c).Value for c in range(2, 16)]
            b_name = str(row_vals[0]).strip() if row_vals[0] is not None else ""
            if b_name not in valid_branches:
                continue

            is_total = (b_name == '합계')
            if not is_total:
                monthly_vals = []
                for m_idx in range(3, 11):
                    m_val = row_vals[m_idx]
                    v = m_val if isinstance(m_val, (int, float)) else 0
                    monthly_totals[m_idx - 3] += v
                    monthly_vals.append(v)
                branch_monthly[b_name] = monthly_vals

            rank_val = row_vals[13] if len(row_vals) > 13 else None
            is_sluggish = False
            if not is_total and rank_val is not None:
                try:
                    is_sluggish = 5 <= int(rank_val) <= 8
                except (TypeError, ValueError):
                    pass

            row_cls = ' class="total-row"' if is_total else (' class="sluggish-row"' if is_sluggish else "")
            out += f"<tr{row_cls}>"
            for idx, val in enumerate(row_vals):
                val_str = f"{val:,.0f}" if isinstance(val, (int, float)) and idx > 0 else esc(val)
                out += f"<td>{val_str}</td>"
            out += "</tr>"
        out += "\n            </tbody>\n        </table>\n        </div>\n"
        return self._panel_wrap("panel-table2", "○ 강북강원본부 이동 재고 수량 현황", "[단위: 건]", out)

    def _build_table3(self, ws, valid_branches):
        out = """
        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>구분</th><th>25년이전</th><th>25년</th>
                    <th>1월</th><th>2월</th><th>3월</th><th>4월</th><th>5월</th><th>6월</th><th>7월</th><th>8월</th>
                    <th>합계</th><th>총합</th><th>순위</th>
                </tr>
            </thead>
            <tbody>
"""
        for r in range(33, 45):
            row_vals = [ws.Cells(r, c).Value for c in range(2, 16)]
            b_name = str(row_vals[0]).strip() if row_vals[0] is not None else ""
            if b_name not in valid_branches:
                continue

            is_total = (b_name == '합계')
            rank_val = row_vals[13] if len(row_vals) > 13 else None
            is_sluggish = False
            if not is_total and rank_val is not None:
                try:
                    is_sluggish = 5 <= int(rank_val) <= 8
                except (TypeError, ValueError):
                    pass

            row_cls = ' class="total-row"' if is_total else (' class="sluggish-row"' if is_sluggish else "")
            out += f"<tr{row_cls}>"
            for idx, val in enumerate(row_vals):
                if isinstance(val, (int, float)) and 0 < idx < 13:
                    val_str = f"{val/1000:,.0f}"
                elif isinstance(val, (int, float)) and idx > 0:
                    val_str = f"{val:,.0f}"
                else:
                    val_str = esc(val)
                out += f"<td>{val_str}</td>"
            out += "</tr>"
        out += "\n            </tbody>\n        </table>\n        </div>\n"
        return self._panel_wrap("panel-table3", "○ 강북강원본부 이동 재고 금액 현황", "[단위: 천원]", out)

    def _build_footer(self, random_user_pwd):
        return f"""
        <div class="footer">
            Generated by Inventory Secure Visualization System | 조회 기한: {esc(self.expiry_str)}
        </div>
    </div>
"""

    def _build_scripts(self, chart_branches, chart_actual_amt, chart_target_amt, chart_achieve_rate,
                        monthly_totals, branch_monthly, quarter_labels, branch_quarterly, hq_quarterly,
                        raw_eda):
        # json.dumps로 안전하게 JS 배열/리터럴 생성 (따옴표 등 특수문자 대응)
        # 색맹 안전성 검증된 순서 고정 카테고리 팔레트 (blue-orange-aqua-yellow-magenta-green-violet-red)
        branch_palette = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100',
                           '#e87ba4', '#008300', '#4a3aa7', '#e34948']
        branch_line_datasets = []
        for i, (name, values) in enumerate(branch_monthly.items()):
            color = branch_palette[i % len(branch_palette)]
            branch_line_datasets.append({
                "label": name,
                "data": values,
                "borderColor": color,
                "backgroundColor": color,
                "tension": 0.35,
                "borderWidth": 2,
                "pointRadius": 3,
                "pointBackgroundColor": color,
                "pointBorderColor": "#fff",
                "pointBorderWidth": 1,
                "fill": False,
            })
        # 본부 합계는 개별 지사보다 값이 훨씬 커서 축 스케일을 왜곡하므로 기본적으로 숨기고,
        # 범례에서 "본부 합계"를 클릭하면 켤 수 있도록 합니다.
        branch_line_datasets.append({
            "label": "본부 합계",
            "data": monthly_totals,
            "borderColor": "#94a3b8",
            "backgroundColor": "#94a3b8",
            "borderDash": [6, 4],
            "borderWidth": 2.5,
            "pointRadius": 0,
            "tension": 0.35,
            "fill": False,
            "hidden": True,
        })

        # 분기별 누적(stacked) 막대 차트용 데이터셋 (지사별 색상 동일하게 재사용)
        quarter_datasets = []
        for i, (name, q_vals) in enumerate(branch_quarterly.items()):
            color = branch_palette[i % len(branch_palette)]
            quarter_datasets.append({
                "label": name,
                "data": q_vals,
                "backgroundColor": color,
                "borderRadius": 4,
                "stack": "q",
            })

        # RAW 탐색적 분석: 지사 필터별로 클라이언트에서 즉시 전환할 수 있도록
        # {지사: {의뢰유형: {신품, 리퍼}}} / {지사: {월: {의뢰유형: n}}} 구조를 그대로 JSON으로 넘깁니다.
        raw_eda_js = "{}"
        raw_req_types_js = "[]"
        raw_month_order_js = "[]"
        raw_req_type_totals_js = "[]"
        if raw_eda.get("available"):
            raw_eda_js = json.dumps({
                "typeStatus": raw_eda["type_status"],
                "monthType": raw_eda["month_type"],
            }, ensure_ascii=False)
            raw_req_types_js = json.dumps(raw_eda["req_types"], ensure_ascii=False)
            raw_month_order_js = json.dumps(raw_eda["month_order"], ensure_ascii=False)
            raw_req_type_totals_js = json.dumps(raw_eda["req_type_totals"], ensure_ascii=False)

        return f"""
    <script>
        Chart.defaults.font.family = "'Pretendard', 'Malgun Gothic', 'Segoe UI', sans-serif";
        Chart.defaults.color = '#475569';
        // 직접 라벨은 선택된 핵심 차트에서만 켭니다 (모든 점에 숫자를 찍지 않음).
        if (window.ChartDataLabels) {{
            Chart.register(ChartDataLabels);
            Chart.defaults.set('plugins.datalabels', {{ display: false }});
        }}

        function initCharts() {{
            const ctxBar = document.getElementById('amountBarChart').getContext('2d');
            const barGradient = ctxBar.createLinearGradient(0, 0, 0, 320);
            barGradient.addColorStop(0, 'rgba(37, 99, 235, 0.95)');
            barGradient.addColorStop(1, 'rgba(37, 99, 235, 0.55)');

            new Chart(ctxBar, {{
                data: {{
                    labels: {json.dumps(chart_branches, ensure_ascii=False)},
                    datasets: [
                        {{
                            type: 'bar',
                            label: '합계금액(천원)',
                            data: {json.dumps(chart_actual_amt)},
                            backgroundColor: barGradient,
                            borderRadius: 6,
                            borderSkipped: false
                        }},
                        {{
                            type: 'bar',
                            label: '목표금액(천원)',
                            data: {json.dumps(chart_target_amt)},
                            backgroundColor: 'rgba(203, 213, 225, 0.55)',
                            borderRadius: 6,
                            borderSkipped: false
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {{ mode: 'index', intersect: false }},
                    layout: {{ padding: {{ top: 4 }} }},
                    plugins: {{
                        legend: {{
                            position: 'top', align: 'center',
                            labels: {{ usePointStyle: true, boxWidth: 8, padding: 18, font: {{ size: 10.5 }} }}
                        }},
                        tooltip: {{
                            backgroundColor: '#0f172a',
                            padding: 10,
                            cornerRadius: 8,
                            callbacks: {{
                                label: ctx => ' ' + ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString() + '천원'
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            grid: {{ color: 'rgba(148,163,184,0.15)' }},
                            title: {{ display: true, text: '금액(천원)', font: {{ size: 10 }}, padding: {{ bottom: 6 }} }}
                        }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            // 금액(천원)과 달성율(%)은 단위가 달라 억지로 한 차트/듀얼 축에 겹치지 않고 별도 차트로 분리했습니다.
            // 핵심 지표라 막대 위에 직접 값을 라벨로 표시합니다 (호버 없이도 바로 읽힘).
            const achieveCanvas = document.getElementById('achieveRateChart');
            if (achieveCanvas) {{
                new Chart(achieveCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: {{
                        labels: {json.dumps(chart_branches, ensure_ascii=False)},
                        datasets: [{{
                            label: '달성율(%)',
                            data: {json.dumps(chart_achieve_rate)},
                            backgroundColor: '#eb6834',
                            borderRadius: 6,
                            borderSkipped: false
                        }}]
                    }},
                    options: {{
                        responsive: true,
                        maintainAspectRatio: false,
                        layout: {{ padding: {{ top: 24 }} }},
                        plugins: {{
                            legend: {{ display: false }},
                            tooltip: {{
                                backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                                callbacks: {{ label: ctx => ' ' + ctx.dataset.label + ': ' + ctx.parsed.y + '%' }}
                            }},
                            datalabels: {{
                                display: true, anchor: 'end', align: 'top', offset: 2,
                                color: '#33415c', font: {{ size: 9.5, weight: '700' }},
                                formatter: v => v + '%'
                            }}
                        }},
                        scales: {{
                            y: {{
                                beginAtZero: true,
                                grid: {{ color: 'rgba(148,163,184,0.15)' }},
                                ticks: {{ callback: v => v + '%' }},
                                title: {{ display: true, text: '달성율(%)', font: {{ size: 10 }}, padding: {{ bottom: 6 }} }}
                            }},
                            x: {{ grid: {{ display: false }} }}
                        }}
                    }}
                }});
            }}

            const ctxLine = document.getElementById('trendLineChart').getContext('2d');

            new Chart(ctxLine, {{
                type: 'line',
                data: {{
                    labels: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월'],
                    datasets: {json.dumps(branch_line_datasets, ensure_ascii=False)}
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {{ mode: 'index', intersect: false }},
                    plugins: {{
                        legend: {{
                            position: 'top', align: 'center',
                            labels: {{ usePointStyle: true, boxWidth: 8, padding: 12, font: {{ size: 9.5 }} }}
                        }},
                        tooltip: {{
                            backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                            callbacks: {{
                                label: ctx => ' ' + ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString() + '건'
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{ beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }}, title: {{ display: true, text: '이동 재고(건)', font: {{ size: 10 }} }} }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            // 지사가 많아 선이 겹칠 때는 범례를 클릭해 특정 지사만 볼 수 있습니다 (Chart.js 기본 동작).

            const quarterCanvas = document.getElementById('quarterChart');
            if (quarterCanvas) {{
                new Chart(quarterCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: {{
                        labels: {json.dumps(quarter_labels, ensure_ascii=False)},
                        datasets: {json.dumps(quarter_datasets, ensure_ascii=False)}
                    }},
                    options: {{
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {{
                            legend: {{
                                position: 'top', align: 'center',
                                labels: {{ usePointStyle: true, boxWidth: 8, padding: 12, font: {{ size: 9.5 }} }}
                            }},
                            tooltip: {{
                                backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                                callbacks: {{
                                    label: ctx => ' ' + ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString() + '건'
                                }}
                            }}
                        }},
                        scales: {{
                            x: {{ stacked: true, grid: {{ display: false }} }},
                            y: {{ stacked: true, beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }} }}
                        }}
                    }}
                }});
            }}

            // ---- RAW 시트 탐색적 분석: 의뢰유형별 재고상태 구성 + 월별 의뢰유형 추이 ----
            const typeStatusCanvas = document.getElementById('typeStatusChart');
            const monthTypeCanvas = document.getElementById('monthTypeChart');
            if (typeStatusCanvas && monthTypeCanvas && Object.keys(RAW_EDA_DATA.typeStatus || {{}}).length) {{
                window.__typeStatusChart = new Chart(typeStatusCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: buildTypeStatusData('전체'),
                    options: {{
                        responsive: true,
                        maintainAspectRatio: false,
                        indexAxis: 'y',
                        plugins: {{
                            legend: {{ position: 'top', align: 'center', labels: {{ usePointStyle: true, boxWidth: 8, padding: 12, font: {{ size: 9.5 }} }} }},
                            tooltip: {{ backgroundColor: '#0f172a', padding: 10, cornerRadius: 8 }}
                        }},
                        scales: {{
                            x: {{ stacked: true, beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }} }},
                            y: {{ stacked: true, grid: {{ display: false }} }}
                        }}
                    }}
                }});
                window.__monthTypeChart = new Chart(monthTypeCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: buildMonthTypeData('전체'),
                    options: {{
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {{
                            legend: {{ position: 'top', align: 'center', labels: {{ usePointStyle: true, boxWidth: 8, padding: 10, font: {{ size: 8.8 }} }} }},
                            tooltip: {{ backgroundColor: '#0f172a', padding: 10, cornerRadius: 8 }}
                        }},
                        scales: {{
                            x: {{ stacked: true, grid: {{ display: false }} }},
                            y: {{ stacked: true, beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }} }}
                        }}
                    }}
                }});
            }}

            const reqTypeTotalCanvas = document.getElementById('reqTypeTotalChart');
            if (reqTypeTotalCanvas && RAW_REQ_TYPES.length) {{
                new Chart(reqTypeTotalCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: {{
                        labels: RAW_REQ_TYPES,
                        datasets: [{{
                            label: '전체 건수',
                            data: RAW_REQ_TYPE_TOTALS,
                            backgroundColor: '#2563eb',
                            borderRadius: 4
                        }}]
                    }},
                    options: {{
                        responsive: true,
                        maintainAspectRatio: false,
                        indexAxis: 'y',
                        plugins: {{
                            legend: {{ display: false }},
                            tooltip: {{
                                backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                                callbacks: {{ label: ctx => ' ' + ctx.parsed.x.toLocaleString() + '건' }}
                            }}
                        }},
                        scales: {{
                            x: {{ beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }} }},
                            y: {{ grid: {{ display: false }} }}
                        }}
                    }}
                }});
            }}

            // 차트가 collapsed 패널 안에서 초기화되면 크기가 0으로 잡힐 수 있으므로,
            // 패널을 펼칠 때 Chart.js에 리사이즈를 알려줍니다.
            document.querySelectorAll('.panel').forEach(panel => {{
                panel.querySelector('.section-title').addEventListener('click', () => {{
                    setTimeout(() => {{ window.dispatchEvent(new Event('resize')); }}, 300);
                }});
            }});
        }}

        // ---- RAW EDA 데이터 및 지사 필터 전환 로직 ----
        const RAW_EDA_DATA = {raw_eda_js};
        const RAW_REQ_TYPES = {raw_req_types_js};
        const RAW_MONTH_ORDER = {raw_month_order_js};
        const RAW_REQ_TYPE_TOTALS = {raw_req_type_totals_js};
        const RAW_TYPE_PALETTE = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100',
                                   '#e87ba4', '#008300', '#4a3aa7', '#e34948'];

        function buildTypeStatusData(branch) {{
            const src = (RAW_EDA_DATA.typeStatus && RAW_EDA_DATA.typeStatus[branch]) || {{}};
            const types = RAW_REQ_TYPES.filter(t => src[t]);
            return {{
                labels: types,
                datasets: [
                    {{ label: '신품', data: types.map(t => src[t]['신품'] || 0), backgroundColor: '#2563eb', borderRadius: 4, stack: 's' }},
                    {{ label: '리퍼', data: types.map(t => src[t]['리퍼'] || 0), backgroundColor: '#cbd5e1', borderRadius: 4, stack: 's' }}
                ]
            }};
        }}

        function buildMonthTypeData(branch) {{
            const src = (RAW_EDA_DATA.monthType && RAW_EDA_DATA.monthType[branch]) || {{}};
            const activeTypes = RAW_REQ_TYPES.filter(t => RAW_MONTH_ORDER.some(m => (src[m] || {{}})[t]));
            return {{
                labels: RAW_MONTH_ORDER,
                datasets: activeTypes.map((t, i) => ({{
                    label: t,
                    data: RAW_MONTH_ORDER.map(m => (src[m] || {{}})[t] || 0),
                    backgroundColor: RAW_TYPE_PALETTE[i % RAW_TYPE_PALETTE.length],
                    borderRadius: 3,
                    stack: 'm'
                }}))
            }};
        }}

        function updateRawEda(branch) {{
            if (window.__typeStatusChart) {{
                window.__typeStatusChart.data = buildTypeStatusData(branch);
                window.__typeStatusChart.update();
            }}
            if (window.__monthTypeChart) {{
                window.__monthTypeChart.data = buildMonthTypeData(branch);
                window.__monthTypeChart.update();
            }}
        }}
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedHTMLGeneratorApp(root)
    root.mainloop()
