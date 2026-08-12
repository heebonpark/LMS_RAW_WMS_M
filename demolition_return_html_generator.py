# -*- coding: utf-8 -*-
"""
@file: demolition_return_html_generator.py
@description: 강북·강원본부 철거반납현황(해지철거 및 장비반납율) 전자동화 & EDA 시각화 대시보드 생성 시스템

워크플로우:
  1단계(RAW 반영): 물류전산 WMS > 입고 > ASN > 입고현황에서
      - 지사코드 삭제, 의뢰유형(해지철거), 입고예정일 2026.01.01~당일 조건으로 조회
      - "전체데이터다운로드"로 엑셀 저장 (파일명에 '입고현황' 포함, 예: ASN-입고현황.xlsx)
      다운로드 파일을 열린 워크북의 A열(창고코드) 기준으로 분리해 붙여넣습니다:
        - R_지역본부 시트: A열이 'WH2'가 아닌 모든 행
        - R_정비실 시트:   A열이 'WH2'인 행
  2단계(대시보드 생성): '정리' 시트의 본부별/지사별 요약표를 기반으로
      보안 HTML 대시보드(반납율 히트맵, 차트, 자동 리뷰)를 생성합니다.
"""

import os
import glob
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
TARGET_WORKBOOK_KEYWORD = "철거반납현황_강북강원본부 집계"
WAREHOUSE_REPAIR_CODE = "WH2"

# 버전 표시: "예전 파일을 실행하고 있는지"를 눈으로 바로 확인할 수 있도록
# 창 제목/생성되는 HTML 하단에 이 값이 그대로 노출됩니다.
SCRIPT_VERSION = "v4 (2026-08-11) - EDA분석리포트 + 철거율히트맵 + 하위4개지사강조"

HQ_NAMES = ['강남/서부본부', '강북/강원본부', '부산/경남본부', '전남/전북본부', '충남/충북본부', '대구/경북본부', '합계']
BRANCH_NAMES = ['중앙지사', '강북지사', '서대문지사', '고양지사', '의정부지사', '남양주지사', '강릉지사', '원주지사', '합계']


def esc(val):
    """HTML 표에 넣기 전 값 이스케이프 (None 안전 처리)"""
    if val is None:
        return ""
    return html.escape(str(val))


class DemolitionReturnApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"강북·강원본부 철거반납현황 전자동화 & EDA 시각화 시스템 [{SCRIPT_VERSION}]")
        self.root.geometry("600x500")
        self.root.configure(bg="#0f172a")

        now = datetime.datetime.now()
        last_day = calendar.monthrange(now.year, now.month)[1]
        self.expiry_str = datetime.date(now.year, now.month, last_day).strftime('%Y-%m-%d')

        self.setup_ui()

    def setup_ui(self):
        header = tk.Frame(self.root, bg="#1e3a8a", height=90)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="철거반납현황 전자동화 및 EDA 시각화 생성", font=("Malgun Gothic", 12, "bold"),
                 bg="#1e3a8a", fg="white").pack(pady=26)

        body = tk.Frame(self.root, bg="#0f172a")
        body.pack(fill="both", expand=True, padx=25, pady=20)

        tk.Label(body, text="📁 작업 디렉토리 (다운로드 경로)", font=("Malgun Gothic", 9, "bold"),
                 bg="#0f172a", fg="#94a3b8").pack(anchor="w", pady=(0, 4))

        self.path_var = tk.StringVar(value=r"C:\Users\user\Downloads")
        path_entry = tk.Entry(body, textvariable=self.path_var, font=("Malgun Gothic", 10),
                               bg="#1e293b", fg="#e2e8f0", insertbackground="white", relief="flat")
        path_entry.pack(fill="x", ipady=8, pady=(0, 15))

        # 버전 배지: 최신 파일을 실행 중인지 창을 여는 즉시 눈으로 확인할 수 있게 합니다.
        tk.Label(body, text=f"🏷️ 스크립트 버전: {SCRIPT_VERSION}", font=("Malgun Gothic", 8, "bold"),
                 bg="#0f172a", fg="#38bdf8").pack(anchor="w", pady=(0, 10))

        info_text = (
            f"• 조회 만료 기한 (당월 말일 자동적용): {self.expiry_str}\n"
            "• 1단계: WMS ASN 입고현황 다운로드 파일을 A열(창고코드) 기준\n"
            "  WH2 제외 → R_지역본부 / WH2만 → R_정비실 시트에 자동 반영\n"
            "• 2단계: '정리' 시트의 본부별·지사별 반납율/철거율을 히트맵/차트로 시각화"
        )
        tk.Label(body, text=info_text, font=("Malgun Gothic", 9), bg="#0f172a", fg="#cbd5e1",
                 justify="left").pack(anchor="w", pady=(0, 20))

        btn_sync = tk.Button(
            body, text="⚡ 1단계: ASN 입고현황 RAW 반영 (R_지역본부/R_정비실)", command=self.run_sync,
            font=("Malgun Gothic", 10, "bold"), bg="#3b82f6", fg="white", activebackground="#2563eb",
            relief="flat", cursor="hand2"
        )
        btn_sync.pack(fill="x", ipady=10, pady=(0, 10))

        btn_html = tk.Button(
            body, text="🎨 2단계: 철거반납현황 대시보드 보고서 생성", command=self.generate_html_report,
            font=("Malgun Gothic", 10, "bold"), bg="#10b981", fg="white", activebackground="#059669",
            relief="flat", cursor="hand2"
        )
        btn_html.pack(fill="x", ipady=12)

    # ---------------- 1단계: RAW 반영 ----------------

    def run_sync(self):
        download_dir = self.path_var.get().strip()
        if not os.path.exists(download_dir):
            download_dir = os.getcwd()

        # 파일명이 정확히 'ASN-입고현황'이 아닐 수 있어 '입고현황'을 포함한 최신 엑셀 파일을 찾습니다.
        candidates = [
            p for p in glob.glob(os.path.join(download_dir, "*입고현황*.xlsx"))
            if not os.path.basename(p).startswith('~$')
        ]
        if not candidates:
            messagebox.showerror(
                "오류",
                "다운로드 경로에서 '입고현황'이 포함된 엑셀 파일을 찾을 수 없습니다.\n"
                "WMS에서 전체데이터다운로드한 파일이 이 폴더에 있는지 확인해 주세요."
            )
            return False
        src_path = max(candidates, key=os.path.getmtime)

        excel = None
        try:
            df = pd.read_excel(src_path, sheet_name=0)
            if df.empty or df.shape[1] == 0:
                messagebox.showerror("오류", "다운로드한 파일에 데이터가 없습니다.")
                return False

            col_a_name = df.columns[0]
            col_a_str = df[col_a_name].astype(str).str.strip()
            df_hq = df[col_a_str != WAREHOUSE_REPAIR_CODE].copy()       # R_지역본부용 (WH2 제외)
            df_repair = df[col_a_str == WAREHOUSE_REPAIR_CODE].copy()   # R_정비실용 (WH2만)

            excel = win32com.client.Dispatch("Excel.Application")
            excel.ScreenUpdating = False
            excel.Calculation = -4135  # xlCalculationManual
            excel.DisplayAlerts = False  # "참조가 잘못되었습니다" 등 Excel 자체 알림 창이 자동화를 막지 않도록 억제

            target_wb = self._find_target_workbook(excel)
            if not target_wb:
                messagebox.showerror("오류", f"현재 엑셀에서 '{TARGET_WORKBOOK_KEYWORD}.xlsx' 파일이 열려있지 않습니다.")
                return False

            warnings = []
            self._paste_dataframe(target_wb, "R_지역본부", df_hq, warnings)
            self._paste_dataframe(target_wb, "R_정비실", df_repair, warnings)

            # 1단계 마지막 순서: 두 시트 반영이 끝난 뒤에만 '모두 새로고침'을 실행해
            # '정리' 시트의 피벗/수식이 최신 데이터를 반영하도록 합니다.
            self._refresh_workbook(target_wb, warnings)

            msg = (
                "RAW 데이터 반영이 완료되었습니다!\n\n"
                f"원본 파일: {os.path.basename(src_path)}\n"
                f"R_지역본부: {len(df_hq):,}건 (WH2 제외)\n"
                f"R_정비실: {len(df_repair):,}건 (WH2만)\n"
                f"✅ '정리' 시트 모두 새로고침 완료"
            )
            if warnings:
                msg += "\n\n⚠️ 일부 경고:\n" + "\n".join(warnings)
            messagebox.showinfo("성공", msg)
            return True
        except Exception as e:
            messagebox.showerror("실행 오류", f"처리 중 오류 발생:\n{e}")
            return False
        finally:
            if excel is not None:
                try:
                    excel.ScreenUpdating = True
                    excel.Calculation = -4105  # xlCalculationAutomatic
                    excel.DisplayAlerts = True
                except Exception:
                    pass

    def _paste_dataframe(self, wb, sheet_name, df, warnings=None):
        """DataFrame을 시트에 반영합니다.
        시트에 엑셀 표(ListObject)가 있으면 표 구조(헤더/서식)를 유지한 채 데이터 행만 교체하고,
        표가 없으면 사용 범위를 지우고 통째로 다시 씁니다.
        표 구조를 무시하고 UsedRange를 지워버리면 '정리' 시트의 피벗/구조적 참조가 깨지면서
        Excel이 "참조가 잘못되었습니다" 알림을 띄우는 문제가 있어, 표가 있을 때는 이를 보존합니다.

        중요: R_지역본부/R_정비실 표에는 원본 다운로드 데이터보다 뒤쪽에 VLOOKUP 등 수식으로
        계산되는 열(예: 물품상태, 입고예정, 변경본부, 변경지사, 대상)이 있을 수 있습니다.
        이런 계산 열은 절대 지우거나 덮어쓰지 않고, 원본 데이터 열(앞쪽)만 값을 기록합니다.
        행 수가 늘어나 새로 생기는 행만 계산 열의 수식을 마지막 기존 행에서 그대로 복사해 채웁니다."""
        try:
            ws = wb.Sheets(sheet_name)
        except Exception:
            raise RuntimeError(f"'{sheet_name}' 시트를 찾을 수 없습니다.")

        clean_df = df.astype(object).where(pd.notnull(df), None)
        values = clean_df.values.tolist()
        raw_ncols = len(df.columns)
        new_data_rows = len(values)

        if ws.ListObjects.Count > 0:
            tbl = ws.ListObjects.Item(1)
            try:
                header_row = tbl.Range.Row
                header_col = tbl.Range.Column
                table_ncols = tbl.Range.Columns.Count
                old_data_rows = tbl.ListRows.Count
                write_ncols = min(raw_ncols, table_ncols) if table_ncols else raw_ncols

                if table_ncols and table_ncols != raw_ncols and warnings is not None:
                    if table_ncols > raw_ncols:
                        warnings.append(
                            f"'{sheet_name}' 표에는 원본 데이터({raw_ncols}열)보다 열이 많습니다({table_ncols}열). "
                            f"앞쪽 {write_ncols}개(원본 데이터) 열만 갱신했고, 나머지 계산 열(수식)은 손대지 않았습니다."
                        )
                    else:
                        warnings.append(
                            f"'{sheet_name}' 표의 기존 열 수({table_ncols})보다 새 데이터 열 수({raw_ncols})가 많아 "
                            f"앞의 {write_ncols}개 열만 반영했습니다. 표 구조를 확인해 주세요."
                        )

                # 계산 열(원본 데이터 범위 밖) 수식을 미리 기억해 뒀다가, 새로 늘어난 행에만 그대로 확장합니다.
                # (기존 행의 계산 열은 절대 건드리지 않습니다.)
                calc_formulas = {}
                if table_ncols > write_ncols and old_data_rows > 0:
                    last_existing_row = header_row + old_data_rows
                    for c in range(write_ncols + 1, table_ncols + 1):
                        f = ws.Cells(last_existing_row, header_col + c - 1).FormulaR1C1
                        if f:
                            calc_formulas[c] = f

                # 행 수만 맞춥니다 (DataBodyRange를 통째로 지우지 않음 -> 계산 열 값 보존).
                if new_data_rows > old_data_rows:
                    for _ in range(new_data_rows - old_data_rows):
                        tbl.ListRows.Add()
                elif new_data_rows < old_data_rows:
                    for _ in range(old_data_rows - new_data_rows):
                        tbl.ListRows(tbl.ListRows.Count).Delete()

                if new_data_rows > 0:
                    start_row = header_row + 1
                    end_row = start_row + new_data_rows - 1
                    trimmed_values = [row[:write_ncols] for row in values]
                    data_range = ws.Range(
                        ws.Cells(start_row, header_col),
                        ws.Cells(end_row, header_col + write_ncols - 1)
                    )
                    data_range.Value = trimmed_values

                    # 새로 추가된 행에 대해서만 계산 열 수식을 이어서 채웁니다 (기존 행은 그대로 둠).
                    if calc_formulas and new_data_rows > old_data_rows:
                        fill_start = header_row + old_data_rows + 1
                        for c, formula in calc_formulas.items():
                            fill_range = ws.Range(
                                ws.Cells(fill_start, header_col + c - 1),
                                ws.Cells(end_row, header_col + c - 1)
                            )
                            fill_range.FormulaR1C1 = formula
                return
            except Exception as e:
                if warnings is not None:
                    warnings.append(f"'{sheet_name}' 표 갱신 중 문제가 있어 일반 범위 쓰기로 대체했습니다 ({e}).")
                # 표 처리 실패 시 아래의 일반 범위 쓰기로 폴백합니다.

        ws.UsedRange.ClearContents()
        mat = [list(df.columns)] + values
        if mat and mat[0]:
            ws.Range(ws.Cells(1, 1), ws.Cells(len(mat), len(mat[0]))).Value = mat

    def _refresh_workbook(self, wb, warnings=None):
        """워크북 전체를 새로고침하고, '정리' 시트의 피벗테이블은 개별적으로도 새로고침해
        일부 피벗이 실패해도 나머지는 정상 반영되도록 합니다."""
        try:
            wb.RefreshAll()
        except Exception as e:
            if warnings is not None:
                warnings.append(f"RefreshAll 중 오류: {e}")

        try:
            summary_ws = wb.Sheets("정리")
        except Exception:
            summary_ws = None

        if summary_ws is not None:
            try:
                for pt in summary_ws.PivotTables():
                    try:
                        pt.RefreshTable()
                    except Exception as e:
                        if warnings is not None:
                            warnings.append(f"'정리' 시트 피벗테이블 새로고침 중 오류: {e}")
            except Exception:
                pass  # 피벗테이블이 없는 시트일 수 있음
            try:
                summary_ws.Calculate()
            except Exception:
                pass

        try:
            wb.Application.CalculateFull()
        except Exception:
            pass

    def _find_target_workbook(self, excel):
        for wb in excel.Workbooks:
            if TARGET_WORKBOOK_KEYWORD in wb.Name:
                return wb
        return None

    # ---------------- 2단계: 대시보드 생성 ----------------

    def generate_html_report(self):
        pwd = simpledialog.askstring("관리자 인증", "관리자 비밀번호를 입력하세요:", show="*")
        if pwd != ADMIN_PASSWORD:
            messagebox.showerror("인증 실패", "관리자 비밀번호가 일치하지 않습니다.")
            return False

        download_dir = self.path_var.get().strip()
        if not os.path.exists(download_dir):
            download_dir = os.getcwd()

        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            target_wb = self._find_target_workbook(excel)
            if not target_wb:
                messagebox.showerror("오류", f"현재 엑셀에서 '{TARGET_WORKBOOK_KEYWORD}.xlsx' 파일이 열려있지 않습니다.")
                return False

            ws = target_wb.Sheets("정리")
            alphabet = string.ascii_letters + string.digits
            random_user_pwd = ''.join(secrets.choice(alphabet) for _ in range(8))

            # 정리 시트 레이아웃(스크린샷 기준):
            #   본부별 표: 헤더 6~7행, 데이터 8~13행(6개 본부), 합계 14행 -> B~O열(14개 컬럼)
            #   지사별 표: 헤더 17~18행, 데이터 19~26행(8개 지사), 합계 27행 -> B~O열(14개 컬럼)
            hq_rows = self._read_summary_rows(ws, range(8, 15), HQ_NAMES)
            branch_rows = self._read_summary_rows(ws, range(19, 28), BRANCH_NAMES)

            if not hq_rows or not branch_rows:
                messagebox.showerror(
                    "오류",
                    "'정리' 시트에서 본부별/지사별 표를 찾지 못했습니다.\n"
                    "시트 레이아웃이 예상과 다른 것 같습니다 (본부별 8~14행, 지사별 19~27행 기준)."
                )
                return False

            eda = self._collect_return_eda(target_wb, BRANCH_NAMES)

            html_content = self._build_html_head(random_user_pwd)
            html_content += self._build_auto_review(hq_rows, branch_rows)
            html_content += self._build_chart_section(branch_rows)
            html_content += self._build_summary_table(
                "panel-hq", "○ 본부별 해지철거 및 장비반납율", hq_rows
            )
            html_content += self._build_summary_table(
                "panel-branch", "○ 지사별 해지철거 및 장비반납율", branch_rows, highlight_bottom_n=4
            )
            html_content += self._build_return_eda_panel(eda)
            html_content += self._build_footer(random_user_pwd)
            html_content += self._build_scripts(branch_rows, eda)

            html_path = os.path.join(
                download_dir, f"demolition_return_dashboard_{datetime.date.today().strftime('%Y%m%d')}.html"
            )
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_content)

            self._show_result_dialog(html_path, random_user_pwd)
            return True

        except Exception as e:
            messagebox.showerror("실행 오류", f"보고서 생성 중 오류 발생:\n{e}")
            return False
        finally:
            excel = None  # 읽기 전용 세션이라 Quit는 호출하지 않음 (사용자의 열린 엑셀을 닫지 않기 위함)

    def _read_summary_rows(self, ws, row_range, valid_names):
        """정리 시트에서 구분(B열)이 valid_names에 속하는 행만 읽어 (이름, 14개값) 리스트로 반환."""
        rows = []
        for r in row_range:
            row_vals = [ws.Cells(r, c).Value for c in range(2, 16)]  # B~O = 14개 컬럼
            name = str(row_vals[0]).strip() if row_vals[0] is not None else ""
            if name in valid_names:
                rows.append((name, row_vals))
        return rows

    def _is_excel_error(self, val):
        """COM으로 셀을 읽을 때 #N/A 등 오류값은 문자열이 아니라 거대한 음수 정수로 반환됩니다."""
        if isinstance(val, bool):
            return False
        if isinstance(val, (int, float)) and -2146826300 <= val <= -2146826000:
            return True
        if isinstance(val, str) and val.strip().startswith('#'):
            return True
        return False

    def _excel_serial_to_date(self, serial):
        """엑셀 날짜 일련번호를 datetime으로 변환 (셀 서식이 날짜가 아닌 숫자로 읽힐 때 대비)."""
        try:
            return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(serial))
        except (ValueError, OverflowError, TypeError):
            return None

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

    def _collect_return_eda(self, target_wb, branch_names):
        """'R_지역본부' 시트에서 물품상태·입고예정·변경지사·대상 컬럼을 찾아
        EDA(탐색적 분석)용 집계 데이터를 만듭니다. 시트/컬럼이 없으면 available=False로 반환합니다."""
        result = {"available": False}
        try:
            ws = target_wb.Sheets("R_지역본부")
        except Exception:
            return result

        try:
            header_map = {}
            for c in range(1, 41):
                v = ws.Cells(1, c).Value
                if v:
                    header_map[str(v).strip()] = c

            def find_col(*keywords):
                for kw in keywords:
                    for name, idx in header_map.items():
                        if name == kw:
                            return idx
                    for name, idx in header_map.items():
                        if kw in name:
                            return idx
                return None

            col_status = find_col('물품상태')
            col_month = find_col('입고예정')
            col_branch = find_col('변경지사')
            col_target = find_col('대상')

            if not all([col_status, col_month, col_branch]):
                return result

            max_col_used = max(header_map.values())
            last_row = ws.Cells(ws.Rows.Count, col_status).End(-4162).Row  # xlUp
            if last_row < 2:
                return result

            data = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, max_col_used)).Value
            if data and not isinstance(data[0], tuple):
                data = (data,)

            status_counts = {}
            month_counts = {}
            branch_counts = {}
            branch_status = {}
            branch_month = {}
            target_counts = {"대상": 0, "제외": 0, "기타": 0}
            month_sort_keys = {}  # 표시용 라벨(MM-DD) -> 정렬용 (년,월,일) 튜플
            skipped_na = 0

            for row in data:
                status = row[col_status - 1]
                month = row[col_month - 1]
                branch = row[col_branch - 1]
                target = row[col_target - 1] if col_target else None

                if self._is_excel_error(status) or self._is_excel_error(branch):
                    skipped_na += 1
                    continue
                if not status:
                    continue

                status_str = str(status).strip()

                # 날짜 셀은 datetime/일련번호/문자열 등 형태가 제각각이라 파싱 후 "MM-DD"로 통일합니다.
                # (파싱 실패 시에만 원본 값을 그대로 문자열로 표시)
                month_date = self._parse_date_value(month)
                if month_date is not None:
                    month_label = month_date.strftime('%m-%d')
                    month_sort_keys[month_label] = (month_date.year, month_date.month, month_date.day)
                elif month not in (None, ""):
                    month_label = str(month).strip()
                    month_sort_keys.setdefault(month_label, (9999, 99, 99))
                else:
                    month_label = "미상"
                    month_sort_keys.setdefault(month_label, (9999, 99, 99))

                branch_str = str(branch).strip() if branch not in (None, "") else None
                if branch_str is not None and (self._is_excel_error(branch_str) or branch_str.startswith('#')):
                    branch_str = None

                status_counts[status_str] = status_counts.get(status_str, 0) + 1
                month_counts[month_label] = month_counts.get(month_label, 0) + 1

                if branch_str:
                    branch_counts[branch_str] = branch_counts.get(branch_str, 0) + 1
                    branch_status.setdefault(branch_str, {})
                    branch_status[branch_str][status_str] = branch_status[branch_str].get(status_str, 0) + 1
                    branch_month.setdefault(branch_str, {})
                    branch_month[branch_str][month_label] = branch_month[branch_str].get(month_label, 0) + 1

                if target is not None:
                    t = str(target).strip()
                    if '제외' in t:
                        target_counts["제외"] += 1
                    elif '대상' in t:
                        target_counts["대상"] += 1
                    else:
                        target_counts["기타"] += 1

            if not status_counts:
                return result

            month_order = sorted(month_counts.keys(), key=lambda m: month_sort_keys.get(m, (9999, 99, 99)))
            status_list = sorted(status_counts.keys(), key=lambda s: status_counts[s], reverse=True)

            # 지사(중앙~원주) 순서로만 정렬합니다. 그 외 지사명(다른 본부/오타 등)은 지사 x 물품상태
            # 교차표에는 노출하지 않습니다 (branch_counts/branch_status에는 원본 그대로 유지).
            branch_order = [b for b in branch_names if b != '합계' and b in branch_counts]

            result.update({
                "available": True,
                "status_counts": status_counts,
                "status_list": status_list,
                "month_counts": month_counts,
                "month_order": month_order,
                "branch_counts": branch_counts,
                "branch_order": branch_order,
                "branch_status": branch_status,
                "branch_month": branch_month,
                "target_counts": target_counts,
                "skipped_na": skipped_na,
            })
            return result
        except Exception:
            return {"available": False}

    def _show_result_dialog(self, html_path, random_user_pwd):
        """messagebox는 텍스트를 마우스로 선택/복사할 수 없으므로,
        비밀번호를 복사-붙여넣기 할 수 있는 전용 결과창을 띄웁니다."""
        dlg = tk.Toplevel(self.root)
        dlg.title("해지고객 장비 반납 현황 리포트 생성 완료")
        dlg.configure(bg="#0f172a")
        dlg.resizable(False, False)
        dlg.grab_set()

        frame = tk.Frame(dlg, bg="#0f172a")
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="🎨 해지고객 장비 반납 현황 리포트 생성 완료!", font=("Malgun Gothic", 12, "bold"),
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

    # ---------------- HTML 빌더 ----------------

    def _to_ratio(self, val):
        """철거율/반납율 값을 0~1 사이의 비율로 정규화합니다.
        엑셀 셀이 진짜 퍼센트 서식(0.817)이 아니라 이미 100을 곱한 숫자(81.7)이거나
        문자열("81.7%")로 되어 있어도 안전하게 처리합니다. 인식 불가하면 None."""
        if val is None:
            return None
        if isinstance(val, bool):
            return None
        if isinstance(val, str):
            s = val.strip()
            has_pct = s.endswith('%')
            s = s.rstrip('%').strip()
            try:
                num = float(s)
            except ValueError:
                return None
            return num / 100 if (has_pct or num > 1) else num
        if isinstance(val, (int, float)):
            return val / 100 if val > 1 else val
        return None

    def _diverging_style(self, value, vmin, vmax):
        """반납율처럼 '높을수록 좋은' 지표에 빨강(낮음)-흰색(중간)-파랑(높음) 색상 스케일을 적용합니다.
        (엑셀 조건부서식 3색 스케일과 동일한 느낌)"""
        if vmax == vmin or value is None:
            return ""
        ratio = max(0.0, min(1.0, (value - vmin) / (vmax - vmin)))
        if ratio < 0.5:
            t = ratio / 0.5
            r, g, b = 239, 68, 68
            r2, g2, b2 = 255, 255, 255
        else:
            t = (ratio - 0.5) / 0.5
            r, g, b = 255, 255, 255
            r2, g2, b2 = 37, 99, 235
        rr = round(r + (r2 - r) * t)
        gg = round(g + (g2 - g) * t)
        bb = round(b + (b2 - b) * t)
        text_color = "#ffffff" if (ratio < 0.18 or ratio > 0.82) else "#1e2233"
        return f' style="background:rgb({rr},{gg},{bb}); color:{text_color}; font-weight:700;"'

    def _blue_heat_style(self, value, vmax):
        """단순 건수 히트맵 (값이 클수록 진한 파란색). 반납율처럼 '좋다/나쁘다' 가치판단이 없는
        단순 빈도 데이터(물품상태·지사 교차표 등)에 사용합니다."""
        if not vmax or value in (None, 0):
            return ""
        ratio = min(value / vmax, 1.0)
        alpha = 0.08 + ratio * 0.55
        text_color = "#0b1b4d" if ratio < 0.6 else "#ffffff"
        weight = "700" if ratio >= 0.35 else "500"
        return f' style="background:rgba(37,99,235,{alpha:.2f}); color:{text_color}; font-weight:{weight};"'

    def _panel_wrap(self, panel_id, title_html, unit_label, body_html):
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

    def _build_html_head(self, random_user_pwd):
        js_user_pwd = json.dumps(random_user_pwd)
        js_admin_pwd = json.dumps(ADMIN_PASSWORD)
        js_expiry = json.dumps(self.expiry_str)

        return f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>'26년도 강북 / 강원본부 해지고객 장비 반납 현황 리포트</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="preconnect" href="https://cdn.jsdelivr.net">
    <link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css" rel="stylesheet">
    <style>
        :root {{
            --navy: #0b1b4d; --blue: #1d4ed8; --blue-light: #3b82f6; --amber: #f59e0b;
            --slate: #64748b; --bg: #eef1f8; --card-border: #e6e9f2;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Pretendard', 'Malgun Gothic', 'Segoe UI', sans-serif;
            background: radial-gradient(1200px 600px at 10% -10%, #dbe4ff 0%, var(--bg) 45%), var(--bg);
            color: #1e2233; margin: 0; padding: 32px 20px; letter-spacing: -0.01em;
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
        #report-content {{
            display: none; max-width: 1480px; margin: 0 auto; background: #ffffff;
            padding: 40px 44px; border-radius: 20px; box-shadow: 0 20px 40px -18px rgba(15,23,42,0.15);
            border: 1px solid var(--card-border);
        }}
        .security-banner {{
            background: linear-gradient(90deg, #0b1230, #16255e); color: #7dd3fc; padding: 13px 22px;
            border-radius: 12px; margin-bottom: 26px; display: flex; justify-content: space-between;
            font-size: 9.5pt; font-weight: 600; align-items: center;
        }}
        .header-card {{
            background: linear-gradient(135deg, var(--navy) 0%, var(--blue) 60%, #2563eb 100%);
            color: white; padding: 34px; border-radius: 16px; margin-bottom: 26px; text-align: center;
            box-shadow: 0 20px 35px -12px rgba(29,78,216,0.4);
        }}
        .header-card h1 {{ margin: 0 0 8px 0; font-size: 22pt; font-weight: 800; letter-spacing: -0.02em; }}
        .header-card p {{ margin: 0; font-size: 10.5pt; color: #bfdbfe; }}
        .chart-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .chart-card {{
            background: #fbfcfe; padding: 22px 20px 16px; border-radius: 14px; border: 1px solid var(--card-border);
            box-shadow: 0 1px 3px rgba(15,23,42,0.03);
        }}
        .chart-card h3 {{ margin-top: 0; font-size: 10.5pt; font-weight: 700; color: var(--navy); text-align: center; margin-bottom: 14px; }}
        .chart-canvas-wrap {{ position: relative; width: 100%; height: 300px; }}

        .panel {{ margin-bottom: 10px; }}
        .section-title {{
            font-size: 13.5pt; font-weight: 800; color: var(--navy); margin: 32px 0 14px 0;
            border-left: 4px solid var(--blue-light); padding: 2px 0 2px 12px; display: flex;
            justify-content: space-between; align-items: center; cursor: pointer; user-select: none;
        }}
        .section-title:hover {{ opacity: 0.75; }}
        .section-title-right {{ display: flex; align-items: center; gap: 10px; }}
        .unit-label {{ font-size: 9pt; font-weight: 400; color: var(--slate); }}
        .chevron {{ display: inline-block; font-size: 10pt; color: var(--slate); transition: transform .25s ease; }}
        .panel.collapsed .chevron {{ transform: rotate(-90deg); }}
        .panel-body {{ display: grid; grid-template-rows: 1fr; transition: grid-template-rows .28s ease; }}
        .panel.collapsed .panel-body {{ grid-template-rows: 0fr; }}
        .panel-body-inner {{ overflow: hidden; min-height: 0; }}

        .review-card {{
            background: linear-gradient(135deg, #f5f8ff 0%, #eef2fb 100%); border: 1px solid #dde5f9;
            border-radius: 14px; padding: 20px 22px 16px;
        }}
        .review-list {{ list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 9px; }}
        .review-list li {{ font-size: 9.6pt; line-height: 1.5; color: #33415c; display: flex; gap: 8px; align-items: flex-start; }}
        .review-list li .badge {{ flex-shrink: 0; font-size: 10pt; }}
        .review-list li b {{ color: var(--navy); }}
        .review-list li.tone-warn b {{ color: #b42318; }}
        .review-list li.tone-good b {{ color: #0f9d58; }}

        .eda-tabs {{ display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 16px; }}
        .eda-tab-btn {{
            border: 1px solid var(--card-border); background: white; color: var(--slate);
            font-size: 8.8pt; font-weight: 600; padding: 6px 12px; border-radius: 999px;
            cursor: pointer; transition: all .15s ease;
        }}
        .eda-tab-btn:hover {{ border-color: var(--blue-light); color: var(--blue); }}
        .eda-tab-btn.active {{ background: var(--blue); border-color: var(--blue); color: white; }}

        table {{ width: 100%; border-collapse: separate; border-spacing: 0; background: white; font-size: 9.2pt; }}
        .table-wrap {{ border-radius: 12px; overflow-x: auto; box-shadow: 0 1px 2px rgba(15,23,42,0.04); border: 1px solid var(--card-border); }}
        th, td {{ border-bottom: 1px solid #eef1f6; padding: 9px 10px; text-align: center; vertical-align: middle; white-space: nowrap; font-variant-numeric: tabular-nums; }}
        th {{ background-color: #f6f8fc; color: #33415c; font-weight: 700; text-transform: uppercase; letter-spacing: 0.02em; font-size: 8.5pt; }}
        tbody tr:last-child td {{ border-bottom: none; }}
        tr:hover td {{ background-color: rgba(245,248,255,0.6); }}
        .total-row td {{ background-color: #e8edfb !important; font-weight: 800; color: var(--navy); }}
        .bottom-row td {{ background-color: #fef2f2; }}
        .bottom-row td:first-child {{ font-weight: 700; color: #b42318; }}
        .bottom-badge {{
            display: inline-block; font-size: 7.3pt; font-weight: 700; color: #b42318; background: #fee2e2;
            padding: 1px 6px; border-radius: 999px; margin-left: 4px; vertical-align: middle;
        }}
        .rank-caption {{ font-size: 8.3pt; color: var(--slate); }}
        .footer {{ text-align: right; font-size: 8.8pt; color: var(--slate); border-top: 1px solid var(--card-border); padding-top: 18px; margin-top: 36px; }}

        @media (max-width: 880px) {{
            body {{ padding: 16px 10px; }}
            #report-content {{ padding: 22px 16px; border-radius: 14px; }}
            .header-card {{ padding: 22px 16px; }}
            .header-card h1 {{ font-size: 16.5pt; }}
            .chart-grid {{ grid-template-columns: 1fr; }}
            .chart-canvas-wrap {{ height: 240px; }}
            .section-title {{ font-size: 12pt; }}
            table {{ font-size: 8.4pt; }}
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
            <h1>'26년도 강북 / 강원본부 해지고객 장비 반납 현황 리포트</h1>
            <p>해지철거 및 장비반납율 EDA 분석 리포트</p>
        </div>
"""

    def _build_chart_section(self, branch_rows):
        body = """
        <div class="chart-grid">
            <div class="chart-card">
                <h3>📊 지사별 반납율 (8월 실적 vs 26년 누계)</h3>
                <div class="chart-canvas-wrap"><canvas id="returnRateChart"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>📦 지사별 미철거 현황 (누계)</h3>
                <div class="chart-canvas-wrap"><canvas id="pendingChart"></canvas></div>
            </div>
        </div>
"""
        return self._panel_wrap("panel-chart", "📊 EDA 시각화 차트", "", body)

    def _build_auto_review(self, hq_rows, branch_rows):
        """본부별/지사별 데이터를 기반으로 자동 인사이트 문구를 생성합니다."""
        items = []  # (tone, badge, text)

        branch_only = [(n, v) for n, v in branch_rows if n != '합계']
        if branch_only:
            # 누계 반납율(idx 12) 기준. _to_ratio로 정규화해 셀이 0~1 소수든 0~100 숫자든 동일하게 처리합니다
            # (정규화하지 않으면 _build_summary_table과 값 표기가 어긋날 수 있습니다).
            rates = [(n, self._to_ratio(v[12])) for n, v in branch_only if self._to_ratio(v[12]) is not None]
            if rates:
                avg_rate = sum(r for _, r in rates) / len(rates) * 100
                best_name, best_rate = max(rates, key=lambda x: x[1])
                worst_name, worst_rate = min(rates, key=lambda x: x[1])
                items.append(("good" if avg_rate >= 80 else "warn", "📊",
                               f"지사 평균 누계 반납율은 <b>{avg_rate:.1f}%</b> 입니다."))
                items.append(("good", "🏆", f"최고 반납율 지사: <b>{esc(best_name)}</b> ({best_rate*100:.1f}%)"))
                items.append(("warn" if worst_rate < 0.75 else "neutral", "🔻",
                               f"최저 반납율 지사: <b>{esc(worst_name)}</b> ({worst_rate*100:.1f}%) — 개선 필요"))

            # 미철거(idx 10, 누계) 기준
            pending = [(n, v[10]) for n, v in branch_only if isinstance(v[10], (int, float))]
            if pending:
                pending_sorted = sorted(pending, key=lambda x: x[1], reverse=True)
                top_name, top_val = pending_sorted[0]
                if top_val > 0:
                    items.append(("warn", "⚠️", f"미철거 최다 지사: <b>{esc(top_name)}</b> ({top_val:,.0f}건) — 우선 정리 대상"))

        hq_only = [(n, v) for n, v in hq_rows if n != '합계']
        rank_row = next((v for n, v in hq_rows if n == '강북/강원본부'), None)
        if rank_row and isinstance(rank_row[13], (int, float)):
            items.append(("neutral", "📍", f"강북/강원본부는 전체 6개 본부 중 누계 반납율 <b>{int(rank_row[13])}위</b>입니다."))

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

    def _build_summary_table(self, panel_id, title, rows, highlight_bottom_n=0):
        header = """
                <tr>
                    <th rowspan="2">구분</th>
                    <th colspan="6">8월 실적</th>
                    <th colspan="6">26년 누계 실적</th>
                    <th rowspan="2">순위</th>
                </tr>
                <tr>
                    <th>대상시설</th><th>반납완료</th><th>반납예정</th><th>미철거</th><th>철거율</th><th>반납율</th>
                    <th>대상시설</th><th>반납완료</th><th>반납예정</th><th>미철거</th><th>철거율</th><th>반납율</th>
                </tr>
"""
        # 누계 반납율(idx 12) 기준 색상 스케일 계산 (합계 행 제외). _to_ratio로 정규화해
        # 값이 0.817이든 81.7이든 "81.7%"든 동일하게 처리합니다.
        rate_vals = [
            self._to_ratio(v[12]) for n, v in rows if n != '합계' and self._to_ratio(v[12]) is not None
        ]
        vmin, vmax = (min(rate_vals), max(rate_vals)) if rate_vals else (0, 1)

        # 누계 철거율(idx 11)도 값이 낮을수록(부진할수록) 빨간색이 되도록 동일한 방식의 색상 스케일 적용
        demo_vals = [
            self._to_ratio(v[11]) for n, v in rows if n != '합계' and self._to_ratio(v[11]) is not None
        ]
        demo_vmin, demo_vmax = (min(demo_vals), max(demo_vals)) if demo_vals else (0, 1)

        # 순위(idx 13) 기준 하위 N개를 가독성 있게 행 전체를 강조합니다 (히트맵은 반납율 칸에만 적용되어
        # 나머지 칸은 그냥 흰 배경이라 어느 행이 부진한지 한눈에 안 들어오는 문제를 보완).
        bottom_names = set()
        if highlight_bottom_n > 0:
            rank_vals = [(n, v[13]) for n, v in rows if n != '합계' and isinstance(v[13], (int, float))]
            rank_vals.sort(key=lambda x: x[1], reverse=True)  # 순위 숫자가 클수록 하위
            bottom_names = {n for n, _ in rank_vals[:highlight_bottom_n]}

        rows_html = ""
        for name, vals in rows:
            is_total = (name == '합계')
            if is_total:
                row_cls = ' class="total-row"'
            elif name in bottom_names:
                row_cls = ' class="bottom-row"'
            else:
                row_cls = ""
            cells = f"<td>{esc(name)}</td>"
            for idx, val in enumerate(vals[1:], start=1):
                if val is None:
                    cells += "<td></td>"
                    continue
                if idx == 5:  # 철거율(8월) - 색상 없이 표시
                    ratio = self._to_ratio(val)
                    cells += f"<td>{ratio*100:.1f}%</td>" if ratio is not None else f"<td>{esc(val)}</td>"
                elif idx == 6:  # 반납율(8월) - 색상 없이 표시
                    ratio = self._to_ratio(val)
                    cells += f"<td>{ratio*100:.1f}%</td>" if ratio is not None else f"<td>{esc(val)}</td>"
                elif idx == 11:  # 철거율(누계) - 부진순위별 히트맵 색상 (낮을수록 빨강)
                    ratio = self._to_ratio(val)
                    if ratio is not None and not is_total:
                        style = self._diverging_style(ratio, demo_vmin, demo_vmax)
                        cells += f"<td{style}>{ratio*100:.1f}%</td>"
                    elif ratio is not None:
                        cells += f"<td>{ratio*100:.1f}%</td>"
                    else:
                        cells += f"<td>{esc(val)}</td>"
                elif idx == 12:  # 반납율(누계) - 히트맵 색상
                    ratio = self._to_ratio(val)
                    if ratio is not None and not is_total:
                        style = self._diverging_style(ratio, vmin, vmax)
                        cells += f"<td{style}>{ratio*100:.1f}%</td>"
                    elif ratio is not None:
                        cells += f"<td>{ratio*100:.1f}%</td>"
                    else:
                        cells += f"<td>{esc(val)}</td>"
                elif idx == 13:  # 순위
                    rank_badge = ""
                    if isinstance(val, (int, float)) and name in bottom_names:
                        rank_badge = ' <span class="bottom-badge">하위</span>'
                    cells += (f"<td>{val:,.0f}{rank_badge}</td>" if isinstance(val, (int, float))
                              else f"<td>{esc(val)}</td>")
                elif isinstance(val, (int, float)):
                    cells += f"<td>{val:,.0f}</td>"
                else:
                    cells += f"<td>{esc(val)}</td>"
            rows_html += f"<tr{row_cls}>{cells}</tr>"

        caption = (
            "※ 누계 반납율·누계 철거율은 값이 높을수록 파란색, 낮을수록(부진할수록) 빨간색으로 "
            "표시됩니다 (엑셀 조건부서식과 동일한 방식)."
        )
        if highlight_bottom_n > 0:
            caption += f' 순위 하위 {highlight_bottom_n}개 지사는 행 전체를 옅은 붉은색으로 강조했습니다.'

        body = f"""
        <div class="table-wrap">
        <table>
            <thead>{header}</thead>
            <tbody>{rows_html}</tbody>
        </table>
        </div>
        <div class="rank-caption" style="font-size:8.3pt; color:var(--slate); margin-top:8px;">
            {caption}
        </div>
"""
        return self._panel_wrap(panel_id, title, "[단위: 건, %]", body)

    def _build_return_eda_panel(self, eda):
        """R_지역본부 원본 데이터 기반 EDA(탐색적 분석) 패널: 물품상태 분포, 입고예정 월별 추이,
        지사별 구성 히트맵을 보여줍니다."""
        if not eda.get("available"):
            body = (
                '<div class="review-card"><ul class="review-list"><li><span class="badge">ℹ️</span>'
                "<span>'R_지역본부' 시트에서 물품상태/입고예정/변경지사 컬럼을 찾지 못해 "
                "이 분석을 건너뛰었습니다. 컬럼명을 확인해 주세요.</span></li></ul></div>"
            )
            return self._panel_wrap("panel-eda", "🔬 EDA 분석 리포트 (R_지역본부)", "", body)

        na_notice = ""
        if eda.get("skipped_na"):
            na_notice = (
                f'<div class="rank-caption" style="margin-bottom:14px;">'
                f'※ 물품상태/변경지사가 #N/A 등 오류값인 {eda["skipped_na"]:,}건은 집계에서 제외했습니다.</div>'
            )

        # 대상/제외 요약 배지
        tgt = eda["target_counts"]
        tgt_total = sum(tgt.values()) or 1
        target_html = ""
        if tgt_total > 1 or tgt["대상"] or tgt["제외"]:
            pct_target = tgt["대상"] / tgt_total * 100
            pct_exclude = tgt["제외"] / tgt_total * 100
            target_html = f"""
        <div class="review-card" style="margin-bottom:18px;">
            <ul class="review-list">
                <li><span class="badge">📌</span><span>대상 <b>{tgt['대상']:,}건</b> ({pct_target:.1f}%) · 제외 <b>{tgt['제외']:,}건</b> ({pct_exclude:.1f}%)</span></li>
            </ul>
        </div>
"""

        # 지사(중앙~원주) × 물품상태 교차표 (히트맵)
        status_list = eda["status_list"]
        branch_order = eda["branch_order"]
        branch_status = eda["branch_status"]
        all_cell_vals = [branch_status.get(b, {}).get(s, 0) for b in branch_order for s in status_list]
        heat_max = max(all_cell_vals) if all_cell_vals else 1

        header_cells = "".join(f"<th>{esc(s)}</th>" for s in status_list) + "<th>합계</th>"
        matrix_rows = ""
        for b in branch_order:
            row_total = 0
            cells = ""
            for s in status_list:
                v = branch_status.get(b, {}).get(s, 0)
                row_total += v
                style = self._blue_heat_style(v, heat_max)
                cells += f"<td{style}>{v:,.0f}</td>" if v else "<td></td>"
            matrix_rows += f"<tr><td>{esc(b)}</td>{cells}<td><b>{row_total:,.0f}</b></td></tr>"

        # 지사 선택 버튼: 강북강원본부(전체) + 중앙~원주 8개 지사
        tab_targets = ["강북강원본부"] + [b for b in BRANCH_NAMES if b != '합계']
        tab_buttons = "".join(
            f'<button type="button" class="eda-tab-btn{" active" if i == 0 else ""}" '
            f'data-branch="{esc(b)}" onclick="selectEdaBranch(this)">{esc(b)}</button>'
            for i, b in enumerate(tab_targets)
        )

        body = f"""
        {na_notice}
        {target_html}
        <div class="eda-tabs">{tab_buttons}</div>
        <div class="chart-grid">
            <div class="chart-card">
                <h3>📦 물품상태별 분포</h3>
                <div class="chart-canvas-wrap"><canvas id="returnStatusChart"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>🗓️ 입고예정 월별 추이</h3>
                <div class="chart-canvas-wrap"><canvas id="returnMonthChart"></canvas></div>
            </div>
        </div>
        <div class="section-title" style="cursor:default; margin-top:28px;">
            <span>📋 지사(중앙~원주) × 물품상태 교차표</span>
            <span class="section-title-right"><span class="unit-label">[단위: 건]</span></span>
        </div>
        <div class="table-wrap">
        <table>
            <thead><tr><th>변경지사 \\ 물품상태</th>{header_cells}</tr></thead>
            <tbody>{matrix_rows}</tbody>
        </table>
        </div>
        <div class="rank-caption" style="margin-top:8px;">※ 색이 진할수록 건수가 많은 셀입니다.</div>
"""
        return self._panel_wrap("panel-eda", "🔬 EDA 분석 리포트 (R_지역본부)", "", body)

    def _build_footer(self, random_user_pwd):
        return f"""
        <div class="footer">
            Generated by Demolition-Return Secure Visualization System {esc(SCRIPT_VERSION)} | 조회 기한: {esc(self.expiry_str)}
        </div>
    </div>
"""

    def _build_scripts(self, branch_rows, eda=None):
        eda = eda or {}
        branch_only = [(n, v) for n, v in branch_rows if n != '합계']
        labels = [n for n, v in branch_only]
        # _to_ratio로 정규화해 셀이 0~1 소수든 0~100 숫자든 동일하게 처리합니다.
        rate_aug = [round((self._to_ratio(v[6]) or 0) * 100, 1) for n, v in branch_only]   # 8월 반납율
        rate_cum = [round((self._to_ratio(v[12]) or 0) * 100, 1) for n, v in branch_only]  # 누계 반납율
        pending_cum = [v[10] if isinstance(v[10], (int, float)) else 0 for n, v in branch_only]  # 누계 미철거

        status_labels = eda.get("status_list", [])
        status_values = [eda.get("status_counts", {}).get(s, 0) for s in status_labels]
        month_labels = eda.get("month_order", [])
        month_values = [eda.get("month_counts", {}).get(m, 0) for m in month_labels]

        # 지사 선택 버튼(강북강원본부 전체 + 중앙~원주)을 클릭했을 때 두 EDA 차트를 갱신하기 위한
        # 지사별 데이터. 축 라벨(status_labels/month_labels)은 전체 기준으로 고정해 지사를 바꿔도
        # 차트 축이 흔들리지 않게 합니다.
        branch_status_map = eda.get("branch_status", {})
        branch_month_map = eda.get("branch_month", {})
        status_by_branch = {"강북강원본부": status_values}
        month_by_branch = {"강북강원본부": month_values}
        for b in BRANCH_NAMES:
            if b == '합계':
                continue
            bs = branch_status_map.get(b, {})
            bm = branch_month_map.get(b, {})
            status_by_branch[b] = [bs.get(s, 0) for s in status_labels]
            month_by_branch[b] = [bm.get(m, 0) for m in month_labels]

        return f"""
    <script>
        Chart.defaults.font.family = "'Pretendard', 'Malgun Gothic', 'Segoe UI', sans-serif";
        Chart.defaults.color = '#475569';

        const EDA_STATUS_BY_BRANCH = {json.dumps(status_by_branch, ensure_ascii=False)};
        const EDA_MONTH_BY_BRANCH = {json.dumps(month_by_branch, ensure_ascii=False)};
        let returnStatusChartInstance = null;
        let returnMonthChartInstance = null;

        function selectEdaBranch(btn) {{
            const branch = btn.getAttribute('data-branch');
            document.querySelectorAll('.eda-tabs .eda-tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            if (returnStatusChartInstance && EDA_STATUS_BY_BRANCH[branch]) {{
                returnStatusChartInstance.data.datasets[0].data = EDA_STATUS_BY_BRANCH[branch];
                returnStatusChartInstance.update();
            }}
            if (returnMonthChartInstance && EDA_MONTH_BY_BRANCH[branch]) {{
                returnMonthChartInstance.data.datasets[0].data = EDA_MONTH_BY_BRANCH[branch];
                returnMonthChartInstance.update();
            }}
        }}

        function initCharts() {{
            const ctxRate = document.getElementById('returnRateChart').getContext('2d');
            new Chart(ctxRate, {{
                type: 'bar',
                data: {{
                    labels: {json.dumps(labels, ensure_ascii=False)},
                    datasets: [
                        {{ label: '8월 반납율(%)', data: {json.dumps(rate_aug)}, backgroundColor: 'rgba(148,163,184,0.6)', borderRadius: 5 }},
                        {{ label: '26년 누계 반납율(%)', data: {json.dumps(rate_cum)}, backgroundColor: '#2563eb', borderRadius: 5 }}
                    ]
                }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ position: 'top', align: 'center', labels: {{ usePointStyle: true, boxWidth: 8, padding: 14, font: {{ size: 9.5 }} }} }},
                        tooltip: {{
                            backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                            callbacks: {{ label: ctx => ' ' + ctx.dataset.label + ': ' + ctx.parsed.y + '%' }}
                        }}
                    }},
                    scales: {{
                        y: {{ beginAtZero: true, max: 100, grid: {{ color: 'rgba(148,163,184,0.15)' }}, ticks: {{ callback: v => v + '%' }} }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            const ctxPending = document.getElementById('pendingChart').getContext('2d');
            new Chart(ctxPending, {{
                type: 'bar',
                data: {{
                    labels: {json.dumps(labels, ensure_ascii=False)},
                    datasets: [{{
                        label: '미철거(건, 누계)', data: {json.dumps(pending_cum)},
                        backgroundColor: '#f59e0b', borderRadius: 5
                    }}]
                }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ display: false }},
                        tooltip: {{
                            backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                            callbacks: {{ label: ctx => ' ' + ctx.parsed.y.toLocaleString() + '건' }}
                        }}
                    }},
                    scales: {{
                        y: {{ beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }} }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            const statusCanvas = document.getElementById('returnStatusChart');
            if (statusCanvas) {{
                returnStatusChartInstance = new Chart(statusCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: {{
                        labels: {json.dumps(status_labels, ensure_ascii=False)},
                        datasets: [{{
                            label: '건수', data: {json.dumps(status_values)},
                            backgroundColor: '#2563eb', borderRadius: 4
                        }}]
                    }},
                    options: {{
                        responsive: true, maintainAspectRatio: false, indexAxis: 'y',
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

            const monthCanvas = document.getElementById('returnMonthChart');
            if (monthCanvas) {{
                returnMonthChartInstance = new Chart(monthCanvas.getContext('2d'), {{
                    type: 'bar',
                    data: {{
                        labels: {json.dumps(month_labels, ensure_ascii=False)},
                        datasets: [{{
                            label: '건수', data: {json.dumps(month_values)},
                            backgroundColor: '#10b981', borderRadius: 4
                        }}]
                    }},
                    options: {{
                        responsive: true, maintainAspectRatio: false,
                        plugins: {{
                            legend: {{ display: false }},
                            tooltip: {{
                                backgroundColor: '#0f172a', padding: 10, cornerRadius: 8,
                                callbacks: {{ label: ctx => ' ' + ctx.parsed.y.toLocaleString() + '건' }}
                            }}
                        }},
                        scales: {{
                            y: {{ beginAtZero: true, grid: {{ color: 'rgba(148,163,184,0.15)' }} }},
                            x: {{ grid: {{ display: false }} }}
                        }}
                    }}
                }});
            }}

            document.querySelectorAll('.panel').forEach(panel => {{
                panel.querySelector('.section-title').addEventListener('click', () => {{
                    setTimeout(() => {{ window.dispatchEvent(new Event('resize')); }}, 300);
                }});
            }});
        }}
    </script>
</body>
</html>
"""


if __name__ == "__main__":
    root = tk.Tk()
    app = DemolitionReturnApp(root)
    root.mainloop()
