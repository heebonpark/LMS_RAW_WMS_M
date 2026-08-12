# -*- coding: utf-8 -*-
"""
@file: main_dashboard.py
@description: 이동재고 및 철거반납 통합 관리 GUI 대시보드
"""
import os
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
import datetime
import pythoncom

# 기존 모듈 임포트
import inventory_html_generator as inv
import demolition_return_html_generator as dem

class UnifiedDashboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("이동재고 및 철거반납 통합 관리 시스템")
        self.root.geometry("650x600")
        self.root.configure(bg="#0f172a")
        
        # 내부 모듈 초기화를 위한 숨겨진 Toplevel (UI 간섭 방지)
        self.hidden_inv = tk.Toplevel(self.root)
        self.hidden_inv.withdraw()
        self.inv_app = inv.AdvancedHTMLGeneratorApp(self.hidden_inv)
        
        self.hidden_dem = tk.Toplevel(self.root)
        self.hidden_dem.withdraw()
        self.dem_app = dem.DemolitionReturnApp(self.hidden_dem)
        
        self.setup_ui()
        
    def setup_ui(self):
        # 헤더
        header = tk.Frame(self.root, bg="#1e3a8a", height=80)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="이동재고 & 철거반납 통합 시스템", font=("Malgun Gothic", 14, "bold"),
                 bg="#1e3a8a", fg="white").pack(pady=22)

        # 바디
        body = tk.Frame(self.root, bg="#0f172a")
        body.pack(fill="both", expand=True, padx=25, pady=20)
        
        tk.Label(body, text="📁 작업 디렉토리 (다운로드 경로)", font=("Malgun Gothic", 9, "bold"),
                 bg="#0f172a", fg="#94a3b8").pack(anchor="w", pady=(0, 4))
                 
        self.path_var = tk.StringVar(value=r"C:\Users\user\Downloads")
        path_entry = tk.Entry(body, textvariable=self.path_var, font=("Malgun Gothic", 10),
                               bg="#1e293b", fg="#e2e8f0", insertbackground="white", relief="flat")
        path_entry.pack(fill="x", ipady=8, pady=(0, 15))
        
        # 버튼 프레임
        btn_frame = tk.Frame(body, bg="#0f172a")
        btn_frame.pack(fill="x", pady=10)
        
        btn_all = tk.Button(
            btn_frame, text="▶️ 한번에 실행 (이동재고 + 철거반납)", command=self.run_all,
            font=("Malgun Gothic", 11, "bold"), bg="#8b5cf6", fg="white", activebackground="#7c3aed",
            relief="flat", cursor="hand2"
        )
        btn_all.pack(fill="x", ipady=12, pady=(0, 10))
        
        sub_btn_frame = tk.Frame(btn_frame, bg="#0f172a")
        sub_btn_frame.pack(fill="x")
        
        btn_inv = tk.Button(
            sub_btn_frame, text="📦 이동재고 실행", command=self.run_inventory,
            font=("Malgun Gothic", 10, "bold"), bg="#3b82f6", fg="white", activebackground="#2563eb",
            relief="flat", cursor="hand2"
        )
        btn_inv.pack(side="left", fill="x", expand=True, ipady=10, padx=(0, 5))
        
        btn_dem = tk.Button(
            sub_btn_frame, text="🛠️ 철거반납 실행", command=self.run_demolition,
            font=("Malgun Gothic", 10, "bold"), bg="#10b981", fg="white", activebackground="#059669",
            relief="flat", cursor="hand2"
        )
        btn_dem.pack(side="right", fill="x", expand=True, ipady=10, padx=(5, 0))
        
        # 로그 콘솔
        tk.Label(body, text="📝 진행 상황", font=("Malgun Gothic", 9, "bold"),
                 bg="#0f172a", fg="#94a3b8").pack(anchor="w", pady=(15, 4))
        
        self.console = scrolledtext.ScrolledText(body, height=10, bg="#1e293b", fg="#a7f3d0", 
                                                 font=("Consolas", 9), relief="flat", state="disabled")
        self.console.pack(fill="both", expand=True)
        
    def log(self, message):
        # run_sync/generate_html_report는 백그라운드 스레드에서 실행되므로
        # Tkinter 위젯 접근은 반드시 root.after()로 메인 스레드에 위임합니다.
        self.root.after(0, self._log_on_main_thread, message)

    def _log_on_main_thread(self, message):
        self.console.configure(state="normal")
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.console.insert("end", f"[{timestamp}] {message}\n")
        self.console.see("end")
        self.console.configure(state="disabled")

    def _show_info(self, title, msg):
        self.root.after(0, lambda: messagebox.showinfo(title, msg))

    def _show_warn(self, title, msg):
        self.root.after(0, lambda: messagebox.showwarning(title, msg))

    def _fail(self, reason):
        self.log(f"=== 작업 실패: {reason} ===")
        self._show_warn("실패", f"{reason}\n진행 상황 로그에서 자세한 오류 내용을 확인하세요.")

    def sync_paths(self):
        # 입력된 경로를 하위 앱들에 동기화 (스레드 시작 전, 메인 스레드에서 호출해야 함)
        path = self.path_var.get()
        self.inv_app.path_var.set(path)
        self.dem_app.path_var.set(path)

    def _execute_with_monkeypatch(self, task_name, logic_func):
        """기존 스크립트의 messagebox를 가로채어 콘솔에 출력하도록 몽키패치"""
        import inventory_html_generator as inv_mod
        import demolition_return_html_generator as dem_mod
        from tkinter import simpledialog

        # 이 함수는 threading.Thread의 target으로 실행되는 워커 스레드다.
        # win32com은 스레드별로 COM 아파트먼트를 초기화해야 하므로 반드시 필요하다.
        pythoncom.CoInitialize()

        original_inv_info = inv_mod.messagebox.showinfo
        original_inv_warn = inv_mod.messagebox.showwarning
        original_inv_err = inv_mod.messagebox.showerror
        original_inv_ask = inv_mod.simpledialog.askstring

        original_dem_info = dem_mod.messagebox.showinfo
        original_dem_warn = dem_mod.messagebox.showwarning
        original_dem_err = dem_mod.messagebox.showerror
        original_dem_ask = dem_mod.simpledialog.askstring

        # 패치 함수들
        def mock_info(title, msg): self.log(f"[INFO] {title}: {msg}")
        def mock_warn(title, msg): self.log(f"[WARN] {title}: {msg}")
        def mock_err(title, msg): self.log(f"[ERROR] {title}: {msg}")
        def mock_ask(title, prompt, **kwargs): return inv_mod.ADMIN_PASSWORD  # 자동 관리자 인증 (두 모듈의 ADMIN_PASSWORD는 항상 동일)

        try:
            # 이동재고 패치
            inv_mod.messagebox.showinfo = mock_info
            inv_mod.messagebox.showwarning = mock_warn
            inv_mod.messagebox.showerror = mock_err
            inv_mod.simpledialog.askstring = mock_ask
            
            # 철거반납 패치
            dem_mod.messagebox.showinfo = mock_info
            dem_mod.messagebox.showwarning = mock_warn
            dem_mod.messagebox.showerror = mock_err
            dem_mod.simpledialog.askstring = mock_ask

            self.log(f"--- {task_name} 시작 ---")
            logic_func()
            self.log(f"--- {task_name} 완료 ---")
        except Exception as e:
            self.log(f"[치명적 오류] {str(e)}")
        finally:
            # 복원
            inv_mod.messagebox.showinfo = original_inv_info
            inv_mod.messagebox.showwarning = original_inv_warn
            inv_mod.messagebox.showerror = original_inv_err
            inv_mod.simpledialog.askstring = original_inv_ask

            dem_mod.messagebox.showinfo = original_dem_info
            dem_mod.messagebox.showwarning = original_dem_warn
            dem_mod.messagebox.showerror = original_dem_err
            dem_mod.simpledialog.askstring = original_dem_ask
            pythoncom.CoUninitialize()

    def run_all_logic(self):
        self.log(">> 이동재고 1단계 (RAW 반영) 실행 중...")
        if not self.inv_app.run_sync():
            self._fail("이동재고 1단계(RAW 반영)에서 실패했습니다.")
            return

        self.log(">> 이동재고 2단계 (보고서 생성) 실행 중...")
        if not self.inv_app.generate_html_report():
            self._fail("이동재고 2단계(보고서 생성)에서 실패했습니다.")
            return

        self.log(">> 철거반납 1단계 (RAW 반영) 실행 중...")
        if not self.dem_app.run_sync():
            self._fail("철거반납 1단계(RAW 반영)에서 실패했습니다.")
            return

        self.log(">> 철거반납 2단계 (보고서 생성) 실행 중...")
        if not self.dem_app.generate_html_report():
            self._fail("철거반납 2단계(보고서 생성)에서 실패했습니다.")
            return

        self.log("=== 모든 작업이 성공적으로 완료되었습니다 ===")
        self._show_info("완료", "모든 작업이 완료되었습니다.\n상세 내용은 진행 상황 로그를 확인하세요.")

    def run_all(self):
        self.sync_paths()
        threading.Thread(target=self._execute_with_monkeypatch, args=("한번에 실행", self.run_all_logic), daemon=True).start()

    def run_inventory_logic(self):
        self.log(">> 이동재고 1단계 (RAW 반영) 실행 중...")
        if not self.inv_app.run_sync():
            self._fail("이동재고 1단계(RAW 반영)에서 실패했습니다.")
            return

        self.log(">> 이동재고 2단계 (보고서 생성) 실행 중...")
        if not self.inv_app.generate_html_report():
            self._fail("이동재고 2단계(보고서 생성)에서 실패했습니다.")
            return

        self._show_info("완료", "이동재고 작업이 완료되었습니다.")

    def run_inventory(self):
        self.sync_paths()
        threading.Thread(target=self._execute_with_monkeypatch, args=("이동재고 실행", self.run_inventory_logic), daemon=True).start()

    def run_demolition_logic(self):
        self.log(">> 철거반납 1단계 (RAW 반영) 실행 중...")
        if not self.dem_app.run_sync():
            self._fail("철거반납 1단계(RAW 반영)에서 실패했습니다.")
            return

        self.log(">> 철거반납 2단계 (보고서 생성) 실행 중...")
        if not self.dem_app.generate_html_report():
            self._fail("철거반납 2단계(보고서 생성)에서 실패했습니다.")
            return

        self._show_info("완료", "철거반납 작업이 완료되었습니다.")

    def run_demolition(self):
        self.sync_paths()
        threading.Thread(target=self._execute_with_monkeypatch, args=("철거반납 실행", self.run_demolition_logic), daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = UnifiedDashboardApp(root)
    root.mainloop()
