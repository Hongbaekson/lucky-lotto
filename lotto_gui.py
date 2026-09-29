"""표준 라이브러리 Tkinter로 실행하는 로또 분석 창."""

from pathlib import Path
from queue import Empty, Queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import lotto


class LottoApp:
    def __init__(self, root):
        self.root = root
        self.events = Queue()
        self.busy = False
        root.title("Lucky Lotto · 당첨 이력과 번호 분석")
        root.geometry("980x860")
        root.minsize(880, 760)
        root.configure(bg="#F3F6FA")
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#F3F6FA")
        style.configure("TLabel", background="#F3F6FA", font=("맑은 고딕", 10))
        style.configure("TButton", font=("맑은 고딕", 10), padding=(14, 9))
        style.configure("Accent.TButton", background="#175D70", foreground="white")
        style.map("Accent.TButton", background=[("active", "#124858"), ("disabled", "#8899A4")])
        style.configure("TEntry", padding=5)
        body = ttk.Frame(root, padding=24)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="LUCKY LOTTO", font=("맑은 고딕", 24, "bold"), foreground="#153E50").pack(anchor="w")
        ttk.Label(body, text="당첨 이력 심층 분석 · 회차별 고정 2게임", foreground="#526575").pack(anchor="w", pady=(4, 16))
        self.path = tk.StringVar(value=str(lotto.DEFAULT_FILE))
        file_row = ttk.Frame(body)
        file_row.pack(fill="x")
        self.path_entry = ttk.Entry(file_row, textvariable=self.path)
        self.path_entry.pack(side="left", fill="x", expand=True)
        self.browse = ttk.Button(file_row, text="엑셀 선택", command=self.select_file)
        self.browse.pack(side="left", padx=(10, 0))
        self.status = tk.StringVar(value="이력을 불러오는 중...")
        ttk.Label(body, textvariable=self.status, foreground="#175D70").pack(anchor="w", pady=(10, 18))

        self.weekly_title = tk.StringVar(value="이번 회차 2게임")
        self.weekly_numbers = tk.StringVar(value="‘이번 회차 2게임 저장’을 누르면 최신 이력을 확인하고 분석합니다.")
        ttk.Label(body, textvariable=self.weekly_title, font=("맑은 고딕", 13, "bold"), foreground="#153E50").pack(anchor="w")
        ttk.Label(body, textvariable=self.weekly_numbers, font=("맑은 고딕", 12), foreground="#175D70").pack(anchor="w", pady=(5, 14))

        settings = ttk.Frame(body)
        settings.pack(fill="x")
        self.mode = tk.StringVar(value=lotto.MODES["recent"])
        self.games = tk.StringVar(value="2")
        self.window = tk.StringVar(value="52")
        self.test_rounds = tk.StringVar(value="104")
        ttk.Label(settings, text="추천 방식").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.mode_input = ttk.Combobox(settings, textvariable=self.mode, values=list(lotto.MODES.values()), state="readonly", width=19)
        self.mode_input.grid(row=1, column=0, sticky="w", padx=(0, 20))
        self.inputs = [self.path_entry, self.browse, self.mode_input]
        for column, (label, variable, maximum) in enumerate([
                ("게임 수", self.games, 100), ("최근 분석 회차 수", self.window, 2000),
                ("과거 검증 회차 수", self.test_rounds, 2000)], 1):
            ttk.Label(settings, text=label).grid(row=0, column=column, sticky="w", pady=(0, 5))
            entry = ttk.Spinbox(settings, from_=1, to=maximum, textvariable=variable, width=12)
            entry.grid(row=1, column=column, sticky="w", padx=(0, 20))
            self.inputs.append(entry)

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=20)
        self.weekly_button = ttk.Button(actions, text="이번 회차 2게임 저장", style="Accent.TButton", command=self.weekly)
        self.weekly_button.pack(side="left", padx=(0, 10))
        self.update_button = ttk.Button(actions, text="이력 업데이트", command=self.update)
        self.update_button.pack(side="left", padx=(0, 10))
        self.report_button = ttk.Button(actions, text="설정대로 추가 분석", command=self.report)
        self.report_button.pack(side="left")
        self.inputs.extend([self.weekly_button, self.update_button, self.report_button])
        ttk.Label(body, text="위 설정은 ‘추가 분석’에만 적용됩니다. 주간 추천은 분석 결과에 따라 2게임을 고정합니다.", foreground="#526575").pack(anchor="w", pady=(0, 8))
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 12))
        self.output = scrolledtext.ScrolledText(body, wrap="word", font=("맑은 고딕", 11),
                                              bg="white", fg="#203544", relief="flat", padx=14, pady=12,
                                              state="disabled", height=15)
        self.output.pack(fill="both", expand=True)
        ttk.Label(body, text=lotto.NOTICE, wraplength=810, foreground="#526575").pack(anchor="w", pady=(14, 0))
        self.append("엑셀을 닫은 뒤 ‘이번 회차 2게임 저장’을 눌러 주세요.\n공식 데이터 확인 → 심층 분석 → 2게임 저장까지 진행합니다.\n같은 회차는 처음 저장한 번호를 유지합니다. 결과는 reports 폴더에 저장됩니다.\n")
        self.refresh_status()
        root.after(100, self.poll)

    def append(self, message):
        self.output.configure(state="normal")
        self.output.insert("end", str(message) + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")

    def refresh_status(self):
        try:
            draws = lotto.load_draws(Path(self.path.get()))
            self.status.set(f"엑셀 기준: 1~{draws[-1].round}회 · 마지막 추첨 {draws[-1].drawn_at:%Y-%m-%d} · 추천 대상 {draws[-1].round + 1}회")
        except Exception as exc:
            self.status.set(f"엑셀 확인 필요: {exc}")

    def select_file(self):
        selected = filedialog.askopenfilename(title="당첨 이력 엑셀 선택", filetypes=[("Excel", "*.xlsx")], initialdir=lotto.BASE)
        if selected:
            self.path.set(selected)
            self.refresh_status()

    def start(self, action):
        if self.busy:
            return
        self.busy = True
        for widget in self.inputs:
            widget.configure(state="disabled")
        self.progress.start(15)
        self.append("─" * 45)

        def work():
            try:
                action()
            except Exception as exc:
                self.events.put(("error", str(exc)))
            finally:
                self.events.put(("done", None))

        threading.Thread(target=work, daemon=True).start()

    def log(self, message):
        self.events.put(("log", message))

    def update(self):
        path = Path(self.path.get())
        self.start(lambda: lotto.update_history(path, log=self.log))

    def weekly(self):
        path = Path(self.path.get())

        def run():
            from lotto_weekly import weekly
            result = weekly(path, log=self.log)
            self.events.put(("weekly", result["record"]))

        self.start(run)

    def report(self):
        try:
            games = lotto.integer(self.games.get(), "게임 수", 1)
            window = lotto.integer(self.window.get(), "최근 분석 회차 수", 1)
            rounds = lotto.integer(self.test_rounds.get(), "과거 검증 회차 수", 1)
            mode = next(key for key, label in lotto.MODES.items() if label == self.mode.get())
            if games > 100:
                raise lotto.LottoError("게임 수는 1~100이어야 합니다.")
        except (lotto.LottoError, StopIteration) as exc:
            messagebox.showerror("입력 확인", str(exc))
            return
        path = Path(self.path.get())
        self.start(lambda: lotto.create_report(path, games, mode, window, test_rounds=rounds, log=self.log))

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "log":
                    self.append(value)
                elif kind == "weekly":
                    self.weekly_title.set(f"{value['target_round']}회 · 추첨 예정 {value['planned_date']} · 저장된 2게임")
                    self.weekly_numbers.set("\n".join(f"게임 {i}    " + "   ".join(f"{n:02}" for n in ticket)
                                                       for i, ticket in enumerate(value["tickets"], 1)))
                elif kind == "error":
                    self.append(f"오류: {value}\n파일이 열려 있으면 엑셀을 닫고 다시 실행하세요.")
                    messagebox.showerror("작업을 마치지 못했습니다", value)
                else:
                    self.busy = False
                    self.progress.stop()
                    for widget in self.inputs:
                        widget.configure(state="normal")
                    self.mode_input.configure(state="readonly")
                    self.refresh_status()
        except Empty:
            pass
        self.root.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo("작업 진행 중", "파일 저장이 끝난 뒤 창을 닫아 주세요.")
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    LottoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
