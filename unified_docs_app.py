import datetime
import tkinter as tk
from tkinter import messagebox, ttk

from tkcalendar import DateEntry

import faks_gui_finish as faks
import nariad_gui_final as naryad
import plan_rabot_GUI as plan


class NaryadTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=18)
        self._build_ui()

    def _build_ui(self):
        ttk.Label(self, text="Станция:").grid(row=0, column=0, padx=8, pady=8, sticky="e")
        self.station_var = tk.StringVar(value=naryad.stations[0])
        self.station_combo = ttk.Combobox(self, textvariable=self.station_var, values=naryad.stations, width=52, state="readonly")
        self.station_combo.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        ttk.Label(self, text="Руководитель работ:").grid(row=1, column=0, padx=8, pady=8, sticky="e")
        self.leader_var = tk.StringVar(value=naryad.leaders[0])
        self.leader_combo = ttk.Combobox(self, textvariable=self.leader_var, values=naryad.leaders, width=52, state="readonly")
        self.leader_combo.grid(row=1, column=1, padx=8, pady=8, sticky="w")

        today = datetime.date.today()
        ttk.Label(self, text="Начало:").grid(row=2, column=0, padx=8, pady=8, sticky="e")
        self.start_cal = DateEntry(self, width=18, date_pattern="dd.mm.yyyy", year=today.year, month=today.month, day=today.day)
        self.start_cal.grid(row=2, column=1, padx=8, pady=8, sticky="w")

        end_default = today + datetime.timedelta(days=14)
        ttk.Label(self, text="Окончание:").grid(row=3, column=0, padx=8, pady=8, sticky="e")
        self.end_cal = DateEntry(
            self,
            width=18,
            date_pattern="dd.mm.yyyy",
            year=end_default.year,
            month=end_default.month,
            day=end_default.day,
        )
        self.end_cal.grid(row=3, column=1, padx=8, pady=8, sticky="w")

        self.filename_label = ttk.Label(self, text="", style="Hint.TLabel")
        self.filename_label.grid(row=4, column=0, columnspan=2, padx=8, pady=(8, 2), sticky="w")

        ttk.Button(self, text="Создать наряд", command=self.create_document, style="Accent.TButton").grid(
            row=5, column=1, padx=8, pady=16, sticky="w"
        )

        self.station_var.trace_add("write", self.update_filename)
        self.leader_var.trace_add("write", self.update_filename)
        self.update_filename()

        self.columnconfigure(1, weight=1)

    def update_filename(self, *_):
        station = self.station_var.get().strip()
        leader = self.leader_var.get().strip()
        if not station or not leader:
            self.filename_label.config(text="Файл сохранится как: ...")
            return
        filename, _ = naryad.build_naryad_filename(station, leader)
        self.filename_label.config(text=f"Файл сохранится как: {filename}")

    def create_document(self):
        try:
            filename = naryad.generate_naryad_document(
                self.station_var.get(),
                self.leader_var.get(),
                self.start_cal.get_date(),
                self.end_cal.get_date(),
            )
            messagebox.showinfo("Готово", f"Наряд создан:\n{filename}")
        except Exception as error:
            messagebox.showerror("Ошибка", str(error))


class PlanTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=18)
        self._build_ui()

    def _build_ui(self):
        ttk.Label(self, text="Станция:").grid(row=0, column=0, padx=8, pady=8, sticky="e")
        self.station_var = tk.StringVar(value=plan.STATIONS[0])
        self.station_combo = ttk.Combobox(self, textvariable=self.station_var, values=plan.STATIONS, width=52, state="readonly")
        self.station_combo.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        ttk.Label(self, text="Руководитель:").grid(row=1, column=0, padx=8, pady=8, sticky="e")
        self.supervisor_var = tk.StringVar(value=plan.SUPERVISORS[0])
        self.supervisor_combo = ttk.Combobox(
            self,
            textvariable=self.supervisor_var,
            values=plan.SUPERVISORS,
            width=52,
            state="readonly",
        )
        self.supervisor_combo.grid(row=1, column=1, padx=8, pady=8, sticky="w")

        today = datetime.date.today()
        ttk.Label(self, text="Начало:").grid(row=2, column=0, padx=8, pady=8, sticky="e")
        self.start_cal = DateEntry(self, width=18, date_pattern="dd.mm.yyyy", year=today.year, month=today.month, day=today.day)
        self.start_cal.grid(row=2, column=1, padx=8, pady=8, sticky="w")

        end_default = today + datetime.timedelta(days=10)
        ttk.Label(self, text="Окончание:").grid(row=3, column=0, padx=8, pady=8, sticky="e")
        self.end_cal = DateEntry(
            self,
            width=18,
            date_pattern="dd.mm.yyyy",
            year=end_default.year,
            month=end_default.month,
            day=end_default.day,
        )
        self.end_cal.grid(row=3, column=1, padx=8, pady=8, sticky="w")

        ttk.Button(self, text="Создать план работ", command=self.create_document, style="Accent.TButton").grid(
            row=4, column=1, padx=8, pady=16, sticky="w"
        )
        self.columnconfigure(1, weight=1)

    def create_document(self):
        try:
            filename = plan.generate_document(
                station=self.station_var.get(),
                supervisor_full=self.supervisor_var.get(),
                start_dt=self.start_cal.get_date(),
                end_dt=self.end_cal.get_date(),
                show_messages=False,
            )
            messagebox.showinfo("Готово", f"План работ создан:\n{filename}")
        except Exception as error:
            messagebox.showerror("Ошибка", str(error))


def _setup_styles(root):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("Root.TFrame", background="#eff3f8")
    style.configure("Header.TFrame", background="#17324f")
    style.configure("Header.TLabel", background="#17324f", foreground="#ffffff", font=("Segoe UI", 14, "bold"))
    style.configure("SubHeader.TLabel", background="#17324f", foreground="#d7e7fb", font=("Segoe UI", 10))

    style.configure("TNotebook", background="#eff3f8", borderwidth=0)
    style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(16, 8), background="#c9d8eb")
    style.map("TNotebook.Tab", background=[("selected", "#ffffff")])

    style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 8), foreground="#ffffff", background="#1f6f43")
    style.map("Accent.TButton", background=[("active", "#185c37")])
    style.configure("Hint.TLabel", foreground="#425466")


def run_app():
    root = tk.Tk()
    root.title("Единый генератор документов")
    root.geometry("1060x820")
    root.minsize(980, 760)

    _setup_styles(root)

    container = ttk.Frame(root, style="Root.TFrame")
    container.pack(fill=tk.BOTH, expand=True)

    header = ttk.Frame(container, style="Header.TFrame", padding=(18, 14))
    header.pack(fill=tk.X)
    ttk.Label(header, text="Единый генератор документов", style="Header.TLabel").pack(anchor="w")
    ttk.Label(
        header,
        text="Исходящее письмо, наряд и план работ в одном приложении",
        style="SubHeader.TLabel",
    ).pack(anchor="w", pady=(4, 0))

    notebook = ttk.Notebook(container)
    notebook.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)

    letter_tab = ttk.Frame(notebook)
    naryad_tab = NaryadTab(notebook)
    plan_tab = PlanTab(notebook)

    letter_app = faks.LetterApp(letter_tab)

    notebook.add(letter_tab, text="Исходящее письмо")
    notebook.add(naryad_tab, text="Наряд")
    notebook.add(plan_tab, text="План работ")

    root.letter_app = letter_app
    root.mainloop()


if __name__ == "__main__":
    run_app()
