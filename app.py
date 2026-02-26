# Import der erforderlichen Module für CSV-Verarbeitung, Dateisystem, GUI und Datenanalyse
import csv
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd

from LIKE import like

# Normalisiert den Pfad aus Drag & Drop Daten


def _normalize_dnd_path(data: str) -> str:
    data = data.strip()
    if data.startswith("{") and "} {" in data:
        first = data.split("} {", 1)[0].lstrip("{")
        return first
    if data.startswith("{") and data.endswith("}"):
        return data[1:-1]
    return data


# Überprüft, ob der Pfad eine CSV-Datei ist
def _is_csv(path: str) -> bool:
    return path.lower().endswith(".csv")


# Erstellt das Standard-Ausgabeverzeichnis im Downloads-Ordner
def _default_output_dir() -> Path:
    downloads = Path.home() / "Downloads"
    outdir = downloads / "LIKE_Output"
    outdir.mkdir(parents=True, exist_ok=True)
    return outdir


# Liest den Header einer CSV-Datei und gibt ihn als Set zurück
def read_csv_header(path: str) -> set[str]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader)
    return {h.strip().lower() for h in header if h and h.strip()}


# Signaturen (aus .csv Headern)
SIG_LARGE = {
    "completion status", "sum time spent", "accuracy-classes",
    "unconscious incompetent", "conscious incompetent",
    "unconscious competent", "conscious competent",
}

SIG_MCP = {
    "initial conscious competence", "initial unconscious competence",
    "improvement conscious competence", "improvement unconscious incompetence",
    "current conscious competence", "current unconscious competence",
}

SIG_SELF = {
    "self assessment", "average progress", "time", "correct", "wrong", "accuracy",
}

SIG_MDLO = {
    "learning objective", "unconsciously incompetent", "wrong answers", "open in curator",
}


# Erkennt den CSV-Typ basierend auf dem Header
def detect_csv_type_by_header(path: str) -> tuple[str | None, str]:
    cols = read_csv_header(path)

    # Large Table
    if {"learner", "module", "completion status"}.issubset(cols) and len(SIG_LARGE & cols) >= 4:
        return "large", "Header-Signatur: Large Table"

    # MCP
    if {"class name", "progress"}.issubset(cols) and len(SIG_MCP & cols) >= 4:
        return "mcp", "Header-Signatur: Metacognition Progress"

    # Self Assessment
    if {"learner", "module", "self assessment"}.issubset(cols) and len(SIG_SELF & cols) >= 4:
        return "self", "Header-Signatur: Self-Assessment"

    # MDLO
    if {"module", "learning objective"}.issubset(cols) and len(SIG_MDLO & cols) >= 3:
        return "mdlo", "Header-Signatur: MDLO"

    return None, f"Unbekannter CSV-Header. Gefundene Spalten (Auszug): {', '.join(sorted(list(cols))[:8])}..."


# Hauptklasse der Anwendung
class LikeApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("LIKE Datenanalyse Tool")
        self.geometry("900x520")
        self.minsize(900, 520)

        # Datei-Pfade
        self.paths = {
            "large": None,
            "mcp": None,
            "self": None,
            "mdlo": None,
        }

        # Output
        self.output_dir = _default_output_dir()
        self.last_output_file = None

        self._build_ui()
        self._refresh_state()

    # Baut die Benutzeroberfläche auf
    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        header = ttk.Label(
            self,
            text="CSV-Dateien hochladen oder per Drag & Drop im Feld ablegen",
            font=("Segoe UI", 13, "bold"),
        )
        header.pack(anchor="w", **pad)

        hint = ttk.Label(
            self,
            text=(
                "Pflicht: Large Table, Metacognition Progress, Self-Assessment\n"
                "Optional: Most Difficult Learning Objectives"
            ),
        )
        hint.pack(anchor="w", padx=12)

        # Drop-Zone
        self.drop_zone = ttk.Label(
            self,
            text="Datei hierher ziehen (CSV)",
            relief="ridge",
            padding=18,
            anchor="center",
        )
        self.drop_zone.pack(fill="x", padx=12, pady=12)

        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

        # File rows container
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=6)

        frm.columnconfigure(1, weight=1)

        # Rows
        self._row_large = self._file_row(
            parent=frm,
            r=0,
            key="large",
            title="1) Large Table (Pflicht)",
            button_text="Datei wählen…",
        )
        self._row_mcp = self._file_row(
            parent=frm,
            r=1,
            key="mcp",
            title="2) Metacognition Progress (Pflicht)",
            button_text="Datei wählen…",
        )
        self._row_self = self._file_row(
            parent=frm,
            r=2,
            key="self",
            title="3) Self-Assessment (Pflicht)",
            button_text="Datei wählen…",
        )
        self._row_mdlo = self._file_row(
            parent=frm,
            r=3,
            key="mdlo",
            title="4) Most Difficult Learning Objectives (Optional)",
            button_text="Datei wählen…",
        )

        # Output info + change button
        outfrm = ttk.Frame(self)
        outfrm.pack(fill="x", padx=12, pady=6)

        outfrm.columnconfigure(0, weight=1)

        self.output_label = ttk.Label(
            outfrm,
            text=f"Ausgabeordner: {str(self.output_dir)}",
        )
        self.output_label.grid(row=0, column=0, sticky="w")

        self.btn_change_output = ttk.Button(
            outfrm, text="Ausgabeordner ändern…", command=self._choose_output_dir
        )
        self.btn_change_output.grid(row=0, column=1, sticky="e")

        # Actions
        actions = ttk.Frame(self)
        actions.pack(fill="x", padx=12, pady=10)

        self.btn_run = ttk.Button(
            actions, text="Analyse starten", command=self._run)
        self.btn_run.pack(side="left")

        self.btn_open_folder = ttk.Button(
            actions, text="Ordner öffnen", command=self._open_output_folder, state="disabled"
        )
        self.btn_open_folder.pack(side="left", padx=(10, 0))

        self.btn_reset = ttk.Button(
            actions, text="Neustart", command=self._reset)
        self.btn_reset.pack(side="right")

        # Status
        self.status = tk.StringVar(value="Bereit.")
        self.status_label = ttk.Label(self, textvariable=self.status)
        self.status_label.pack(anchor="w", padx=12, pady=(0, 10))

    # Erstellt eine Zeile für die Dateiauswahl in der UI
    def _file_row(self, parent: ttk.Frame, r: int, key: str, title: str, button_text: str):
        lbl_title = ttk.Label(parent, text=title)
        lbl_title.grid(row=r, column=0, sticky="w", padx=(0, 8), pady=6)

        var = tk.StringVar(value="(keine Datei)")
        lbl_path = ttk.Label(parent, textvariable=var)
        lbl_path.grid(row=r, column=1, sticky="we", pady=6)

        btn = ttk.Button(
            parent,
            text=button_text,
            command=lambda k=key: self._choose_file(k),
        )
        btn.grid(row=r, column=2, sticky="e", padx=(8, 0), pady=6)

        btn_clear = ttk.Button(
            parent,
            text="Entfernen",
            command=lambda k=key: self._clear_file(k),
        )
        btn_clear.grid(row=r, column=3, sticky="e", padx=(8, 0), pady=6)

        return {"var": var, "lbl": lbl_path, "btn": btn, "btn_clear": btn_clear}

    # Behandelt das Ablegen von Dateien per Drag & Drop
    def _on_drop(self, event):
        path = _normalize_dnd_path(event.data)

        if not path or not os.path.isfile(path):
            messagebox.showerror(
                "Ungültig", "Bitte eine gültige Datei ablegen.")
            return
        if not _is_csv(path):
            messagebox.showerror(
                "Falsches Format", "Bitte eine CSV-Datei ablegen.")
            return

        try:
            key, reason = detect_csv_type_by_header(path)
        except Exception as e:
            messagebox.showerror(
                "Fehler", f"Konnte CSV-Header nicht lesen:\n{e}")
            return

        if key is None:
            messagebox.showerror(
                "Unbekannte CSV",
                "Die Datei konnte keinem Typ zugeordnet werden.\n"
                "Bitte prüfe, ob es eine der vier Exportdateien ist.\n\n"
                f"Datei: {Path(path).name}\n{reason}"
            )
            return

        if self.paths[key] is not None:
            overwrite = messagebox.askyesno(
                "Datei bereits gesetzt",
                f"{key.upper()} ist bereits gesetzt.\n\n"
                f"Neue Datei: {Path(path).name}\n"
                f"Erkennung: {reason}\n\n"
                "Überschreiben?"
            )
            if not overwrite:
                return

        self._set_file(key, path)
        self.status.set(
            f"Zugeordnet zu {key.upper()} – {Path(path).name} ({reason})")

    # Öffnet einen Dateidialog zur Auswahl einer CSV-Datei
    def _choose_file(self, key: str):
        path = filedialog.askopenfilename(
            title="CSV auswählen",
            filetypes=[("CSV Dateien", "*.csv"), ("Alle Dateien", "*.*")],
        )
        if not path:
            return
        if not _is_csv(path):
            messagebox.showerror(
                "Falsches Format", "Bitte eine CSV-Datei auswählen.")
            return
        self._set_file(key, path)
        self.status.set(f"Datei gesetzt: {Path(path).name}")

    # Öffnet einen Dialog zur Auswahl des Ausgabeordners
    def _choose_output_dir(self):
        d = filedialog.askdirectory(title="Ausgabeordner auswählen")
        if not d:
            return
        self.output_dir = Path(d)
        self.output_label.config(text=f"Ausgabeordner: {str(self.output_dir)}")
        self.status.set("Ausgabeordner geändert.")
        self._refresh_state()

    # Öffnet den Ausgabeordner im Datei-Explorer
    def _open_output_folder(self):
        try:
            os.startfile(str(self.output_dir))
        except Exception as e:
            messagebox.showerror(
                "Fehler", f"Ordner konnte nicht geöffnet werden:\n{e}")

    # ----- State -----
    # Setzt den Pfad für einen bestimmten Dateityp
    def _set_file(self, key: str, path: str):
        self.paths[key] = path
        row = self._row_for_key(key)
        row["var"].set(Path(path).name)
        self._refresh_state()

    # Entfernt die Datei für einen bestimmten Typ
    def _clear_file(self, key: str):
        self.paths[key] = None
        row = self._row_for_key(key)
        row["var"].set("(keine Datei)")
        self.status.set("Datei entfernt.")
        self._refresh_state()

    # Gibt die UI-Zeile für einen Schlüssel zurück
    def _row_for_key(self, key: str):
        return {
            "large": self._row_large,
            "mcp": self._row_mcp,
            "self": self._row_self,
            "mdlo": self._row_mdlo,
        }[key]

    # Überprüft, ob alle Pflichtdateien gesetzt sind
    def _mandatory_ready(self) -> bool:
        return all(self.paths[k] for k in ["large", "mcp", "self"])

    # Aktualisiert den Zustand der UI basierend auf den gesetzten Dateien
    def _refresh_state(self):
        self.btn_run.config(
            state=("normal" if self._mandatory_ready() else "disabled"))

    # Setzt die Anwendung zurück
    def _reset(self):
        for k in self.paths:
            self.paths[k] = None
        self._row_large["var"].set("(keine Datei)")
        self._row_mcp["var"].set("(keine Datei)")
        self._row_self["var"].set("(keine Datei)")
        self._row_mdlo["var"].set("(keine Datei)")
        self.btn_open_folder.config(state="disabled")
        self.last_output_file = None
        self.status.set("Zurückgesetzt. Bitte Dateien auswählen oder ablegen.")
        self._refresh_state()

    # ----- Run -----
    # Führt die Datenanalyse aus
    def _run(self):
        try:
            self.status.set("CSV-Dateien werden geladen…")

            df_large = pd.read_csv(
                self.paths["large"], sep=";", encoding="utf-8")
            df_mcp = pd.read_csv(self.paths["mcp"], sep=";", encoding="utf-8")
            df_self = pd.read_csv(
                self.paths["self"], sep=";", encoding="utf-8")

            df_mdlo = None
            if self.paths["mdlo"]:
                df_mdlo = pd.read_csv(
                    self.paths["mdlo"], sep=";", encoding="utf-8")

            if df_mdlo is None:
                df_mdlo = pd.DataFrame()

            self.status.set("Datenanalyse läuft…")
            old_cwd = os.getcwd()
            os.chdir(self.output_dir)

            try:
                like(df_large, df_mcp, df_self, df_mdlo)
            finally:
                os.chdir(old_cwd)

            self.status.set(
                f"Fertig. Ergebnis liegt in: {str(self.output_dir)}")
            self.btn_open_folder.config(state="normal")
            messagebox.showinfo(
                "Erfolg",
                f"Datenanalyse erfolgreich.\nDie Datei wurde in folgendem Ordner gespeichert:\n{self.output_dir}",
            )

        except Exception as e:
            self.status.set("Fehler aufgetreten.")
            messagebox.showerror("Fehler", f"Fehler in der Datenanalyse:\n{e}")


# Startet die Anwendung
if __name__ == "__main__":
    app = LikeApp()
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    app.mainloop()
