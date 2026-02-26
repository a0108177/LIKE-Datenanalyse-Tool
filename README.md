# LIKE Datenanalyse Tool

## Dateiübersicht

- **app.py** – GUI-Anwendung mit tkinter. Ermöglicht CSV-Upload via Dialog oder Drag & Drop und führt die Datenanalyse aus.
- **LIKE.py** – Core-Funktion `like()` für die Datenverarbeitung und Excel-Export.
- **LIKE.ipynb** – Jupyter Notebook mit Entwicklungs-/Testcode (optional, dient als Referenz).

## Installation

### Erforderliche Pakete

```bash
pip install pandas openpyxl tkinterdnd2 python-pptx
```

### Optional (für Notebook)
```bash
pip install jupyter
```

## Anwendung starten

```bash
python app.py
```

## One-File-Anwendung (.exe)

Zum Erstellen einer Windows-Anwendung als einzelne .exe-Datei:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "LIKE_Tool" --icon=icon.ico app.py
```

**Optionen:**
- `--onefile` – Alles in einer Datei
- `--windowed` – Keine Konsole anzeigen
- `--icon=icon.ico` – App-Icon setzen (optional)

Die .exe befindet sich danach in `dist/LIKE_Tool.exe`.

## Verwendung

1. CSV-Dateien hochladen (Pflicht: Large Table, MCP, Self-Assessment; Optional: MDLO)
2. Dateien per Dialog oder Drag & Drop ablegen
3. Ausgabeordner optional ändern
4. "Analyse starten" klicken
5. Excel-Datei wird mit Ergebnissen erstellt und kann direkt geöffnet werden
