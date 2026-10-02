# Persian Date Converter

**Convert Persian (Shamsi) dates to Gregorian in 25 formats across 8 languages — with a GUI, batch converter, and a Word macro.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)
![Language: Python](https://img.shields.io/badge/Language-Python%203.10-blue.svg)

---

## Overview

A complete tool for translators, offices, and anyone who works with Persian dates. Pick a format once, then:

- **Type a date in the GUI** → get the Gregorian equivalent instantly
- **Batch-convert a folder of `.docx` files** with one click
- **Press a keyboard shortcut inside Word** → the date pastes at the cursor

No Python required on the target machine — the executables are self-contained.

---

## Features

| Feature | Description |
|---|---|
| **25 date formats** | British, American, ISO, European, dot-separated, and more |
| **8 languages** | English, French, German, Italian, Spanish, Turkish, Arabic, Russian |
| **Live two-way conversion** | Type Persian → get Gregorian. Type Gregorian → get Persian. |
| **Batch converter** | Convert whole folders of `.docx` files at once |
| **Word macro** | Insert dates at the cursor with a single keystroke |
| **Copy to clipboard** | One click to copy the converted date |
| **Protected documents** | Works inside protected Word forms, prompts for passwords when needed |
| **Activity log** | Every conversion timestamped in `dist\log.txt` |

---

## Download

### Option 1 — Download the ready-to-use package (recommended)

1. Go to the [Releases](../../releases) page.
2. Download the latest `PersianDateTool_vX.X.X.zip`.
3. Extract it so the final folder is at `C:\PersianDateTool\`.
4. Follow `QUICK_START.txt` inside.

### Option 2 — Build from source

Requirements:

- Windows 10/11
- Python 3.10 (or 3.11 / 3.12)
- Git

```cmd
git clone https://github.com/Saeedmg/persian-date-converter.git
cd persian-date-converter
pip install jdatetime python-docx pyinstaller
build_cli.bat
build_gui.bat

```
Both executables are produced in `dist\`.

---

## File Layout

After extraction, the folder should look like this:

```
C:\PersianDateTool\
├── dist\
│   ├── PersianDateConverter.exe                 GUI (double-click)
│   ├── PersianDateConverterCLI\
│   │   └── PersianDateConverterCLI.exe          CLI (called by Word)
│   ├── RunHidden.vbs                            hides the console window
│   ├── settings.ini                             remembers your format
│   └── log.txt                                  conversion history
└── word\
    └── ConvertPersianDate.bas                   macro to import into Word
```

---

## Quick Start (Word Macro)

1. Extract the tool to `C:\PersianDateTool\`.
2. Open `dist\PersianDateConverter.exe` once, pick a format, close it.
3. In Word, press **Alt + F11**.
4. **File → Import File…** → select `word\ConvertPersianDate.bas`.
5. **Ctrl + S** → click **Yes** to save to the Normal template.
6. Bind a shortcut (**File → Options → Customize Ribbon → Keyboard shortcuts**).

Now, in any Word document:

```
Press Ctrl + Shift + D   →   type 1403/06/27   →   OK
```

The converted date pastes at the cursor.

---

## Formats

| Key | Output | Description |
|---|---|---|
| 1 | 17 September 2024 | English — British full |
| 2 | 09/17/2024 | English — American slash |
| 3 | 17-09-2024 | English — European dash |
| 4 | 2024-09-17 | English — ISO |
| 5 | September 17, 2024 | English — American full *(default)* |
| 6 | 17 Sep. 2024 | English — British abbreviated |
| 7 | Sep. 17, 2024 | English — American abbreviated |
| 8 | 17 septembre 2024 | French |
| 9 | 17 sept. 2024 | French abbreviated |
| 10 | 17. September 2024 | German |
| 11 | 17. Sep. 2024 | German abbreviated |
| 12 | 17 settembre 2024 | Italian |
| 13 | 17 set. 2024 | Italian abbreviated |
| 14 | 17 de septiembre de 2024 | Spanish |
| 15 | 17 sep. 2024 | Spanish abbreviated |
| 16 | 17 Eylül 2024 | Turkish |
| 17 | 17 Eyl. 2024 | Turkish abbreviated |
| 18 | 17 سبتمبر 2024 | Arabic |
| 19 | 17/09/2024 | Arabic numeric |
| 20 | 17.09.2024 | Dot — day.month.year |
| 21 | 2024.09.17 | Dot — ISO |
| 22 | 09.17.2024 | Dot — American |
| 23 | 17. September 2024 | Dot — day. Month year |
| 24 | 17 сентября 2024 | Russian |
| 25 | 17 сен. 2024 | Russian abbreviated |

---

## Keyboard Shortcuts (Word)

After binding in Word's Customize dialog:

| Shortcut | Macro | Output |
|---|---|---|
| Ctrl+Shift+D | `ConvertPersianDate` | Uses the format set in the GUI |
| Alt+F8 | `ConvertPersianDate_Full` | September 17, 2024 |
| Alt+F9 | `ConvertPersianDate_Abbrev` | Sep. 17, 2024 |
| Alt+F10 | `ConvertPersianDate_British` | 17 September 2024 |
| Ctrl+Alt+1 | `ConvertPersianDate_BritishAbbrev` | 17 Sep. 2024 |
| Ctrl+Alt+2 | `ConvertPersianDate_ISO` | 2024-09-17 |
| Ctrl+Alt+3 | `ConvertPersianDate_DotISO` | 2024.09.17 |
| Ctrl+Alt+4 | `ConvertPersianDate_SlashUS` | 09/17/2024 |

---

## Batch Converter

Convert whole folders of Word files with one click:

1. Open `dist\PersianDateConverter.exe`.
2. Click the **Word Files (Batch)** tab.
3. Pick an **input folder** (contains your `.docx` files).
4. Pick an **output folder** (where results are saved).
5. Pick a **format** from the dropdown.
6. Click **Convert Word Files**.

Each input file produces an output file named `OriginalName_updated.docx`. Originals are untouched.

Works on paragraphs, tables, headers, and footers.

---

## Protected Documents

The macro handles three cases automatically:

- **Unprotected document** — inserts normally.
- **Form-protected document** — fills the form field directly.
- **Password-protected document** — asks for the password once.

No manual unprotecting needed.

---

## Building from Source

Both executables are built with PyInstaller:

```cmd
build_cli.bat    →  dist\PersianDateConverterCLI\PersianDateConverterCLI.exe
build_gui.bat    →  dist\PersianDateConverter.exe
```

Key build flags:

| Flag | Purpose |
|---|---|
| `--onedir` | CLI: fast startup (no unpacking on each run) |
| `--noconsole` | CLI: no console window when Word calls it |
| `--onefile` | GUI: single-file distribution |
| `--windowed` | GUI: no console |
| `--icon` | Custom icon for both exes |
| `--version-file` | File-properties metadata |
| `--hidden-import=jdatetime` | Ensures jdatetime is bundled |
| `--collect-all=jdatetime` | Bundles jdatetime's internal data |

**Never run `pyinstaller` directly** — always use `python -m PyInstaller` to guarantee the right interpreter.

---

## Documentation

- **`QUICK_START.txt`** — one-page English install guide
- **`Persian Guide.txt`** — Persian install guide
- **`Installation Guide_Persian.docx` / `.pdf`** — fully formatted guide

---

## Compatibility

| Component | Supported |
|---|---|
| Operating System | Windows 10, Windows 11 |
| Microsoft Word | 2016, 2019, 2021, Microsoft 365 |
| Python (build only) | 3.10, 3.11, 3.12 |

---

## Contributing

Suggestions and improvements are welcome. Open an [issue](../../issues) or submit a pull request.

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

## Contact

**Saeed Majidi** — Translation Office 1324
📧 rassamtranslation@gmail.com

---

*Built with Python, `jdatetime`, `python-docx`, and PyInstaller.*

