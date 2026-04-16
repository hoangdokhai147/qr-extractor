# 📷 QR Code Extractor

A production-ready Python tool providing both a **Command Line Interface (CLI)** and a **Graphical User Interface (GUI)**. It **batch-reads QR codes from device images** and exports a structured **Excel report** with per-folder summaries and per-file details — ready to be dropped into any data pipeline or operations workflow.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
  - [macOS](#macos)
  - [Linux (Ubuntu / Debian)](#linux-ubuntu--debian)
  - [Windows](#windows)
- [Project Structure](#project-structure)
- [Usage](#usage)
  - [GUI (Graphical User Interface)](#gui-graphical-user-interface)
  - [CLI Basic](#cli-basic)
  - [CLI All Options](#cli-all-options)
  - [CLI Examples](#cli-examples)
- [Expected Input Layout](#expected-input-layout)
- [Output Format](#output-format)
- [Behavior & Edge Cases](#behavior--edge-cases)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## Features

| Capability | Detail |
|---|---|
| 🖥️ UI / UX | Clean Tkinter GUI to preview extracted data streams before exporting |
| 🗂️ Batch processing | Scans every subfolder inside a root directory |
| 🤖 AI-Powered Decode | Runs `WeChatQRCode` AI models to read highly-reflective tiny tags! |
| 🔎 Comprehensive fallback | Uses advanced blob tracing, upscaling, rotations, and standard `pyzbar` |
| 📊 Excel export | Summary table + Detail table in one sheet (`QR_Results`) |
| 🔢 Multi-QR per image | Multiple codes joined with `; ` |
| 🛡️ Graceful error handling | Corrupted / unreadable images logged as `fail`, never crash |
| 📝 Verbose logging | `--verbose` flag for full per-file progress |
| 🔄 Overwrite guard | Prompts before overwriting an existing output file |

---

## Requirements

- **Python 3.8+**
- **Native `zbar` shared library** (see [Installation](#installation))
- **Internet Connection** (only for the exact first run, to download AI models)
- Python packages (installed via `pip`):

```
opencv-contrib-python
pyzbar
pandas
openpyxl
Pillow
```

---

## Installation

### macOS

```bash
# 1. Install the native zbar library (REQUIRED — pyzbar wraps this)
brew install zbar

# 2. Clone the repository
git clone https://github.com/your-org/qr-extractor.git
cd qr-extractor

# 3. Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 4. Install Python dependencies
pip install -r requirements.txt
```

> **Why `brew install zbar`?**
> `pyzbar` is a thin Python wrapper that calls the native `libzbar.dylib` at
> runtime via `ctypes`. Installing only `pip install pyzbar` puts the Python
> glue code in your venv but **not** the native library — you will get
> `ImportError: Unable to find zbar shared library` without this step.

---

### Linux (Ubuntu / Debian)

```bash
# 1. Install native zbar
sudo apt-get update && sudo apt-get install -y libzbar0

# 2. Clone and set up
git clone https://github.com/hoangdokhai147/qr-extractor.git
cd qr-extractor
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

### Windows

```powershell
# 1. Install native zbar — choose one method:
#    (a) Chocolatey
choco install zbar
#    (b) Manual: download the DLL from https://zbar.sourceforge.net/
#        and add it to your PATH or place it next to qr_extractor.py

# 2. Clone and set up
git clone https://github.com/hoangdokhai147/qr-extractor.git
cd qr-extractor
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

---

## Project Structure

```
qr-extractor/
├── app.py               # Graphical User Interface (Tkinter)
├── main.py              # Command line interface & orchestrator
├── requirements.txt     # Python dependencies
├── README.md            # This file
├── src/
│   ├── qrcode/
│   │   └── scanner.py   # WeChatQRCode AI logic & auto-downloads
│   ├── io/
│   │   ├── file_utils.py
│   │   └── excel_writer.py
│   └── utils/
│       └── logger.py
└── raw_images/          # Example: root directory with images
```

---

## Usage

### GUI (Graphical User Interface)

You can launch the visual app for an interactive experience:

```bash
python app.py
```

**GUI Workflow:**
1. **Select Input Target:** Click **Browse** next to "Thư mục chứa ảnh" to select your root image folder.
2. **Scan Process:** Click **Bắt đầu quét** to start tracing. Real-time extraction logs and decoded results will populate the on-screen table. The results are initially kept in memory without generating loose files.
3. **Manual Export:** Once processing is complete, click the **Xuất Excel** button. A system file picker dialog will prompt you to choose the exact location and filename to save your `.xlsx` generated report.

---

### CLI Basic

```bash
python main.py --root ./raw_images
```

Produces a timestamped Excel file (e.g., `qr_results_20250313_143022.xlsx`) in the current directory.

---

### CLI All Options

```
python main.py [--root DIR] [--output FILE] [--verbose]

Required:
  --root DIR        Root directory containing subfolders of images.

Optional:
  --output FILE     Output Excel file path.
                    Default: qr_results_<YYYYMMDD_HHMMSS>.xlsx
  --verbose         Print detailed per-file progress to the console.
  -h, --help        Show this help message and exit.
```

---

### CLI Examples

```bash
# Standard run — timestamped output
python main.py --root ./raw_images

# Custom output path + verbose logging
python main.py --root ./raw_images --output results.xlsx --verbose

# Absolute paths
python main.py --root /data/devices --output /reports/qr_report.xlsx --verbose
```

**Verbose console output example:**
```
Processing folder: device_A
  Decoding: img1.jpg
  Decoding: img2.png
  Files: 2, Success: 1, Fail: 1

Processing folder: device_B
  Decoding: photo1.jpeg
  Files: 1, Success: 1, Fail: 0

Results written to: /path/to/results.xlsx
```

---

## Expected Input Layout

The script **auto-detects** which layout you are using — no configuration needed.

### ✅ Nested layout (one subfolder per device / group)

```
root/
├── device_A/
│   ├── img001.jpg
│   └── img002.png
└── device_B/
    ├── photo1.jpeg
    └── photo2.bmp
```

Each subfolder becomes its own row in the **Summary** table.

```bash
python main.py --root ./root
```

---

### ✅ Flat layout (all images in one folder, no subfolders)

```
raw_images/
├── img001.jpg
├── img002.png
└── photo1.jpeg
```

The root folder itself is treated as a single group.

```bash
python main.py --root ./raw_images
```

> The console will print:
> ```
> No subfolders found. Processing 3 image(s) directly in 'raw_images'.
> ```

---

**Supported image formats:** `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.tif`

Non-image files (PDFs, text files, etc.) are silently ignored in both layouts.

---

## Output Format

A single `.xlsx` file with one sheet named **`QR_Results`**.

### Table 1 — Summary (top of sheet)

| Folder | Date | Total Files | Success | Fail |
|---|---|---|---|---|
| device_A | 2025-03-10 14:30:22 | 5 | 4 | 1 |
| device_B | 2025-03-10 14:30:22 | 3 | 3 | 0 |

### *(blank separator row)*

### Table 2 — Details (below summary)

| folder | file_name | qr_content | status |
|---|---|---|---|
| device_A | img1.jpg | https://example.com/data | success |
| device_A | img2.png | | fail |
| device_B | photo1.jpeg | SN:12345; LOT:ABC | success |

> **Note on multiple QR codes per image:** all decoded strings are joined with `; ` in the `qr_content` column.

Column widths are automatically adjusted for readability.

---

## Behavior & Edge Cases

| Scenario | Behavior |
|---|---|
| Image has no QR code | `qr_content` = empty, `status` = `fail` |
| Image has multiple QR codes | All contents joined with `"; "` |
| Image file is corrupted | Logged as `fail`, processing continues |
| Subfolder has no images | Warning printed, subfolder skipped |
| `--root` does not exist | Error message printed, script exits with code 1 |
| Output file already exists | Interactive prompt: **[O]verwrite / [T]imestamped copy / [Q]uit** |

---

## Troubleshooting

### `ImportError: Unable to find zbar shared library`

The native `zbar` C library is not installed on your OS.

| OS | Fix |
|---|---|
| macOS | `brew install zbar` |
| Ubuntu/Debian | `sudo apt-get install libzbar0` |
| Windows | `choco install zbar` or place the DLL in PATH |

---

### `ModuleNotFoundError: No module named 'cv2'` (or similar)

Your virtual environment is not active, or dependencies were not installed.

```bash
source .venv/bin/activate      # macOS/Linux
# or
.venv\Scripts\activate         # Windows

pip install -r requirements.txt
```

---

### Download Error for WeChat Models

If the program complains about failing to retrieve the `WeChat QR models`:
- Ensure you have an active internet connection.
- During the first execution, it downloads about 4 small `.prototxt` & `.caffemodel` files (around 20MB in total). Once downloaded, they work perfectly offline!

---

### Low QR decode success rate

- The AI detector typically secures ~90% accuracy.
- Ensure images are **in focus**.
- Try increasing image resolution or lighting before running.

---

## Contributing

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-improvement`.
3. Commit your changes with clear messages.
4. Open a Pull Request — all PRs should include a description of what was changed and why.

---

## License

This project is licensed under the **MIT License** — see [`LICENSE`](LICENSE) for details.
