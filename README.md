# 3MF Converter for Orca Slicer

Converts old-format BambuStudio 3MF files to the current Orca Slicer format so you don't have to manually recolor your models every time you open them.

## The Problem

3MF files created with older versions of BambuStudio store mesh data inline in a single model file. Current versions of Orca Slicer expect mesh data split into separate object files with production extension metadata (UUIDs, paths). When you open an old-format file in Orca, it loads but loses all color/material assignments, forcing you to repaint everything by hand.

## Download

Grab the latest `3MF Converter.exe` from the [Releases](../../releases) page. No Python or dependencies needed — just download and double-click.

## Usage

### GUI (recommended)
1. Run `3MF Converter.exe` (or `python gui.py`)
2. Click the drop zone to browse for 3MF files (or drag and drop)
3. Click **Convert All**
4. Converted files are saved next to the originals as `*_orca.3mf`

Files already in the new format are automatically detected and skipped.

### Command Line
```bash
# Convert a single file
python convert_3mf.py "My Model.3mf"

# Convert with specific output
python convert_3mf.py input.3mf -o output.3mf

# Batch convert
python convert_3mf.py *.3mf

# Convert in place (overwrites original)
python convert_3mf.py input.3mf --in-place
```

## What It Does

- Extracts inline mesh data into separate `3D/Objects/` files
- Adds production extension (`p:UUID`, `p:path`) that current Orca requires
- Creates `3D/_rels/3dmodel.model.rels` linking object files
- Preserves all `paint_color` and `paint_supports` data (your color assignments)
- Auto-detects already-converted files and skips them

## Building the exe yourself

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "3MF Converter" gui.py
```

The exe will be in the `dist/` folder.

## Requirements (for running from source)

- Python 3.8+
- No external dependencies (uses only stdlib + tkinter)
- Optional: `tkinterdnd2` for drag-and-drop support in the GUI
