# 3MF Converter for Orca Slicer

Having trouble opening 3MF files in Orca Slicer? Models loading without colors, or not loading at all? This tool fixes that.

## Quick Start

1. Download **`3MF Converter.exe`** from the [Releases](../../releases) page
2. Double-click to run (no install needed)
3. Click the **+** area to select your 3MF files
4. Optionally pick an output folder (defaults to same folder as input)
5. Click **Convert All**
6. Open the new `*_orca.3mf` file in Orca Slicer

That's it! Your converted files will be saved next to the originals with `_orca` added to the name.

## What does this fix?

3MF files come in different formats depending on what software created them. Current versions of Orca Slicer expect a specific structure (separate object files, production metadata, etc.). Files from older BambuStudio versions, other modeling tools, or downloaded models often don't match this structure.

This tool automatically converts any 3MF into the format Orca expects:

- **Old BambuStudio files** — restructures inline mesh data and adds required metadata
- **Bare 3MF files** (from Blender, 3D Builder, modeling tools, etc.) — wraps them in the full Orca-compatible structure
- **Already-compatible files** — detected and skipped automatically

All paint colors, support painting, and material assignments are preserved during conversion.

## Features

- Dark-themed GUI — no command line needed
- Batch convert multiple files at once
- Progress bar with per-file status
- Debug log panel to see exactly what's happening
- Handles files of any size (tested with 95MB+ 3MF files)
- Auto-skips files that are already in the right format

## Command Line Usage

For power users, the converter also works from the command line:

```bash
# Convert a single file
python convert_3mf.py "My Model.3mf"

# Convert with specific output
python convert_3mf.py input.3mf -o output.3mf

# Batch convert everything in a folder
python convert_3mf.py *.3mf

# Convert in place (overwrites original)
python convert_3mf.py input.3mf --in-place
```

## Building from source

If you want to run from source or build the exe yourself:

```bash
# Run directly (Python 3.8+, no dependencies needed)
python gui.py

# Build standalone exe
pip install pyinstaller
pyinstaller --onefile --windowed --name "3MF Converter" gui.py
# Output: dist/3MF Converter.exe
```

## FAQ

**Q: Will this mess up my original files?**
A: No. The converter creates a new file with `_orca` appended to the name. Your original is never modified.

**Q: My file was "skipped" — is that bad?**
A: Nope! That means your file is already in the correct format for Orca. No conversion needed.

**Q: The conversion takes a while on big files — is it stuck?**
A: Large 3MF files (50MB+) can take 15-30 seconds. Watch the progress bar and log panel — if it's working, you'll see activity there.

**Q: My model loads but has no colors after converting.**
A: That means the original file didn't have color data to begin with. This tool preserves all existing colors, but it can't add colors that aren't there. You'll need to paint the model in Orca Slicer.

---

Made by [PrintShack3D](https://github.com/ajs1616) to help the maker community.
