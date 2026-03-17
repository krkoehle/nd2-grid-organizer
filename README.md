# nd2grid

Organize Nikon `.nd2` confocal microscopy images into PowerPoint grids on black backgrounds with white labels. Fluorescent channels are separated by color; phase contrast / brightfield channels are grouped together.

## Install

Requires Python 3.9+. Install with [pipx](https://pipx.pypa.io/) (recommended) or pip:

```bash
# One-command install (recommended)
pipx install git+https://github.com/koehlerlab/nd2-grid-organizer.git

# Or with pip
pip install git+https://github.com/koehlerlab/nd2-grid-organizer.git
```

If you don't have pipx yet:
```bash
# macOS
brew install pipx && pipx ensurepath

# Windows
pip install pipx && pipx ensurepath

# Linux
pip install pipx && pipx ensurepath
```

## Usage

```bash
# Process all .nd2 files in a folder
nd2grid /path/to/nd2/folder

# Process specific files
nd2grid file1.nd2 file2.nd2 file3.nd2

# Customize output filename and grid columns
nd2grid /path/to/folder -o my_figures.pptx --cols 3
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `-o`, `--output` | `microscopy_grid.pptx` | Output PowerPoint file path |
| `-c`, `--cols` | `4` | Number of columns in the image grid |

## What it does

1. Reads `.nd2` files and extracts each channel
2. Auto-detects phase contrast / brightfield / DIC vs fluorescent channels
3. Applies percentile-based contrast normalization
4. Colors fluorescent channels using their native Nikon color assignments
5. Creates widescreen (16:9) PowerPoint slides:
   - One slide for all phase/brightfield images
   - One slide per fluorescent channel (e.g., all GFP on one slide, all Cy5 on another)
6. Handles Z-stacks (max intensity projection), multi-position, and time series data

## Supported channel detection

Phase/brightfield keywords (auto-detected): `Phase`, `Brightfield`, `BF`, `DIC`, `Transmitted`, `Trans`, `TD`, `Diascopic`

All other channels are treated as fluorescent and colored according to their Nikon metadata.
