# nd2grid

Organize Nikon `.nd2` confocal microscopy images into PowerPoint grids on black backgrounds with white labels. Fluorescent channels are separated by color; phase contrast / brightfield channels are grouped together.

---

## Quick-Start Guide (No Coding Experience Needed)

You only need to do **Step 1** once per computer. After that, skip straight to **Step 2** every time.

### Step 1: One-Time Setup

You'll be copying and pasting commands into **Terminal** (Mac/Linux) or **Command Prompt** (Windows). Don't worry — just follow along exactly.

#### 1a. Check if Python is installed

Open Terminal (Mac: search "Terminal" in Spotlight) or Command Prompt (Windows: search "cmd") and type:

```
python3 --version
```

If you see something like `Python 3.9.6`, you're good — skip to 1b. If you get an error:
- **Mac**: Install from [python.org/downloads](https://www.python.org/downloads/). Download, open the installer, click through the prompts.
- **Windows**: Install from [python.org/downloads](https://www.python.org/downloads/). **Important**: check the box that says "Add Python to PATH" during install.

#### 1b. Install pipx (a tool that safely installs command-line apps)

Copy and paste this into Terminal / Command Prompt and press Enter:

**Mac / Linux:**
```
python3 -m pip install --user pipx && python3 -m pipx ensurepath
```

**Windows:**
```
pip install pipx && pipx ensurepath
```

Close and reopen your Terminal / Command Prompt after this step.

#### 1c. Install nd2grid

Copy and paste this and press Enter:

```
pipx install git+https://github.com/krkoehle/nd2-grid-organizer.git
```

That's it! The `nd2grid` command is now available on your computer.

---

### Step 2: Using nd2grid

#### The basics

Put all your `.nd2` files in one folder, then run:

```
nd2grid /path/to/your/folder
```

This creates a file called `microscopy_grid.pptx` in your current directory.

#### Tips for finding your folder path

- **Mac**: Right-click the folder in Finder → "Get Info" → copy the path after "Where:", then add the folder name. Example: `/Users/yourname/Desktop/my_images`
- **Windows**: Click the address bar in File Explorer and copy the path. Example: `C:\Users\yourname\Desktop\my_images`
- **Drag and drop** (Mac): Type `nd2grid ` (with a space), then drag your folder from Finder into the Terminal window. It will paste the path for you!

#### Choosing an output name

```
nd2grid /path/to/your/folder -o my_figure.pptx
```

#### Changing the number of columns

By default you get 4 columns. For 3 columns:

```
nd2grid /path/to/your/folder --cols 3
```

#### Combining options

```
nd2grid /path/to/your/folder -o experiment_results.pptx --cols 3
```

---

### Updating to the latest version

If the tool gets updated, run:

```
pipx upgrade nd2grid
```

---

### Troubleshooting

| Problem | Solution |
|---------|----------|
| `command not found: nd2grid` | Close and reopen Terminal, then try again. If still broken, re-run Step 1b. |
| `command not found: python3` | Install Python (see Step 1a). |
| `command not found: pipx` | Close and reopen Terminal. If still broken, re-run Step 1b. |
| A specific `.nd2` file gives an error | The file may be corrupted or in an older format. Try other files first. |

---

## How It Works (For the Curious)

1. Reads `.nd2` files and extracts each channel
2. Auto-detects phase contrast / brightfield / DIC vs fluorescent channels
3. Applies contrast normalization so images look good
4. Colors fluorescent channels using their Nikon color assignments (green for GFP, red for DsRed, etc.)
5. Creates widescreen PowerPoint slides:
   - One slide for all phase/brightfield images
   - One slide per fluorescent channel (e.g., all GFP on one slide, all Cy5 on another)
6. Handles Z-stacks (max intensity projection), multi-position, and time series data

### Recognized phase/brightfield channel names

`Phase`, `Brightfield`, `BF`, `DIC`, `Transmitted`, `Trans`, `TD`, `Diascopic`

All other channels are treated as fluorescent.

---

## Developer Install

```bash
pipx install git+https://github.com/krkoehle/nd2-grid-organizer.git

# Or with pip directly
pip install git+https://github.com/krkoehle/nd2-grid-organizer.git
```

### Options reference

| Flag | Default | Description |
|------|---------|-------------|
| `-o`, `--output` | `microscopy_grid.pptx` | Output PowerPoint file path |
| `-c`, `--cols` | `4` | Number of columns in the image grid |
