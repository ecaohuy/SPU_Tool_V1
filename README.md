# SPU Processing Tool

A GUI tool for processing CDD input files with SPU templates to generate network configuration output files.

## Features

- Modern, clean Tkinter GUI interface
- Load CDD input Excel files
- Apply SPU templates
- Generate processed output files
- Support for multiple sheet types (IP, Radio 2G/3G/4G/5G, RET, Mapping, etc.)

## Installation

### From Release (Windows)

1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file
3. Run `SPU_Tool.exe`

### From Source

```bash
# Clone the repository
git clone https://github.com/your-username/SPU_Tool_V1.git
cd SPU_Tool_V1

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py --tkinter
```

## Usage

1. Click **"Select CDD Input File"** to load your CDD Excel file
2. Click **"Select SPU Template"** to choose your template file
3. Click **"Process SPU Output"** to generate the output file

The output file will be saved in the `Output` folder.

## Folder Structure

```
SPU_Tool_V1/
├── main.py              # Entry point
├── config.json          # Configuration file
├── icon.png             # Application icon
├── Input/               # Place your CDD input files here
├── Template/            # Place your SPU template files here
├── Output/              # Generated output files
├── src/
│   ├── gui.py           # Tkinter GUI
│   ├── excel_handler.py # Excel file operations
│   ├── processor.py     # Data processing logic
│   ├── mapping_engine.py# Mapping rules
│   └── utils.py         # Utility functions
└── PARAMETER_GUIDE.txt  # Guide for modifying parameters
```

## Building Windows Executable

The project uses GitHub Actions to automatically build Windows executables.

### Manual Build

```bash
pip install pyinstaller
pyinstaller SPU_Tool.spec
```

The executable will be created in the `dist/SPU_Tool` folder.

## Configuration

See `PARAMETER_GUIDE.txt` for details on modifying tool parameters.

## License

MIT License
