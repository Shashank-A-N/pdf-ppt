# PPT to PDF Converter

A cross-platform desktop application for converting PowerPoint presentations (.ppt, .pptx) to PDF format offline.

## Features

- **Multiple Engine Support**: Uses LibreOffice or Microsoft PowerPoint (Windows only) for conversion
- **Batch Processing**: Convert multiple files or entire folders at once
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Simple GUI**: Easy-to-use interface with progress tracking
- **Recursive Folder Scanning**: Option to scan subdirectories for presentations
- **No Internet Required**: All processing happens locally on your machine

## Requirements

### For All Platforms
- Python 3.6 or higher

### Engine-Specific Requirements
- **LibreOffice** (recommended for cross-platform compatibility)
  - Windows: Install LibreOffice from [libreoffice.org](https://www.libreoffice.org/)
  - macOS: Install LibreOffice from [libreoffice.org](https://www.libreoffice.org/) or via Homebrew
  - Linux: Install via your distribution's package manager (typically `libreoffice` or `libreoffice-core`)

- **Microsoft PowerPoint** (Windows only)
  - Microsoft Office installed
  - Python `pywin32` package (`pip install pywin32`)

## Installation

1. **Install Python** if not already installed from [python.org](https://python.org)

2. **Install required packages**:
   ```bash
   pip install pywin32  # Windows only, for PowerPoint support
   ```

3. **Download the application**:
   - Save the `ppt-pdf.py` file to your preferred location

## Usage

1. **Run the application**:
   ```bash
   python ppt-pdf.py
   ```

2. **Add files/folders**:
   - Click "Add Files" to select individual PowerPoint files
   - Click "Add Folder" to select a directory containing presentations

3. **Configure options**:
   - **Engine**: Choose between Auto, LibreOffice, or PowerPoint (Windows only)
   - **Output Folder**: Specify where to save converted PDFs
   - **Recursive Scanning**: Enable to include subdirectories when selecting folders

4. **Start conversion**:
   - Click "Start Conversion" to begin processing
   - Monitor progress in the log area
   - Click "Cancel" to stop the current conversion

## Engine Selection

- **Auto**: Automatically selects the best available engine (PowerPoint on Windows if available, otherwise LibreOffice)
- **LibreOffice**: Uses LibreOffice's command-line interface (cross-platform)
- **PowerPoint**: Uses Microsoft PowerPoint COM automation (Windows only)

Use the "Detect Engines" button to check which conversion engines are available on your system.

## Output

Converted PDFs are saved to the specified output directory with the same base name as the original presentation. Existing files with the same name will be overwritten.

## Troubleshooting

### Common Issues

1. **"No conversion engine found"**
   - Install LibreOffice or ensure it's in your system PATH
   - On Windows, ensure PowerPoint is installed if selecting that option

2. **Conversion fails**
   - Check that the input files are valid PowerPoint presentations
   - Try using a different conversion engine

3. **Application won't start**
   - Ensure Python is properly installed
   - On Windows, if using PowerPoint engine, ensure pywin32 is installed: `pip install pywin32`

### Getting Help

If you encounter issues:
1. Check the log output in the application for specific error messages
2. Ensure your PowerPoint files are not corrupted or password-protected
3. Verify your selected conversion engine is properly installed

## License

This tool is provided as-is for personal use. Please ensure you have the right to convert any presentations you process with this tool.

## Contributing

This is a standalone script. For improvements or bug fixes, you can modify the source code directly.
