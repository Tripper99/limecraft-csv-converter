# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python desktop application that converts Limecraft transcription CSV files to Word documents (.docx) and Inqscribe format (.inqscr). The application features a modern GUI built with ttkbootstrap and handles timecode adjustments.

## Main Application File

- `limecraft_converter_v55 (1.3).py` - Main application file containing both the core conversion logic (`LimecraftConverter` class) and GUI (`LimecraftGUI` class)

## Development Commands

### Testing Code Syntax
```bash
python -m ruff check limecraft_converter_v55\ \(1.3\).py
```

### Running the Application
```bash
python limecraft_converter_v55\ \(1.3\).py
```

### Building Executable
```bash
pyinstaller "Limecraft csv converter (v1.3).spec"
```

## Dependencies

Required Python packages:
- `ttkbootstrap` - Modern Tkinter theme library for GUI
- `python-docx` - Word document creation (optional, app works without it)
- `Pillow` (PIL) - Image handling for the help window icon

Install dependencies:
```bash
pip install ttkbootstrap python-docx Pillow
```

## Architecture

### Core Classes

1. **LimecraftConverter**: Handles all CSV processing and file conversion logic
   - CSV parsing with multiple encoding support
   - Timecode parsing and adjustment
   - Word document generation
   - Inqscribe format export

2. **LimecraftGUI**: Manages the desktop GUI interface
   - Built with ttkbootstrap (superhero theme)
   - Scrollable interface with progress indicators
   - File selection dialogs and validation

### Key Features

- **Timecode Handling**: Supports multiple input formats (HH:MM:SS:FF, HH:MM:SS.FF, HHMMSSFF)
- **Multiple Output Formats**: Word documents and Inqscribe files
- **Time Adjustment**: Adds specified offset to all timecodes
- **Encoding Support**: Handles CSV files with UTF-8, UTF-8-sig, Latin-1, CP1252 encodings
- **Input Validation**: Comprehensive error handling and user feedback

### File Structure

- `build/` - PyInstaller build artifacts
- `dist/` - Compiled executables and distribution files
- `Old versions/` - Historical versions of the application
- `Csv-filer från Limecraft/` - Sample CSV files for testing
- `Output folder (testkörningar)/` - Test output files
- `*.spec` files - PyInstaller configuration files
- `Agg_med_smor_v4_transperant.png` - Help window image asset

## Constants and Configuration

- `FRAME_RATE = 30` - Video frame rate for timecode calculations
- `VERSION = "55"` - Application version number
- Application uses 30fps timecode format internally

## Error Handling

Custom exception hierarchy:
- `LimecraftConverterError` (base)
- `CSVValidationError`
- `TimecodeFormatError`
- `FileProcessingError`
- `ConversionError`

## Version Management

The application follows incremental version numbering (currently v55). Each version maintains the core filename pattern while incrementing the version number in constants and spec files.