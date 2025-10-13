# Limecraft CSV Converter

A Python desktop application that converts Limecraft transcription CSV files to Word documents (.docx) and Inqscribe format (.inqscr), with support for timecode adjustments and speaker identification.

## Features

- **Multiple CSV Format Support**: Handles both standard CSV format and Limecraft's combined format
- **Timecode Adjustment**: Add offset time to all timecodes (e.g., camera start time)
- **Multiple Output Formats**:
  - Word documents (.docx) with proper formatting
  - Inqscribe files (.inqscr) for video editing
- **Speaker Integration**: Includes speaker names in output format: `Speaker: transcript text`
- **Modern GUI**: Built with ttkbootstrap for a modern interface
- **Flexible Timecode Input**: Supports multiple input formats (HH:MM:SS:FF, HH:MM:SS.FF, HHMMSSFF)

## Requirements

### Required Dependencies
- Python 3.8+
- `ttkbootstrap` - Modern GUI framework
- `python-docx` - Word document creation (optional, app works without it)
- `Pillow` (PIL) - Image handling

### Installation
```bash
pip install ttkbootstrap python-docx Pillow
```

## Usage

### From Limecraft Export
1. In Limecraft, select "Export" → "CSV file"
2. Ensure "Media Start", "Transcript", and "Speakers" columns are checked
3. Export the CSV file

### In the Application
1. **Select CSV File**: Browse and select your Limecraft CSV file
2. **Adjust Start Time** (optional): Add time offset to all timecodes
3. **Set Output Filename**: Choose custom filename or use CSV filename
4. **Select Output Format**: Choose Word document and/or Inqscribe file
5. **Optional Filename Prefix**: Include filename in timecode display
6. **Convert**: Choose output directory and convert files

### Supported Timecode Formats
- `HH:MM:SS:FF` (e.g., 01:30:45:12)
- `HH:MM:SS.FF` (e.g., 01:30:45.12)
- `HH.MM.SS.FF` (e.g., 01.30.45.12)
- `HHMMSSFF` (e.g., 01304512)

## Output Format Example

```
[00:03:15.08]
Maria: Men det är dåligt skrivet av mig. Varför ens gjorde ni scenen?

[00:03:33.10]
DJ: Jag förstår precis vad du menar. För det är ju det här med att...
```

## File Structure

- `limecraft_converter_v56 (1.5).py` - Main application
- `Limecraft csv converter (v1.5).spec` - PyInstaller build configuration
- `CLAUDE.md` - Development guidelines
- `Agg_med_smor_v4_transperant.png` - Help window image

## Building Executable

```bash
pyinstaller "Limecraft csv converter (v1.5).spec"
```

## Technical Details

### CSV Format Support

The application handles two CSV formats:

1. **Standard Format**: Separate columns for Media Start, Speakers, Transcript
2. **Limecraft Combined Format**: All data in quoted single field: `"timecode,speaker,transcript"`

### Architecture

- **LimecraftConverter**: Core conversion logic and CSV parsing
- **LimecraftGUI**: Modern desktop interface with ttkbootstrap
- **Error Handling**: Comprehensive exception handling with user feedback

### Frame Rate
- Uses 30fps timecode format internally
- Converts between timecode formats automatically

## Version History

- **v1.5**: Added support for Limecraft's combined CSV format, improved speaker integration
- **v1.4**: Added speaker name support in output format
- **v1.3**: Enhanced GUI and timecode handling
- **v1.2**: Improved CSV parsing and validation
- **v1.1**: Added Inqscribe export functionality
- **v1.0**: Initial release with Word export

## Author

Dan Josefsson, 2025

## License

This project is licensed under the MIT License - see the LICENSE file for details.