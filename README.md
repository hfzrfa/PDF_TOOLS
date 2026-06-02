# PDF Tools

PDF Tools is a Windows-friendly desktop utility for common PDF workflows. It provides PDF compression, document conversion, table extraction, and image-to-PDF merging in a single CustomTkinter interface.

## Features

- Compress PDF files with multiple compression levels.
- Convert PDF files to Word documents (`.docx`).
- Convert Word documents (`.doc`, `.docx`) to PDF.
- Convert Excel workbooks (`.xls`, `.xlsx`, `.xlsm`) to PDF.
- Extract tables or text from PDF files into Excel workbooks (`.xlsx`).
- Merge multiple image files into one PDF.
- Drag-and-drop file input when `tkinterdnd2` is available.
- Remembers the last folder used for file selection and output.

## Supported Inputs

| Tool | Input | Output |
| --- | --- | --- |
| Compress | `.pdf` | `_compressed.pdf` |
| PDF to Word | `.pdf` | `.docx` |
| Word to PDF | `.doc`, `.docx` | `.pdf` |
| Excel to PDF | `.xls`, `.xlsx`, `.xlsm` | `.pdf` |
| PDF to Excel | `.pdf` | `.xlsx` |
| Images to PDF | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tif`, `.tiff`, `.webp` | `.pdf` |

## Requirements

- Python 3.9 or newer
- Windows is recommended for Word and Excel export support
- Microsoft Word and Microsoft Excel for the best Word/Excel to PDF conversion
- LibreOffice can be used as a fallback for Word/Excel to PDF when available

Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Running the App

```bash
python pdf_compressor.py
```

On Windows, you can also use:

```bat
install_and_run.bat
```

## Building the Executable

The project includes a PyInstaller spec file.

```bash
pip install pyinstaller
python -m PyInstaller "PDF Compressor.spec" --noconfirm
```

The executable will be created in:

```text
dist/PDF Compressor.exe
```

## Notes and Limitations

- PDF compression is best-effort and depends on the structure of the source PDF.
- PDF to Word conversion may not perfectly preserve complex layouts.
- PDF to Excel prioritizes detected tables; if no table is found, it exports page text into the worksheet.
- Word to PDF and Excel to PDF require Microsoft Office COM automation on Windows, or LibreOffice as a fallback.
- Generated build artifacts are intentionally ignored by Git.

## Project Structure

```text
.
|-- pdf_compressor.py       # Main application
|-- requirements.txt        # Python dependencies
|-- PDF Compressor.spec     # PyInstaller build configuration
|-- install_and_run.bat     # Windows setup and launch helper
`-- README.md
```

## Author

Created by `hfzrfa`.
