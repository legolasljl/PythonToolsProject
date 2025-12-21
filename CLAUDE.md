# CLAUDE.md - AI Assistant Guide for PythonToolsProject

## Project Overview

This repository contains a collection of Python GUI desktop applications designed for **insurance document processing**. The tools focus on extracting, comparing, and analyzing insurance clause documents in Chinese, with support for English translation.

**Primary Use Case**: Insurance industry document workflow automation - extracting clauses from Word documents, comparing them against standard clause libraries, and generating Excel reports.

## Repository Structure

```
PythonToolsProject/
├── clause_diff_gui_ultimate.py    # Smart Clause Comparison Tool v12.0
├── claim_auto_filler_gui_v1_4.py  # (Empty/placeholder file)
├── word_extractor_gui.py          # Word Extractor Tool - Modern UI Edition
├── word_extractor_gui_final.py    # Word Extractor Tool - Final v5.0 (concurrent)
├── word_extractor_gui_v7_1.py     # Word Extractor Tool v7.1 (incremental mode)
├── make_icon.py                   # macOS .icns icon generator utility
├── icon.jpg                       # Source icon image
├── icon.icns                      # Generated macOS app icon
├── README.md                      # Basic project documentation
└── .gitignore                     # Python-specific ignores
```

## Core Applications

### 1. Smart Clause Comparison Tool (`clause_diff_gui_ultimate.py`)
**Version**: v12.0 (Smart Splitter Edition)

**Purpose**: Compare customer insurance clauses against a standard clause library and generate detailed matching reports.

**Key Features**:
- Smart document splitting (handles documents without blank line separators)
- English-to-Chinese translation using `deep_translator`
- Built-in insurance terminology glossary (财产险专业术语字典)
- Risk analysis and coverage gap detection
- Excel report generation with color-coded match scores

**Main Classes**:
- `ClauseMatcherLogic`: Core matching algorithm with NLP-based text similarity
- `MatchWorker`: QThread-based background processing
- `ClauseDiffGUI`: PyQt5 main window interface

### 2. Word Extractor Tools (`word_extractor_gui*.py`)
**Multiple Versions** with progressive enhancements:

| File | Version | Key Features |
|------|---------|--------------|
| `word_extractor_gui.py` | v3.0 | Basic extraction, cross-platform .doc/.docx support |
| `word_extractor_gui_final.py` | v5.0 | Multi-threaded processing, page noise filtering |
| `word_extractor_gui_v7_1.py` | v7.1 | Incremental mode, smart sheet categorization, row height auto-calculation |

**Purpose**: Batch extract insurance clause content from Word documents into structured Excel files.

**Key Features**:
- Cross-platform .doc to .docx conversion (Windows: COM, macOS: textutil)
- Automatic registration number extraction
- Header/footer noise filtering with regex patterns
- Category-based Excel sheet organization
- Horizontal (data analysis) and Vertical (print-friendly) output formats

### 3. Icon Generator (`make_icon.py`)
**Purpose**: Convert JPG/PNG images to macOS .icns format with proper Retina display sizes.

## Technology Stack

### Core Dependencies
```
PyQt5                  # GUI framework
python-docx            # Word document parsing (.docx)
openpyxl              # Excel file creation/styling
pandas                # Data manipulation
deep_translator       # Google Translate API wrapper (optional)
Pillow                # Image processing (make_icon.py)
```

### Platform-Specific
- **Windows**: `pywin32` for COM-based .doc conversion
- **macOS**: Native `textutil` for .doc conversion, `iconutil` for .icns

## Development Conventions

### Code Style
- **Encoding**: UTF-8 with explicit declaration (`# -*- coding: utf-8 -*-`)
- **Language**: Code comments and UI strings are primarily in Chinese (Simplified)
- **Author Attribution**: Files include author and date in docstrings

### Architecture Patterns

1. **Thread-based Processing**: All heavy operations use `QThread` with signals:
   ```python
   class WorkerThread(QThread):
       log_signal = pyqtSignal(str, str)      # (message, level)
       progress_signal = pyqtSignal(int, int) # (current, total)
       finished_signal = pyqtSignal(bool, str) # (success, result)
   ```

2. **Global Exception Handling**: Custom exception hook for user-friendly error dialogs:
   ```python
   sys.excepthook = global_exception_handler
   ```

3. **macOS Packaging Safety**: NullWriter pattern to prevent frozen app crashes:
   ```python
   if getattr(sys, 'frozen', False):
       sys.stdout = NullWriter()
       sys.stderr = NullWriter()
   ```

4. **Fusion UI Style**: Consistent cross-platform appearance using Qt Fusion style with custom light palette

### UI Design Patterns
- Card-based layouts with drop shadow effects
- Terminal-style log display (dark background, monospace font)
- Progress bars hidden until processing starts
- Gradient-styled action buttons
- High DPI scaling support

### File Processing Patterns
- Temporary file cleanup in `finally` blocks
- Regex-based noise filtering for headers/footers
- Smart text similarity using `difflib.SequenceMatcher`

## Building & Running

### Development Setup
```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # macOS/Linux
# or: venv\Scripts\activate  # Windows

# Install dependencies
pip install PyQt5 python-docx openpyxl pandas deep_translator Pillow

# Windows additional (for .doc support)
pip install pywin32
```

### Running Applications
```bash
# Run any of the GUI tools
python clause_diff_gui_ultimate.py
python word_extractor_gui_v7_1.py
```

### Creating macOS App Bundle
The tools are designed to be packaged with PyInstaller:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.icns clause_diff_gui_ultimate.py
```

## Key Glossary & Domain Knowledge

The clause comparison tool includes an extensive insurance terminology dictionary (`INSURANCE_GLOSSARY`) mapping English clause names to Chinese translations. Key terms include:

- **Deductible/Excess**: 免赔额
- **Reinstatement Value**: 重置价值
- **Removal of Debris**: 清理残骸费用
- **Waiver of Subrogation**: 放弃代位求偿权
- **Strike, Riot & Civil Commotion**: 罢工、暴乱及民众骚乱

## Important Notes for AI Assistants

1. **Bilingual Context**: This codebase serves Chinese-speaking insurance professionals. UI text, comments, and domain terminology are in Chinese.

2. **Version Evolution**: Multiple versions of word_extractor exist - prefer `v7_1` for new features or `final` for stability.

3. **Cross-Platform Considerations**:
   - .doc file conversion differs between Windows (COM) and macOS (textutil)
   - Test on both platforms when modifying conversion logic

4. **Excel Formatting**: The tools apply extensive styling (colors, borders, column widths). Preserve formatting logic when modifying output functions.

5. **Translation Dependency**: `deep_translator` is optional. Code gracefully degrades with `HAS_TRANSLATOR` flag.

6. **No Test Suite**: This project currently lacks automated tests. Manual testing with sample Word/Excel files is required.

## Common Modification Tasks

### Adding New Insurance Terms
Edit `ClauseMatcherLogic.INSURANCE_GLOSSARY` in `clause_diff_gui_ultimate.py`:
```python
INSURANCE_GLOSSARY = {
    # Add new entries here
    "your english term": "中文翻译",
}
```

### Modifying Noise Filters
Edit `NOISE_PATTERNS` in the word extractor files:
```python
NOISE_PATTERNS = [
    r'your_regex_pattern',
    # ...
]
```

### Changing Excel Output Format
Modify `save_to_excel()` methods - look for column width settings, header styling, and cell formatting.
