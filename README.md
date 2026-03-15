# Paper Formatter Android App

A Kivy-based Android app that converts Markdown/Word documents to Chinese thesis-formatted Word documents.

## Setup on Your Machine

```bash
# 1. Install Python dependencies
pip install kivy buildozer python-docx markdown

# 2. Initialize buildozer
buildozer init

# 3. Build APK (first time takes ~10-20 min)
buildozer android debug
```

The APK will be in `bin/` directory.

## Project Structure

```
paper-formatter-android/
├── main.py              # Main Kivy app
├── buildozer.spec       # Buildozer configuration
├── requirements.txt     # Python dependencies
├── src/
│   └── app/
│       ├── __init__.py
│       ├── main.kv      # UI layout
│       └── converter.py # Document conversion logic
└── templates/
    └── china_master.json
```
