#!/bin/bash
# Build script for Paper Formatter Android App

echo "====================================="
echo "Paper Formatter - Android APK Builder"
echo "====================================="

# Check Python version
python3 --version

# Install dependencies
echo "Installing Python dependencies..."
pip install -r requirements.txt

# Initialize buildozer if not done
if [ ! -f "buildozer.spec" ]; then
    echo "Initializing buildozer..."
    buildozer init
fi

# Build APK
echo "Building debug APK (first time takes 10-20 minutes)..."
buildozer android debug

echo "====================================="
echo "Build complete! APK is in bin/ folder"
echo "====================================="
