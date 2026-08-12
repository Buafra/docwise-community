#!/bin/bash
# DocWise Community - macOS launcher.
# Double-click in Finder, or run from Terminal:  ./start.command
# First run installs Python 3.13, Tesseract OCR and app dependencies
# (via Homebrew if available). Later runs start instantly.

set -u
cd "$(dirname "$0")"
export TESSDATA_PREFIX="$PWD/tessdata"

die() {
    echo ""
    echo "SETUP FAILED: $1"
    echo "Fix the issue above, then run start.command again."
    read -r -p "Press Enter to close..." _
    exit 1
}

brew_exe() {
    if command -v brew >/dev/null 2>&1; then command -v brew; return 0; fi
    for b in /opt/homebrew/bin/brew /usr/local/bin/brew; do
        if [ -x "$b" ]; then echo "$b"; return 0; fi
    done
    return 1
}

# Find a Python >= 3.10. Skips the /usr/bin/python3 shim unless Xcode
# Command Line Tools are installed (running it would pop Apple's dialog).
find_python() {
    local c exe v
    local candidates=()
    candidates+=(python3.14 python3.13 python3.12 python3.11 python3.10)
    for c in /opt/homebrew/bin/python3* /usr/local/bin/python3* \
             /Library/Frameworks/Python.framework/Versions/3.*/bin/python3; do
        [ -x "$c" ] && candidates+=("$c")
    done
    if xcode-select -p >/dev/null 2>&1; then
        candidates+=(python3)
    fi
    for c in "${candidates[@]}"; do
        exe=$(command -v "$c" 2>/dev/null) || exe="$c"
        [ -x "$exe" ] || continue
        v=$("$exe" -c 'import sys; print("%d %d" % sys.version_info[:2])' 2>/dev/null) || continue
        set -- $v
        if [ "$1" -eq 3 ] && [ "$2" -ge 10 ]; then
            echo "$exe"
            return 0
        fi
    done
    return 1
}

echo ""
echo "[1/3] Checking Python..."
PYEXE=$(find_python) || PYEXE=""
if [ -z "$PYEXE" ]; then
    echo "  Python 3.10+ not found. Installing Python 3.13..."
    BREW=$(brew_exe) || BREW=""
    if [ -z "$BREW" ]; then
        echo ""
        echo "  Homebrew (the standard macOS package manager) is needed to"
        echo "  install Python and Tesseract automatically."
        read -r -p "  Install Homebrew now? [Y/n] " answer
        case "${answer:-Y}" in
            [Yy]*|"")
                /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" || die "Homebrew installation failed."
                BREW=$(brew_exe) || die "Homebrew installed but not found. Open a new Terminal and run start.command again."
                ;;
            *)
                die "Install Python 3.10+ yourself from https://www.python.org/downloads/ then run start.command again."
                ;;
        esac
    fi
    "$BREW" install python@3.13 || die "Homebrew could not install Python."
    export PATH="$("$BREW" --prefix)/bin:$PATH"
    PYEXE=$(find_python) || die "Python installed but not found. Open a new Terminal and run start.command again."
fi
echo "  OK: $("$PYEXE" --version) at $PYEXE"

echo ""
echo "[2/3] Checking app dependencies..."
PY=".venv/bin/python"
if [ ! -x "$PY" ]; then
    echo "  Creating virtual environment..."
    "$PYEXE" -m venv .venv || die "Could not create the virtual environment."
fi
if ! "$PY" -c 'import fastapi, uvicorn, onnxruntime, qrcode' >/dev/null 2>&1; then
    echo "  Installing dependencies (a few minutes on first run)..."
    "$PY" -m pip install --upgrade pip || die "pip upgrade failed. Check your internet connection."
    "$PY" -m pip install -r requirements.txt || die "pip install failed. Check your internet connection."
fi
echo "  OK: dependencies ready"

echo ""
echo "[3/3] Checking Tesseract OCR (for scanned documents and images)..."
BREW=$(brew_exe) || BREW=""
if [ -n "$BREW" ]; then
    export PATH="$(dirname "$BREW"):$PATH"
fi
if command -v tesseract >/dev/null 2>&1; then
    echo "  OK: Tesseract at $(command -v tesseract)"
elif [ -n "$BREW" ]; then
    echo "  Tesseract not found. Installing with Homebrew..."
    if "$BREW" install tesseract; then
        echo "  OK: Tesseract installed"
    else
        echo "  WARNING: Tesseract install failed. The app still works, but OCR of"
        echo "  scanned PDFs and images is disabled. Try later: brew install tesseract"
    fi
else
    echo "  WARNING: Tesseract not found and Homebrew unavailable. The app still"
    echo "  works, but OCR of scanned PDFs and images is disabled."
    echo "  To enable OCR later: install Homebrew, then: brew install tesseract"
fi

echo ""
echo "Starting DocWise Community..."
echo "Open: http://127.0.0.1:8120   (press Ctrl+C here to stop)"
echo ""
( sleep 2; open "http://127.0.0.1:8120" ) >/dev/null 2>&1 &
exec "$PY" app.py
