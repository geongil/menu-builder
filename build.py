#!/usr/bin/env python3
"""
Build script for standalone executable.
Output: dist/menu-builder.exe (Windows) or dist/menu-builder.app (macOS).
"""
import platform
import sys

def main():
    try:
        import PyInstaller.__main__
    except ImportError:
        print("PyInstaller required. Install with:")
        print("  pip install -r requirements-build.txt")
        sys.exit(1)

    name = "menu-builder"
    args = [
        "main.py",
        "--name", name,
        "--onefile",
        "--windowed",
        "--clean",
    ]

    PyInstaller.__main__.run(args)
    if platform.system() == "Windows":
        print("\nBuild OK. Output: dist/*.exe")
    elif platform.system() == "Darwin":
        print("\nBuild OK. Output: dist/*.app")
    else:
        print("\nBuild OK. Output: dist/")

if __name__ == "__main__":
    main()
