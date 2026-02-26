#!/usr/bin/env python3
"""
실행 파일 빌드 스크립트 (개발자용).
빌드 후 사용자에게 dist 안의 실행 파일만 전달하면 됨.
사용자는 커맨드 없이 더블클릭만 하면 앱이 실행됨.
- Windows: dist/한달식단.exe
- macOS: dist/한달식단.app
"""
import platform
import subprocess
import sys

def main():
    try:
        import PyInstaller.__main__
    except ImportError:
        print("PyInstaller가 필요합니다. 먼저 설치하세요:")
        print("  pip install -r requirements-build.txt")
        sys.exit(1)

    name = "한달식단"
    args = [
        "main.py",
        "--name", name,
        "--onefile",
        "--windowed",   # Windows: 콘솔 창 없음. macOS: .app 번들로 더블클릭 시 터미널 안 뜸
        "--clean",
    ]

    PyInstaller.__main__.run(args)
    if platform.system() == "Windows":
        print(f"\n빌드 완료. 사용자는 dist/{name}.exe 더블클릭만 하면 됩니다.")
    elif platform.system() == "Darwin":
        print(f"\n빌드 완료. 사용자는 dist/{name}.app 더블클릭만 하면 됩니다.")
    else:
        print(f"\n빌드 완료. 실행 파일: dist/{name}")

if __name__ == "__main__":
    main()
