"""Nuitka build script for IA Workspace."""

import subprocess
import sys


def build():
    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--enable-plugin=tk-inter",
        "--windows-icon-from-ico=assets/icon.ico",
        '--windows-company-name=Internal Audit - Agung Sedayu Group',
        '--windows-product-name=IA Workspace',
        "--windows-file-version=1.0.0",
        "--output-filename=IAWorkspace.exe",
        "main.py",
    ]
    print("Building with Nuitka...")
    print(" ".join(cmd))
    subprocess.run(cmd, check=True)
    print("Build complete!")


if __name__ == "__main__":
    build()
