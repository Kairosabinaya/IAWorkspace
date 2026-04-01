"""Nuitka build script for IA Workspace."""

import os
import subprocess
import sys


def build():
    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--assume-yes-for-downloads",
        "--enable-plugin=tk-inter",
        '--windows-company-name=Internal Audit - Agung Sedayu Group',
        '--windows-product-name=IA Workspace',
        "--windows-file-version=1.0.0",
        "--output-filename=IAWorkspace.exe",
        "main.py",
    ]

    # Add icon only if it exists
    if os.path.exists("assets/icon.ico"):
        cmd.insert(5, "--windows-icon-from-ico=assets/icon.ico")

    print("Building with Nuitka...")
    print(" ".join(cmd))
    subprocess.run(cmd, check=True)
    print("Build complete!")


if __name__ == "__main__":
    build()
