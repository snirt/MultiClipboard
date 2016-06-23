import sys
from cx_Freeze import setup, Executable

target = Executable(
    script="MultiClipboard.py",
    base="Win32GUI",
    compress=False,
    copyDependentFiles=True,
    appendScriptToExe=True,
    appendScriptToLibrary=False,
    icon="MultiClipboard.ico"
    )

setup(
    name="MultiClipboard",
    version="0.7",
    description="MultiClipboard - Clipboard tool",
    author="Snir Turgeman",
    # options={"build_exe": options},
    executables=[target]
    )