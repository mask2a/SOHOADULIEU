"""Run from repository root with .venv Python."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from sohoa.runtime import prepare_tk
prepare_tk()
import PyInstaller.__main__

if __name__ == '__main__':
    PyInstaller.__main__.run([
        '--noconfirm', '--clean', '--onefile', '--windowed',
        '--name', 'THCS_Offline', '--exclude-module', 'pandas.tests',
        '--add-data', str(ROOT / 'config' / 'schools.csv') + ':config',
        str(ROOT / 'main.py'),
        *sys.argv[1:],
    ])
