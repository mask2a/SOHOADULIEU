from pathlib import Path
import importlib.util
import subprocess
import sys


def main():
    if not getattr(sys, 'frozen', False):
        missing = [name for name in ('openpyxl', 'pandas') if importlib.util.find_spec(name) is None]
        if missing:
            root = Path(__file__).resolve().parent
            venv_python = root / '.venv' / ('Scripts/python.exe' if sys.platform == 'win32' else 'bin/python')
            if venv_python.is_file() and Path(sys.executable).resolve() != venv_python.resolve():
                return subprocess.call([str(venv_python), str(root / 'main.py'), *sys.argv[1:]])
            print('Thiếu thư viện: ' + ', '.join(missing) + '\nCài bằng: python -m pip install -r "' + str(root / 'requirements.txt') + '"', file=sys.stderr)
            return 1
    from sohoa.runtime import prepare_tk
    prepare_tk()
    if len(sys.argv)==3 and sys.argv[1]=='--self-test':
        from sohoa.smoke import run
        run(sys.argv[2])
    else:
        from sohoa.presentation import launch
        launch()
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
