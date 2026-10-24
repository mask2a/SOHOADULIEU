"""Resolve editable configuration independently of the working directory."""
from pathlib import Path
import shutil
import sys


def resolve_data(saved=None):
    if saved and Path(saved).is_dir():
        return Path(saved).resolve()
    root = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent
    # Executables built into dist or dist/fixed can still find project data.
    roots = [root]
    if root.parent.name.casefold() == 'dist':
        roots.append(root.parent)
    if roots[-1].name.casefold() == 'dist':
        roots.append(roots[-1].parent)
    for candidate in roots:
        folder = candidate / 'data'
        if folder.is_dir():
            return folder.resolve()
    return root / 'data'


def resolve_alias(config_dir, saved=None):
    if saved and Path(saved).is_file():
        return Path(saved)
    resource_root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent.parent))
    app_root = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else resource_root
    external = app_root / 'config' / 'schools.csv'
    if external.is_file():
        return external
    destination = Path(config_dir) / 'schools.csv'
    if not destination.exists():
        destination.parent.mkdir(parents=True, exist_ok=True)
        template = resource_root / 'config' / 'schools.csv'
        if template.is_file():
            shutil.copyfile(template, destination)
        else:
            destination.write_text('alias,school_id,school_name,from_year,to_year\n', encoding='utf-8-sig')
    return destination
