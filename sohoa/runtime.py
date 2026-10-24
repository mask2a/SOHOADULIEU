"""Support Python distributions shipping Tcl/Tk scripts inside zip archives."""
import os
from pathlib import Path
import sys
import zipfile

def prepare_tk():
    if getattr(sys, 'frozen', False):
        return  # PyInstaller's Tcl/Tk runtime hook selects bundled data.
    base = Path(sys.base_prefix) / 'tcl'
    target = Path(__file__).resolve().parent.parent / '.runtime'
    for pattern, folder, variable in [('libtcl*.zip','tcl_library','TCL_LIBRARY'),('libtk*.zip','tk_library','TK_LIBRARY')]:
        if variable in os.environ:
            continue
        archives = sorted(base.glob(pattern))
        if archives:
            dest = target / folder
            if not dest.exists():
                target.mkdir(exist_ok=True)
                with zipfile.ZipFile(archives[-1]) as archive:
                    for member in archive.infolist():
                        resolved = (target / member.filename).resolve()
                        if not resolved.is_relative_to(target.resolve()):
                            raise RuntimeError('Unsafe Tcl archive entry')
                    archive.extractall(target)
            os.environ[variable] = str(dest)
