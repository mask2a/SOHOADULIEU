import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from sohoa.runtime import prepare_tk


class ConfigurationPathTests(unittest.TestCase):
    def test_exe_in_fixed_finds_project_data(self):
        from sohoa.paths import resolve_data
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory); (root/'data').mkdir()
            exe=root/'dist'/'fixed'/'THCS.exe'
            with patch('sys.frozen',True,create=True), patch('sys.executable',str(exe)):
                self.assertEqual(resolve_data(root/'dist'/'data'),root/'data')

    def test_existing_custom_data_is_preserved(self):
        from sohoa.paths import resolve_data
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(resolve_data(directory),Path(directory).resolve())

    def test_standalone_exe_seeds_editable_template_and_preserves_edits(self):
        from sohoa.paths import resolve_alias
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory); bundle=root/'bundle'; (bundle/'config').mkdir(parents=True)
            template=bundle/'config'/'schools.csv'
            template.write_text('alias,school_id,school_name,from_year,to_year\n',encoding='utf-8')
            exe=root/'portable'/'THCS.exe'; config=root/'user'
            with patch('sys.frozen',True,create=True), patch('sys._MEIPASS',str(bundle),create=True), patch('sys.executable',str(exe)):
                path=resolve_alias(config,root/'missing.csv')
                self.assertEqual(path,config/'schools.csv')
                self.assertEqual(path.read_text(encoding='utf-8'),template.read_text(encoding='utf-8'))
                path.write_text('USER EDIT',encoding='utf-8')
                self.assertEqual(resolve_alias(config).read_text(encoding='utf-8'),'USER EDIT')

    def test_valid_custom_selection_is_preserved(self):
        from sohoa.paths import resolve_alias
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'custom.csv'; path.write_text('custom',encoding='utf-8')
            self.assertEqual(resolve_alias(Path(directory)/'settings',path),path)

    def test_gui_started_from_dist_resolves_missing_saved_alias(self):
        prepare_tk()
        import tkinter as tk
        import json
        from sohoa.presentation import App
        previous = Path.cwd()
        with tempfile.TemporaryDirectory(dir=previous) as directory:
            base = Path(directory).resolve()
            dist = base / 'dist'; dist.mkdir()
            config = base / 'settings'; config.mkdir()
            missing = dist / 'config' / 'schools.csv'
            (config / 'settings.json').write_text(json.dumps({'alias': str(missing), 'data': str(dist/'data')}), encoding='utf-8')
            root = tk.Tk(); root.withdraw()
            try:
                os.chdir(dist)
                app = App(root, config_dir=config)
                self.assertTrue(Path(app.data.get()).is_dir(), f'Missing data folder: {app.data.get()}')
                alias = Path(app.alias.get())
                self.assertTrue(alias.is_file(), f'GUI selects nonexistent CSV: {alias}')
                self.assertIn('school_id', alias.read_text(encoding='utf-8-sig'))
            finally:
                os.chdir(previous)
                root.destroy()
