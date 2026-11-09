"""Deterministic smoke check usable from the frozen executable."""
import json
from pathlib import Path
import tkinter as tk
from openpyxl import Workbook, load_workbook
from .domain import HEADERS
from .presentation import App
from .application import compare

def run(folder):
    folder=Path(folder).resolve(); folder.mkdir(parents=True,exist_ok=True)
    root=tk.Tk(); root.withdraw()
    app=App(root,config_dir=folder/'config')
    assert Path(app.alias.get()).is_file(), 'Missing schools CSV in frozen app'
    root.update(); root.destroy()
    record=['']*24
    record[4]='HỌC SINH MẪU'; record[6]='01/02/1980'; record[12]='1995'
    for name in ('source.xlsx','school.xlsx'):
        wb=Workbook(); ws=wb.active; ws.title='Data'; ws.append(HEADERS); ws.append(record)
        wb.save(folder/name); wb.close()
    output=compare([folder/'source.xlsx'],folder/'school.xlsx',folder/'output',years={1995})
    wb=load_workbook(output/'source_Check.xlsx')
    assert wb.sheetnames==[
        'TONGHOP','DULIEU_OK','DULIEU_CHI_CO_TRONG_SGD',
        'DULIEU_KHONGCO_TRONG_SGD','DULIEU_LECH_THONG_TIN',
        'DULIEU_CAN_XAC_MINH',
    ]
    assert wb['DULIEU_OK']['A2'].value=='Khớp hoàn toàn'
    wb.close()
    imported=load_workbook(output/'import'/'import_1995.xlsx',read_only=True)
    assert imported['Data'].max_row==2
    imported.close()
    (folder/'result.json').write_text(json.dumps({'tkinter':'ok','excel_sqlite_report':'ok'},indent=2),encoding='utf-8')
