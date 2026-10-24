"""Disk-backed record store, candidate index and lineage adapter."""
import json
import sqlite3
from .domain import StudentRecord

class RecordStore:
    def __init__(self, path):
        self.db = sqlite3.connect(path)
        self.db.execute('''CREATE TABLE IF NOT EXISTS records (
          id INTEGER PRIMARY KEY, side TEXT, group_id TEXT, year TEXT, seq REAL,
          name TEXT, dob TEXT, grad TEXT, cert TEXT, valid INTEGER,
          payload TEXT, file TEXT, sheet TEXT, rownum INTEGER, origin TEXT, used INTEGER DEFAULT 0)''')
        self.db.execute('CREATE INDEX IF NOT EXISTS lookup ON records(side,name,dob,grad)')
        for col in ('cert', 'dob', 'grad'):
            self.db.execute(f'CREATE INDEX IF NOT EXISTS idx_{col} ON records(side,{col})')

    def add(self, record, side, group=''):
        name, dob, grad = record.keys
        try:
            seq = float(record.values[0])
        except ValueError:
            seq = 1e100
        self.db.execute('''INSERT INTO records(side,group_id,year,seq,name,dob,grad,cert,valid,payload,file,sheet,rownum,origin)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (side, group, grad, seq, name, dob, grad, record.certificate,
          int(record.valid), json.dumps(record.values, ensure_ascii=False), record.file, record.sheet, record.row, record.origin))

    @staticmethod
    def unpack(row):
        return row[0], StudentRecord(tuple(json.loads(row[10])), row[11], row[12], row[13], row[14])

    def query(self, condition='1', args=(), order='id'):
        for row in self.db.execute(f'SELECT * FROM records WHERE {condition} ORDER BY {order}', args):
            yield self.unpack(row)

    def exact(self, record, side):
        return list(self.query('side=? AND name=? AND dob=? AND grad=? AND valid=1', (side, *record.keys)))

    def candidates(self, record):
        n, d, y = record.keys
        # Empty values never count as agreement. Certificate is supporting evidence only.
        return list(self.query('''side='source' AND (
          (name=? AND name<>'' AND dob=? AND dob<>'') OR
          (name=? AND name<>'' AND grad=? AND grad<>'') OR
          (dob=? AND dob<>'' AND grad=? AND grad<>'') OR
          (cert=? AND cert<>''))''', (n,d,n,y,d,y,record.certificate)))

    def mark(self, ids):
        self.db.executemany('UPDATE records SET used=1 WHERE id=?', [(i,) for i in ids])

    def close(self):
        self.db.commit()
        self.db.close()
