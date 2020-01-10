import json
import struct
import pyodbc

from os.path import join
from os import name


def handle_datetimeoffset(dto_value):
    # ref: https://github.com/mkleehammer/pyodbc/issues/134#issuecomment-281739794
    tup = struct.unpack("<6hI2h", dto_value)  # e.g., (2017, 3, 16, 10, 35, 18, 0, -6, 0)
    tweaked = [tup[i] // 100 if i == 6 else tup[i] for i in range(len(tup))]
    return "{:04d}-{:02d}-{:02d} {:02d}:{:02d}:{:02d}.{:07d} {:+03d}:{:02d}".format(*tweaked)


def connect(data_folder=False):
    if data_folder:
        path = join('..', 'config', 'config.json')
    else:
        path = join('config', 'config.json')
    with open(path, encoding='utf8') as fin:
        data = json.load(fin)
        dsn = data['dsn']
        uid = data['uid']
        pwd = data['pwd']
        base = data['base']
        driver = data['driver']
        server = data['server']
        database = data['database']

    if name == 'nt':
        connect_db = pyodbc.connect(f'DRIVER={driver};SERVER={server};'
                                    f'DATABASE={database};UID={uid};PWD={pwd}')
    else:
        connect_db = pyodbc.connect(f'DSN={dsn};UID={uid};PWD={pwd}')
    connect_db.add_output_converter(-155, handle_datetimeoffset)

    return {'base': base,
            'connect_db': connect_db}
