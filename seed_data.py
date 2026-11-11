from pathlib import Path
import sqlite3, re, hashlib
from datetime import datetime
from openpyxl import load_workbook
BASE=Path(__file__).resolve().parent; DB=BASE/'wms.db'
with sqlite3.connect(DB) as c:
 c.executescript('''CREATE TABLE IF NOT EXISTS employees(id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,role TEXT,category TEXT,active INTEGER DEFAULT 1);
 CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,unit_weight_kg REAL,units_per_pack INTEGER,location_code TEXT,zone TEXT,active INTEGER DEFAULT 1);
 CREATE TABLE IF NOT EXISTS routes(id INTEGER PRIMARY KEY,route_date TEXT NOT NULL,name TEXT NOT NULL,status TEXT DEFAULT 'Ждёт',picker_id INTEGER,trays_out INTEGER DEFAULT 0,driver TEXT,source_hash TEXT,updated_at TEXT,UNIQUE(route_date,name));
 CREATE TABLE IF NOT EXISTS route_items(id INTEGER PRIMARY KEY,route_id INTEGER NOT NULL,product_id INTEGER,product_name TEXT NOT NULL,quantity REAL NOT NULL DEFAULT 0,pack_note TEXT,picked INTEGER DEFAULT 0);
 CREATE TABLE IF NOT EXISTS penalties(id INTEGER PRIMARY KEY,employee_id INTEGER NOT NULL,work_date TEXT NOT NULL,amount REAL NOT NULL,note TEXT);
 CREATE TABLE IF NOT EXISTS shifts(id INTEGER PRIMARY KEY,employee_id INTEGER NOT NULL,work_date TEXT NOT NULL,present INTEGER DEFAULT 1,UNIQUE(employee_id,work_date));
 CREATE TABLE IF NOT EXISTS tray_ledger(id INTEGER PRIMARY KEY,driver TEXT NOT NULL,work_date TEXT NOT NULL,route_id INTEGER,delta INTEGER NOT NULL,note TEXT);
 CREATE TABLE IF NOT EXISTS pallet_ledger(id INTEGER PRIMARY KEY,work_date TEXT NOT NULL,movement TEXT NOT NULL,source TEXT,quantity INTEGER NOT NULL,note TEXT);
 CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,at TEXT NOT NULL,actor TEXT,action TEXT NOT NULL,details TEXT);
 CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT NOT NULL);
 INSERT OR IGNORE INTO settings VALUES('milk_rate','0.79'); INSERT OR IGNORE INTO settings VALUES('bread_rate','2.54'); INSERT OR IGNORE INTO settings VALUES('pay_mode','individual');''')
 ws=load_workbook(BASE/'11.26.xlsx',data_only=True)['Табель']
 for r in range(7,207):
  name=ws.cell(r,2).value; role=ws.cell(r,3).value; category=ws.cell(r,4).value
  if name and role: c.execute('INSERT OR IGNORE INTO employees(name,role,category) VALUES(?,?,?)',(str(name).strip(),str(role).strip(),str(category or '').strip()))
 ws=load_workbook(BASE/'группировка на 22.10.2026.xlsx',data_only=True,read_only=True).active
 report_date='2026-10-22'
 for rr in ws.iter_rows(min_row=1,max_row=8,values_only=True):
  for cell in rr:
   if isinstance(cell,str) and 'Период:' in cell:
    m=re.search(r'(\d{2})\.(\d{2})\.(\d{4})',cell)
    if m: report_date=f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
 current=None; routes=[]
 for row in ws.iter_rows(values_only=True):
  a=row[0] if len(row)>0 else None; b=row[1] if len(row)>1 else None; q=row[8] if len(row)>8 else None; note=row[9] if len(row)>9 else None
  if isinstance(a,str) and a.strip() and 'УХУД ООО' in a.upper() and isinstance(q,(int,float)):
   current={'name':a.strip(),'items':[]}; routes.append(current)
  elif current is not None and isinstance(a,str) and a.strip() and isinstance(q,(int,float)):
   current['items'].append((a.strip(),float(q),str(note).strip() if note else None))
 for route in routes:
  digest=hashlib.sha256((route['name']+str(route['items'])).encode()).hexdigest()
  c.execute('INSERT OR IGNORE INTO routes(route_date,name,source_hash,updated_at) VALUES(?,?,?,?)',(report_date,route['name'],digest,datetime.now().isoformat(timespec='seconds')))
  rid=c.execute('SELECT id FROM routes WHERE route_date=? AND name=?',(report_date,route['name'])).fetchone()[0]
  c.execute('DELETE FROM route_items WHERE route_id=?',(rid,))
  for name,qty,note in route['items']:
   c.execute('INSERT INTO route_items(route_id,product_name,quantity,pack_note) VALUES(?,?,?,?)',(rid,name,qty,note))
 c.commit()
 print('Сотрудников:',c.execute('SELECT COUNT(*) FROM employees').fetchone()[0])
 print('Рейсов:',c.execute('SELECT COUNT(*) FROM routes').fetchone()[0])
 print('Товарных строк:',c.execute('SELECT COUNT(*) FROM route_items').fetchone()[0])
