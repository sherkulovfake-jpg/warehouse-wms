from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from pathlib import Path
import sqlite3, os, re, hashlib
from datetime import datetime, date

BASE=Path(__file__).resolve().parent
DB=BASE/'wms.db'
app=Flask(__name__); app.secret_key=os.environ.get('WMS_SECRET','change-this-secret')
STATUSES=['Ждёт','Собирается','Собран','Отгружен']

def db():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; c.execute('PRAGMA foreign_keys=ON'); return c

def init_db():
    with db() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS employees(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, role TEXT, category TEXT, active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, unit_weight_kg REAL, units_per_pack INTEGER, location_code TEXT, zone TEXT, active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS routes(id INTEGER PRIMARY KEY, route_date TEXT NOT NULL, name TEXT NOT NULL, status TEXT DEFAULT 'Ждёт', picker_id INTEGER, trays_out INTEGER DEFAULT 0, driver TEXT, source_hash TEXT, updated_at TEXT, UNIQUE(route_date,name), FOREIGN KEY(picker_id) REFERENCES employees(id));
        CREATE TABLE IF NOT EXISTS route_items(id INTEGER PRIMARY KEY, route_id INTEGER NOT NULL, product_id INTEGER, product_name TEXT NOT NULL, quantity REAL NOT NULL DEFAULT 0, pack_note TEXT, picked INTEGER DEFAULT 0, FOREIGN KEY(route_id) REFERENCES routes(id) ON DELETE CASCADE, FOREIGN KEY(product_id) REFERENCES products(id));
        CREATE TABLE IF NOT EXISTS penalties(id INTEGER PRIMARY KEY, employee_id INTEGER NOT NULL, work_date TEXT NOT NULL, amount REAL NOT NULL, note TEXT, FOREIGN KEY(employee_id) REFERENCES employees(id));
        CREATE TABLE IF NOT EXISTS shifts(id INTEGER PRIMARY KEY, employee_id INTEGER NOT NULL, work_date TEXT NOT NULL, present INTEGER DEFAULT 1, FOREIGN KEY(employee_id) REFERENCES employees(id), UNIQUE(employee_id,work_date));
        CREATE TABLE IF NOT EXISTS tray_ledger(id INTEGER PRIMARY KEY, driver TEXT NOT NULL, work_date TEXT NOT NULL, route_id INTEGER, delta INTEGER NOT NULL, note TEXT, FOREIGN KEY(route_id) REFERENCES routes(id));
        CREATE TABLE IF NOT EXISTS pallet_ledger(id INTEGER PRIMARY KEY, work_date TEXT NOT NULL, movement TEXT NOT NULL, source TEXT, quantity INTEGER NOT NULL, note TEXT);
        CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, at TEXT NOT NULL, actor TEXT, action TEXT NOT NULL, details TEXT);
        CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT NOT NULL);
        ''')
        for k,v in [('milk_rate','0.79'),('bread_rate','2.54'),('pay_mode','individual')]:
            c.execute('INSERT OR IGNORE INTO settings VALUES(?,?)',(k,v))

def audit(action,details,actor='Администратор'):
    with db() as c: c.execute('INSERT INTO audit(at,actor,action,details) VALUES(?,?,?,?)',(datetime.now().isoformat(timespec='seconds'),actor,action,details))

def parse_import(path):
    ws=load_workbook(path,data_only=True,read_only=True).active
    report_date='2026-10-22'
    for rr in ws.iter_rows(min_row=1,max_row=8,values_only=True):
        for cell in rr:
            if isinstance(cell,str) and 'Период:' in cell:
                match=re.search(r'(\d{2})\.(\d{2})\.(\d{4})',cell)
                if match: report_date=f'{match.group(3)}-{match.group(2)}-{match.group(1)}'
    current=None; routes=[]; current_doc=None
    for row in ws.iter_rows(values_only=True):
        a=row[0] if len(row)>0 else None; b=row[1] if len(row)>1 else None
        qty=row[8] if len(row)>8 else None; note=row[9] if len(row)>9 else None
        partner=row[5] if len(row)>5 else None; docdate=row[4] if len(row)>4 else None
        if isinstance(a,str) and a.strip() and 'УХУД ООО' in a.upper() and isinstance(qty,(int,float)):
            current={'name':a.strip(),'total':qty,'items':[],'docs':[]}; routes.append(current); current_doc=None
        elif current is not None and isinstance(a,(int,float)) and b:
            current_doc={'number':str(b),'date':str(docdate or ''),'partner':str(partner or '')}; current['docs'].append(current_doc)
        elif current is not None and isinstance(a,str) and a.strip() and isinstance(qty,(int,float)):
            current['items'].append({'name':a.strip(),'quantity':float(qty),'note':str(note).strip() if note else None})
    return routes,report_date

def import_file(path):
    routes,report_date=parse_import(path); imported=0; items=0
    with db() as c:
        for r in routes:
            # The source report period is taken from the filename/report metadata where possible.
            route_date=report_date
            digest=hashlib.sha256((r['name']+str(r['items'])).encode()).hexdigest()
            old=c.execute('SELECT id FROM routes WHERE route_date=? AND name=?',(route_date,r['name'])).fetchone()
            if old:
                rid=old['id']; c.execute('DELETE FROM route_items WHERE route_id=?',(rid,))
                c.execute('UPDATE routes SET source_hash=?,updated_at=? WHERE id=?',(digest,datetime.now().isoformat(timespec='seconds'),rid))
            else:
                cur=c.execute('INSERT INTO routes(route_date,name,source_hash,updated_at) VALUES(?,?,?,?)',(route_date,r['name'],digest,datetime.now().isoformat(timespec='seconds'))); rid=cur.lastrowid; imported+=1
            for it in r['items']:
                p=c.execute('SELECT id FROM products WHERE name=?',(it['name'],)).fetchone()
                pid=p['id'] if p else None
                c.execute('INSERT INTO route_items(route_id,product_id,product_name,quantity,pack_note) VALUES(?,?,?,?,?)',(rid,pid,it['name'],it['quantity'],it['note']))
                items+=1
    audit('Импорт 1С',f'Рейсов: {len(routes)}; новых: {imported}; товарных строк: {items}')
    return len(routes),items

@app.route('/')
def index():
    with db() as c:
        stats={s:c.execute('SELECT COUNT(*) FROM routes WHERE status=?',(s,)).fetchone()[0] for s in STATUSES}
        total=c.execute('SELECT COUNT(*) FROM routes').fetchone()[0]
        items=c.execute('SELECT COUNT(*) FROM route_items').fetchone()[0]
        weight=c.execute('SELECT COALESCE(SUM(ri.quantity*p.unit_weight_kg),0) FROM route_items ri LEFT JOIN products p ON p.id=ri.product_id').fetchone()[0]
        recent=c.execute('SELECT r.*,e.name picker FROM routes r LEFT JOIN employees e ON e.id=r.picker_id ORDER BY route_date DESC,id DESC LIMIT 12').fetchall()
        unreturned=c.execute('SELECT COALESCE(SUM(delta),0) FROM tray_ledger').fetchone()[0]
        emps=c.execute('SELECT COUNT(*) FROM employees WHERE active=1').fetchone()[0]
    return render_template('index.html',stats=stats,total=total,items=items,weight=weight,recent=recent,unreturned=unreturned,emps=emps)

@app.route('/import',methods=['GET','POST'])
def import_page():
    if request.method=='POST':
        f=request.files.get('file')
        if not f or not f.filename.lower().endswith('.xlsx'):
            flash('Выберите файл Excel формата .xlsx','error'); return redirect(url_for('import_page'))
        target=BASE/'uploads'; target.mkdir(exist_ok=True); p=target/secure_filename(f.filename); f.save(p)
        try:
            n,m=import_file(p); flash(f'Импорт завершён: обработано рейсов {n}, товарных строк {m}.','success')
        except Exception as e: flash(f'Ошибка импорта: {e}','error')
        return redirect(url_for('routes_page'))
    return render_template('import.html')

@app.route('/routes')
def routes_page():
    with db() as c:
        rows=c.execute('''SELECT r.*,e.name picker,(SELECT COUNT(*) FROM route_items i WHERE i.route_id=r.id) item_count,
        (SELECT COALESCE(SUM(i.quantity*p.unit_weight_kg),0) FROM route_items i LEFT JOIN products p ON p.id=i.product_id WHERE i.route_id=r.id) weight
        FROM routes r LEFT JOIN employees e ON e.id=r.picker_id ORDER BY route_date DESC,name''').fetchall()
        emps=c.execute('SELECT * FROM employees WHERE active=1 ORDER BY name').fetchall()
    return render_template('routes.html',routes=rows,employees=emps,statuses=STATUSES)

@app.route('/route/<int:rid>',methods=['GET','POST'])
def route_detail(rid):
    with db() as c:
        if request.method=='POST':
            status=request.form.get('status'); picker=request.form.get('picker_id') or None; driver=request.form.get('driver','').strip(); trays=int(request.form.get('trays_out') or 0)
            old=c.execute('SELECT status,trays_out,driver FROM routes WHERE id=?',(rid,)).fetchone()
            if status not in STATUSES: status=old['status']
            c.execute('UPDATE routes SET status=?,picker_id=?,driver=?,trays_out=? WHERE id=?',(status,picker,driver,trays,rid))
            if status=='Отгружен' and old['status']!='Отгружен' and trays:
                c.execute('INSERT INTO tray_ledger(driver,work_date,route_id,delta,note) VALUES(?,?,?,?,?)',(driver or 'Не указан',date.today().isoformat(),rid,trays,'Выдача при отгрузке'))
            audit('Изменение рейса',f'Рейс #{rid}: {old["status"]} → {status}; лотков {trays}')
            return redirect(url_for('route_detail',rid=rid))
        route=c.execute('SELECT r.*,e.name picker FROM routes r LEFT JOIN employees e ON e.id=r.picker_id WHERE r.id=?',(rid,)).fetchone()
        if not route: return 'Рейс не найден',404
        lines=c.execute('SELECT i.*,p.unit_weight_kg,p.units_per_pack,p.location_code FROM route_items i LEFT JOIN products p ON p.id=i.product_id WHERE i.route_id=? ORDER BY COALESCE(p.location_code,"~"),i.product_name',(rid,)).fetchall()
        emps=c.execute('SELECT * FROM employees WHERE active=1 ORDER BY name').fetchall()
    return render_template('route.html',route=route,lines=lines,employees=emps,statuses=STATUSES)

@app.route('/employees')
def employees():
    with db() as c: rows=c.execute('SELECT * FROM employees ORDER BY active DESC,role,name').fetchall()
    return render_template('employees.html',employees=rows)

@app.route('/products',methods=['GET','POST'])
def products():
    with db() as c:
        if request.method=='POST':
            name=request.form['name'].strip(); weight=float(request.form['weight'] or 0); pack=int(request.form['pack'] or 1); loc=request.form.get('location','').strip(); zone=request.form.get('zone','').strip()
            c.execute('INSERT INTO products(name,unit_weight_kg,units_per_pack,location_code,zone) VALUES(?,?,?,?,?) ON CONFLICT(name) DO UPDATE SET unit_weight_kg=excluded.unit_weight_kg,units_per_pack=excluded.units_per_pack,location_code=excluded.location_code,zone=excluded.zone',(name,weight,pack,loc,zone)); audit('Карточка товара',name); return redirect(url_for('products'))
        rows=c.execute('SELECT * FROM products ORDER BY zone,location_code,name').fetchall()
    return render_template('products.html',products=rows)

@app.route('/payroll')
def payroll():
    with db() as c:
        rows=c.execute('''SELECT e.id,e.name,e.role,
        COALESCE(SUM(CASE WHEN p.zone='Молоко' THEN ri.quantity*p.unit_weight_kg ELSE 0 END),0) milk,
        COALESCE(SUM(CASE WHEN p.zone='Хлеб' THEN ri.quantity*p.unit_weight_kg ELSE 0 END),0) bread
        FROM employees e LEFT JOIN routes r ON r.picker_id=e.id LEFT JOIN route_items ri ON ri.route_id=r.id LEFT JOIN products p ON p.id=ri.product_id
        GROUP BY e.id ORDER BY e.name''').fetchall()
        settings=dict(c.execute('SELECT key,value FROM settings').fetchall())
    return render_template('payroll.html',rows=rows,settings=settings)

@app.route('/mobile')
def mobile():
    with db() as c:
        emps=c.execute('SELECT * FROM employees WHERE active=1 ORDER BY name').fetchall()
        routes=c.execute("SELECT r.*,e.name picker FROM routes r LEFT JOIN employees e ON e.id=r.picker_id WHERE r.status IN ('Ждёт','Собирается') ORDER BY route_date,name").fetchall()
    return render_template('mobile.html',employees=emps,routes=routes)

@app.route('/api/pick/<int:item_id>',methods=['POST'])
def pick_item(item_id):
    with db() as c:
        row=c.execute('SELECT route_id,picked FROM route_items WHERE id=?',(item_id,)).fetchone()
        if not row: return jsonify(ok=False),404
        new=0 if row['picked'] else 1; c.execute('UPDATE route_items SET picked=? WHERE id=?',(new,item_id))
        left=c.execute('SELECT COUNT(*) FROM route_items WHERE route_id=? AND picked=0',(row['route_id'],)).fetchone()[0]
        if left==0: c.execute("UPDATE routes SET status='Собран' WHERE id=?",(row['route_id'],))
        else: c.execute("UPDATE routes SET status='Собирается' WHERE id=? AND status='Ждёт'",(row['route_id'],))
        return jsonify(ok=True,picked=bool(new),remaining=left)

@app.route('/map')
def map_page(): return render_template('map.html')

@app.route('/trays',methods=['GET','POST'])
def trays():
    with db() as c:
        if request.method=='POST':
            c.execute('INSERT INTO tray_ledger(driver,work_date,delta,note) VALUES(?,?,?,?)',(request.form['driver'],request.form.get('work_date') or date.today().isoformat(),int(request.form['delta']),request.form.get('note',''))); audit('Движение лотков',request.form['driver']); return redirect(url_for('trays'))
        balances=c.execute('SELECT driver,SUM(delta) balance FROM tray_ledger GROUP BY driver ORDER BY driver').fetchall()
        ledger=c.execute('SELECT * FROM tray_ledger ORDER BY work_date DESC,id DESC LIMIT 100').fetchall()
    return render_template('trays.html',balances=balances,ledger=ledger)

@app.route('/pallets',methods=['GET','POST'])
def pallets():
    with db() as c:
        if request.method=='POST':
            c.execute('INSERT INTO pallet_ledger(work_date,movement,source,quantity,note) VALUES(?,?,?,?,?)',(request.form.get('work_date') or date.today().isoformat(),request.form['movement'],request.form.get('source',''),int(request.form['quantity']),request.form.get('note',''))); audit('Движение поддонов',request.form['movement']); return redirect(url_for('pallets'))
        rows=c.execute('SELECT * FROM pallet_ledger ORDER BY work_date DESC,id DESC').fetchall()
        total=c.execute("SELECT COALESCE(SUM(CASE WHEN movement='Приход' THEN quantity ELSE -quantity END),0) FROM pallet_ledger").fetchone()[0]
    return render_template('pallets.html',rows=rows,total=total)

if __name__=='__main__':
    init_db(); app.run(host='0.0.0.0',port=5000,debug=True)
