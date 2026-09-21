# Control-UCR'S - CORREGIDO 13 PARAMETROS - FINAL
import os, sqlite3
from flask import Flask, request, redirect, session, render_template_string, g, send_file
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
import openpyxl

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'control-ucr-2024-secreto')

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        DATABASE_URL = os.environ.get('DATABASE_URL')
        if DATABASE_URL:
            try:
                import psycopg2
                db = psycopg2.connect(DATABASE_URL, sslmode='require')
            except:
                import psycopg
                db = psycopg.connect(DATABASE_URL, sslmode='require')
            g._is_postgres = True
        else:
            db = sqlite3.connect('ucrs_local.db')
            db.row_factory = sqlite3.Row
            g._is_postgres = False
        g._database = db
    return g._database

def is_postgres(): return getattr(g, '_is_postgres', False)
def adapt(q):
    if not is_postgres(): return q.replace('%s','?').replace('SERIAL PRIMARY KEY','INTEGER PRIMARY KEY AUTOINCREMENT')
    return q

def fetchall(q,p=()):
    db=get_db(); q2=adapt(q)
    if is_postgres():
        try:
            import psycopg2.extras
            cur=db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        except:
            import psycopg.rows
            cur=db.cursor(row_factory=psycopg.rows.dict_row)
    else: cur=db.cursor()
    cur.execute(q2,p); r=cur.fetchall(); cur.close()
    return [dict(x) if not isinstance(x, dict) else x for x in r]

def fetchone(q,p=()): r=fetchall(q,p); return r[0] if r else None
def execute(q,p=()):
    db=get_db(); cur=db.cursor(); cur.execute(adapt(q),p); db.commit(); cur.close()
def safe(r,k,d=""):
    try: v=r.get(k,d) if isinstance(r,dict) else d; return v if v is not None else d
    except: return d

def init_db():
    execute("CREATE TABLE IF NOT EXISTS users (id SERIAL PRIMARY KEY, username TEXT UNIQUE, password TEXT, role TEXT);")
    execute("CREATE TABLE IF NOT EXISTS regiones (id SERIAL PRIMARY KEY, nombre TEXT UNIQUE);")
    execute("CREATE TABLE IF NOT EXISTS tiendas (id SERIAL PRIMARY KEY, region_id INTEGER, nombre TEXT);")
    execute("CREATE TABLE IF NOT EXISTS movimientos (id SERIAL PRIMARY KEY, tipo TEXT, nombre_tienda TEXT, cantidad INTEGER, cantidad_corregida INTEGER, observaciones TEXT, fecha TEXT, usuario TEXT, region TEXT, verif_tienda INTEGER DEFAULT 0, verif_bodega INTEGER DEFAULT 0, tamano TEXT, color TEXT, proveedor TEXT, tipo_ucrs TEXT, placa TEXT);")
    if not fetchall("SELECT * FROM users"):
        execute("INSERT INTO users (username,password,role) VALUES (%s,%s,%s)",('admin',generate_password_hash('1234'),'admin'))

@app.teardown_appcontext
def close_db(e):
    db=getattr(g,'_database',None)
    if db: db.close()

def login_required(f):
    @wraps(f)
    def w(*a,**k):
        if 'user' not in session: return redirect('/login')
        return f(*a,**k)
    return w

BASE='''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"><style>body{background:#e9ecef;padding-bottom:70px;font-family:system-ui}.header-cedi{background:#E70A29;color:#fff;padding:10px 14px;font-weight:900}.card-mov{background:#fff;border:1px solid #dee2e6;border-radius:10px;margin-bottom:8px}.footer{background:#111;color:#fff;text-align:center;padding:10px;position:fixed;bottom:0;left:0;right:0;z-index:999;font-size:.85rem}</style></head><body><nav class="navbar bg-white shadow-sm" style="border-bottom:5px solid #E70A29"><div class="container-fluid px-3"><b>Control-UCR'S</b><div class="d-flex gap-1 align-items-center"><small class="me-2">{{session.user}} ({{session.role}})</small><a href="/" class="btn btn-sm btn-dark">Inicio</a>{% if session.role=='admin' %}<a href="/configuracion" class="btn btn-sm btn-secondary">Config</a><a href="/exportar" class="btn btn-sm btn-success">Excel</a>{% endif %}<a href="/logout" class="btn btn-sm btn-outline-danger">Salir</a></div></div></nav><div class="container-fluid mt-3 px-2">{{content|safe}}</div><div class="footer">Desarrollado por John Carmona</div></body></html>'''

@app.route('/login', methods=['GET','POST'])
def login():
    with app.app_context(): init_db()
    if request.method=='POST':
        u=fetchone("SELECT * FROM users WHERE username=%s",(request.form['username'],))
        if u and check_password_hash(safe(u,'password'), request.form['password']):
            session['user']=safe(u,'username'); session['role']=safe(u,'role'); return redirect('/')
    return '<html><head><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"></head><body style="background:#f2f2f2"><div style="max-width:380px;margin:90px auto;background:#fff;padding:32px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.1);text-align:center"><h3 style="color:#E70A29;font-weight:900">Control-UCR\'S</h3><form method="post"><input name="username" class="form-control mb-2" placeholder="Usuario" required><input name="password" type="password" class="form-control mb-3" placeholder="Contraseña" required><button class="btn w-100 btn-lg" style="background:#E70A29;color:#fff;font-weight:700">Entrar</button></form><small class="text-muted d-block mt-3">admin / 1234</small></div></body></html>'

@app.route('/logout')
def logout(): session.clear(); return redirect('/login')

@app.route('/')
@login_required
def index():
    movs=fetchall("SELECT * FROM movimientos ORDER BY fecha DESC")
    eg=[m for m in movs if safe(m,'tipo')=='egreso']
    def card(m):
        mid=safe(m,'id'); orig=safe(m,'cantidad'); corr=safe(m,'cantidad_corregida')
        mostrar=corr if corr not in [None,"",0,"0"] else orig
        chk='checked' if safe(m,'verif_tienda') else ''; t='no_tienda' if chk else 'tienda'
        return f"<div class='card-mov p-2 d-flex align-items-center'><div style='flex:1'><div><b>{safe(m,'nombre_tienda')}</b> <small class='text-muted'>({safe(m,'region')}) - {safe(m,'placa')}</small></div><small>{safe(m,'tipo_ucrs')} {safe(m,'tamano')} {safe(m,'color')} | Cant: <b>{orig}</b></small><br><small class='text-muted'>{safe(m,'fecha')} - {safe(m,'usuario')}</small></div><div style='width:170px' class='text-end'><label class='border rounded-pill px-2 py-1 bg-light' style='font-size:.75rem'><input type='checkbox' {chk} onchange=\"location.href='/verificar/{mid}/{t}'\"> check tienda</label><form method='post' action='/editar_cantidad'><input type='hidden' name='id' value='{mid}'><input type='number' name='cantidad_corregida' value='{mostrar}' class='form-control form-control-sm rounded-pill mt-1 text-center' onchange='this.form.submit()'></form></div></div>"
    lista="".join([card(m) for m in eg]) or "<div class='p-4 text-center bg-white rounded'>Sin despachos aún</div>"
    role=session.get('role')
    botones="<div style='position:fixed;bottom:40px;left:0;right:0;background:#fff;padding:12px;display:flex;gap:10px;justify-content:center;z-index:1000;border-top:2px solid #eee'><a href='/nuevo?tipo=egreso' class='btn btn-dark px-4 fw-bold'>CEDI → TIENDA</a><a href='/nuevo?tipo=ingreso' class='btn btn-outline-dark px-4 fw-bold'>TIENDA → CEDI</a></div>" if role=='admin' else "<div style='position:fixed;bottom:40px;left:0;right:0;background:#E70A29;padding:12px;display:flex;justify-content:center;z-index:1000'><a href='/nuevo?tipo=ingreso' class='btn btn-light fw-bold px-5'>TIENDA-CEDI</a></div>"
    return render_template_string(BASE, content=f"<div class='header-cedi'>CEDI → TIENDA ({len(eg)})</div><div class='mt-2'>{lista}</div>{botones}<div style='height:90px'></div>")

@app.route('/nuevo', methods=['GET','POST'])
@login_required
def nuevo():
    tipo=request.args.get('tipo','egreso')
    regiones=fetchall("SELECT * FROM regiones ORDER BY nombre")
    tiendas_all=fetchall("SELECT tiendas.nombre as t_nombre, regiones.nombre as r_nombre FROM tiendas JOIN regiones ON tiendas.region_id=regiones.id")
    if request.method=='POST':
        # AQUI ESTABA EL ERROR - AHORA SON 13 %s
        execute("INSERT INTO movimientos (tipo,nombre_tienda,cantidad,cantidad_corregida,observaciones,fecha,usuario,region,tamano,color,proveedor,tipo_ucrs,placa) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(request.form['tipo'],request.form['nombre_tienda'],int(request.form['cantidad']),int(request.form['cantidad']),request.form.get('observaciones',''),datetime.now().isoformat(sep=' ',timespec='minutes'),session['user'],request.form.get('region',''),request.form.get('tamano',''),request.form.get('color',''),request.form.get('proveedor',''),request.form.get('tipo_ucrs',''),request.form.get('placa','')))
        return redirect('/')
    import json
    tiendas_json=[{"region":safe(t,'r_nombre'),"tienda":safe(t,'t_nombre')} for t in tiendas_all]
    opciones=''.join([f"<option value='{safe(r,'nombre')}'>{safe(r,'nombre')}</option>" for r in regiones])
    html=f"""
    <div class='card p-4 mx-auto shadow' style='max-width:540px;border-radius:14px'>
    <h5 class='fw-bold mb-3' style='color:#E70A29'>{'CEDI → TIENDA' if tipo=='egreso' else 'TIENDA → CEDI'}</h5>
    <form method='post'>
    <input type='hidden' name='tipo' value='{tipo}'>
    <select id='regionSelect' name='region' class='form-select mb-2' required><option value=''>Región</option>{opciones}</select>
    <select id='tiendaSelect' name='nombre_tienda' class='form-select mb-3' required><option value=''>Tienda</option></select>
    <div class='row g-2 mb-2'>
        <div class='col-4'><input name='cantidad' type='number' class='form-control' placeholder='Cantidad' required></div>
        <div class='col-8'><input name='placa' class='form-control' placeholder='Placa' required></div>
    </div>
    <div class='row g-2 mb-2'>
        <div class='col-4'>
            <select name='tipo_ucrs' class='form-select' required>
                <option value=''>Tipo UCR</option>
                <option value='Canastillas'>Canastillas</option>
                <option value='Estibas'>Estibas</option>
            </select>
        </div>
        <div class='col-4'>
            <select name='tamano' class='form-select' required>
                <option value=''>Tamaño</option>
                <option value='Grande'>Grande</option>
                <option value='Mediana'>Mediana</option>
                <option value='Pequeña'>Pequeña</option>
            </select>
        </div>
        <div class='col-4'><input name='color' class='form-control' placeholder='Color'></div>
    </div>
    <div class='row g-2 mb-3'>
        <div class='col-6'>
            <select name='proveedor' class='form-select' required>
                <option value=''>Proveedor</option>
                <option value='D1'>D1</option>
                <option value='Proveedor'>Proveedor</option>
            </select>
        </div>
        <div class='col-6'><input name='observaciones' class='form-control' placeholder='Obs'></div>
    </div>
    <button class='btn w-100 btn-lg' style='background:#E70A29;color:#fff;font-weight:800'>Guardar</button>
    </form></div>
    <script>
    const tiendas={json.dumps(tiendas_json)};
    document.getElementById('regionSelect').addEventListener('change',function(){{
        let r=this.value;let sel=document.getElementById('tiendaSelect');
        sel.innerHTML='<option value=>Tienda</option>';
        tiendas.filter(t=>t.region===r).forEach(t=>{{let o=document.createElement('option');o.value=t.tienda;o.textContent=t.tienda;sel.appendChild(o);}})
    }});
    </script>
    """
    return render_template_string(BASE, content=html)

@app.route('/verificar/<int:id>/<string:tipo>')
@login_required
def verificar(id,tipo): execute("UPDATE movimientos SET verif_tienda=%s WHERE id=%s",(1 if tipo=='tienda' else 0, id)); return redirect('/')
@app.route('/editar_cantidad', methods=['POST'])
@login_required
def editar_cantidad():
    try: execute("UPDATE movimientos SET cantidad_corregida=%s WHERE id=%s",(int(request.form.get('cantidad_corregida') or 0), request.form.get('id')))
    except: pass
    return redirect('/')
@app.route('/configuracion', methods=['GET','POST'])
@login_required
def configuracion():
    if session.get('role')!='admin': return "SOLO ADMIN",403
    if request.method=='POST':
        a=request.form.get('action')
        try:
            if a=='crear_region' and request.form.get('nueva_region'): execute("INSERT INTO regiones (nombre) VALUES (%s)",(request.form['nueva_region'].strip(),))
            elif a=='crear_tienda' and request.form.get('nueva_tienda'): execute("INSERT INTO tiendas (region_id,nombre) VALUES (%s,%s)",(request.form['region_id'],request.form['nueva_tienda'].strip()))
            elif a=='crear_usuario': execute("INSERT INTO users (username,password,role) VALUES (%s,%s,%s)",(request.form['new_username'].strip(),generate_password_hash(request.form['new_password']),request.form['new_role']))
        except Exception as e: print(e)
        return redirect('/configuracion')
    regiones=fetchall("SELECT * FROM regiones ORDER BY nombre"); usuarios=fetchall("SELECT id, username, role FROM users ORDER BY username")
    opciones=''.join([f"<option value='{safe(r,'id')}'>{safe(r,'nombre')}</option>" for r in regiones])
    lista_u="".join([f"<span class='badge bg-dark me-1 mb-1 p-2'>{safe(u,'username')} ({safe(u,'role')}) <a href='/borrar_usuario/{safe(u,'id')}' style='color:#ff6b6b'> X</a></span>" for u in usuarios])
    html=f"<div class='container' style='max-width:750px'><div class='card p-3 mb-3'><h6 style='color:#E70A29'>USUARIOS</h6><form method='post' class='row g-2'><input type='hidden' name='action' value='crear_usuario'><div class='col-4'><input name='new_username' class='form-control' placeholder='Usuario' required></div><div class='col-3'><input name='new_password' class='form-control' value='1234' required></div><div class='col-3'><select name='new_role' class='form-select'><option value='operador'>Operador</option><option value='admin'>Admin</option></select></div><div class='col-2'><button class='btn w-100' style='background:#E70A29;color:#fff'>Crear</button></div></form><div class='mt-3'>{lista_u}</div></div><div class='card p-3 mb-3'><h6 style='color:#E70A29'>REGIÓN</h6><form method='post' class='row g-2'><input type='hidden' name='action' value='crear_region'><div class='col-8'><input name='nueva_region' class='form-control'></div><div class='col-4'><button class='btn btn-success w-100'>Crear</button></div></form></div><div class='card p-3'><h6 style='color:#E70A29'>TIENDA</h6><form method='post' class='row g-2'><input type='hidden' name='action' value='crear_tienda'><div class='col-4'><select name='region_id' class='form-select'><option value=''>Región</option>{opciones}</select></div><div class='col-5'><input name='nueva_tienda' class='form-control' placeholder='Tienda'></div><div class='col-3'><button class='btn btn-primary w-100'>Crear</button></div></form></div><a href='/' class='btn btn-secondary w-100 mt-4'>Volver</a></div>"
    return render_template_string(BASE, content=html)
@app.route('/borrar_usuario/<int:id>')
@login_required
def borrar_usuario(id): execute("DELETE FROM users WHERE id=%s",(id,)); return redirect('/configuracion')
@app.route('/exportar')
@login_required
def exportar():
    movs=fetchall("SELECT * FROM movimientos ORDER BY fecha DESC"); wb=openpyxl.Workbook(); ws=wb.active; ws.append(["Fecha","Region","Tienda","Cantidad","Corregida","Placa","Tipo","Tam","Color","Proveedor","Obs","Usuario"])
    for m in movs: ws.append([safe(m,'fecha'),safe(m,'region'),safe(m,'nombre_tienda'),safe(m,'cantidad'),safe(m,'cantidad_corregida'),safe(m,'placa'),safe(m,'tipo_ucrs'),safe(m,'tamano'),safe(m,'color'),safe(m,'proveedor'),safe(m,'observaciones'),safe(m,'usuario')])
    wb.save("reporte.xlsx"); return send_file("reporte.xlsx", as_attachment=True)
with app.app_context(): init_db()
if __name__=='__main__': app.run(host='0.0.0.0', port=5001, debug=True)
