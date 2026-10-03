# Control-UCR'S - FINAL CON MIGRACION DB
import os, sqlite3
from flask import Flask, request, redirect, session, render_template_string, g, send_file
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
try:
    import openpyxl
except:
    openpyxl=None

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'control-ucr-2024')

def get_db():
    db=getattr(g,'_database',None)
    if db is None:
        url=os.environ.get('DATABASE_URL')
        if url:
            import psycopg2
            db=psycopg2.connect(url, sslmode='require')
            g._is_postgres=True
        else:
            db=sqlite3.connect('ucrs_local.db')
            db.row_factory=sqlite3.Row
            g._is_postgres=False
        g._database=db
    return g._database

def is_pg(): return getattr(g,'_is_postgres',False)
def adapt(q):
    if not is_pg(): return q.replace('%s','?').replace('SERIAL PRIMARY KEY','INTEGER PRIMARY KEY AUTOINCREMENT')
    return q
def fetchall(q,p=()):
    db=get_db()
    if is_pg():
        import psycopg2.extras
        cur=db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    else: cur=db.cursor()
    cur.execute(adapt(q),p); r=cur.fetchall(); cur.close()
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
    # MIGRACION PARA BASE VIEJA
    for col in ["tamano","color","proveedor","tipo_ucrs","placa","cantidad_corregida","region","verif_tienda","verif_bodega","usuario"]:
        try:
            if is_pg(): execute(f"ALTER TABLE movimientos ADD COLUMN IF NOT EXISTS {col} TEXT;")
            else: execute(f"ALTER TABLE movimientos ADD COLUMN {col} TEXT;")
        except: pass
        try:
            if is_pg(): execute(f"ALTER TABLE movimientos ADD COLUMN IF NOT EXISTS {col} INTEGER;")
            else: execute(f"ALTER TABLE movimientos ADD COLUMN {col} INTEGER;")
        except: pass
    if not fetchall("SELECT * FROM users"):
        execute("INSERT INTO users (username,password,role) VALUES (%s,%s,%s)",('admin',generate_password_hash('1234'),'admin'))

@app.teardown_appcontext
def close_db(e):
    db=getattr(g,'_database',None)
    if db:
        try: db.close()
        except: pass

def login_required(f):
    @wraps(f)
    def w(*a,**k):
        if 'user' not in session: return redirect('/login')
        return f(*a,**k)
    return w

BASE='''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"><style>body{background:#e9ecef;padding-bottom:70px;font-family:system-ui}.header-cedi{background:#E70A29;color:#fff;padding:10px 14px;font-weight:900;border-radius:8px}.card-mov{background:#fff;border:1px solid #dee2e6;border-radius:10px;margin-bottom:8px}.footer{background:#111;color:#fff;text-align:center;padding:10px;position:fixed;bottom:0;left:0;right:0;z-index:999}</style></head><body><nav class="navbar bg-white shadow-sm" style="border-bottom:5px solid #E70A29"><div class="container-fluid px-3"><b>Control-UCR'S</b><div class="d-flex gap-1"><small class="me-2">{{session.user}}</small><a href="/" class="btn btn-sm btn-dark">Inicio</a>{% if session.role=='admin' %}<a href="/configuracion" class="btn btn-sm btn-secondary">Config</a>{% endif %}<a href="/logout" class="btn btn-sm btn-outline-danger">Salir</a></div></div></nav><div class="container-fluid mt-3 px-2">{{content|safe}}</div><div class="footer">Desarrollado por John Carmona</div></body></html>'''

@app.route('/login', methods=['GET','POST'])
def login():
    init_db()
    if request.method=='POST':
        u=fetchone("SELECT * FROM users WHERE username=%s",(request.form['username'],))
        if u and check_password_hash(safe(u,'password'), request.form['password']):
            session['user']=safe(u,'username'); session['role']=safe(u,'role'); return redirect('/')
    return '<html><head><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"><style>body{background:#fff;display:flex;align-items:center;justify-content:center;min-height:100vh}.login-box{width:100%;max-width:340px;padding:20px;text-align:center}.login-title{color:#D5001F;font-weight:900;font-size:28px;margin-bottom:18px}.form-control{border-radius:10px;padding:13px}.btn-entrar{background:#D5001F;color:#fff;border:none;border-radius:10px;padding:13px;font-weight:700;width:100%}.badge-dev{margin-top:14px;display:inline-block;border:1px solid #eee;border-radius:8px;padding:6px 14px;font-size:12px}</style></head><body><div class="login-box"><div class="login-title">Control-UCR\'S</div><form method="post"><input name="username" class="form-control mb-3" placeholder="Usuario" required><input name="password" type="password" class="form-control mb-3" placeholder="Contraseña" required><button class="btn-entrar">Entrar</button></form><div class="badge-dev">Desarrollado por John Carmona</div></div></body></html>'

@app.route('/logout')
def logout(): session.clear(); return redirect('/login')

@app.route('/')
@login_required
def index():
    movs=fetchall("SELECT * FROM movimientos ORDER BY fecha DESC")
    eg=[m for m in movs if safe(m,'tipo')=='egreso']
    ing=[m for m in movs if safe(m,'tipo')=='ingreso']
    role=session.get('role','operador')
    def card(m):
        mid=safe(m,'id'); orig=safe(m,'cantidad'); corr=safe(m,'cantidad_corregida'); mostrar=corr if corr not in [None,"",0,"0"] else orig
        chk='checked' if safe(m,'verif_tienda') else ''; t='no_tienda' if chk else 'tienda'
        if role=='operador':
            opts=''.join([f"<option value='{i}' {'selected' if str(i)==str(mostrar) else ''}>{i}</option>" for i in range(0, max(30, int(orig or 0)+15))])
            return f"<div class='card-mov p-2 d-flex align-items-center' style='border-left:5px solid #E70A29'><div style='flex:1'><b>{safe(m,'nombre_tienda')}</b><br><small>{safe(m,'tipo_ucrs')} | Env: {orig}</small></div><div style='width:170px;text-align:right'><label class='border rounded-pill px-2 py-1 bg-light' style='font-size:.7rem'><input type='checkbox' {chk} onchange=\"location.href='/verificar/{mid}/{t}'\"> check tienda</label><form method='post' action='/editar_cantidad'><input type='hidden' name='id' value='{mid}'><select name='cantidad_corregida' class='form-select form-select-sm rounded-pill mt-1 text-center' onchange='this.form.submit()'>{opts}</select></form></div></div>"
        else:
            return f"<div class='card-mov p-2'><b>{safe(m,'nombre_tienda')}</b> ({safe(m,'region')}) - Cant: {orig} -> {mostrar}</div>"
    lista_eg="".join([card(m) for m in eg]) or "<div class='p-3 bg-white rounded text-center'>Sin CEDI->TIENDA</div>"
    lista_ing="".join([f"<div class='card-mov p-2' style='border-left:5px solid #198754'><b>{safe(m,'nombre_tienda')}</b> - {safe(m,'cantidad')} - {safe(m,'placa')}</div>" for m in ing]) or "<div class='p-3 bg-white rounded text-center'>Sin TIENDA->CEDI</div>"
    if role=='operador': botones="<div style='position:fixed;bottom:40px;left:0;right:0;background:#E70A29;padding:12px;display:flex;justify-content:center'><a href='/nuevo?tipo=ingreso' class='btn btn-light fw-bold px-5 rounded-pill'>TIENDA-CEDI</a></div>"
    else: botones="<div style='position:fixed;bottom:40px;left:0;right:0;background:#fff;padding:12px;display:flex;gap:10px;justify-content:center'><a href='/nuevo?tipo=egreso' class='btn btn-dark px-4 rounded-pill'>CEDI→TIENDA</a><a href='/nuevo?tipo=ingreso' class='btn btn-outline-dark px-4 rounded-pill'>TIENDA→CEDI</a></div>"
    return render_template_string(BASE, content=f"<div class='header-cedi'>CEDI→TIENDA ({len(eg)})</div><div class='mt-2 mb-4'>{lista_eg}</div><div class='header-cedi' style='background:#198754'>TIENDA→CEDI ({len(ing)})</div><div class='mt-2'>{lista_ing}</div>{botones}<div style='height:90px'></div>")

@app.route('/nuevo', methods=['GET','POST'])
@login_required
def nuevo():
    tipo=request.args.get('tipo','egreso')
    regiones=fetchall("SELECT * FROM regiones ORDER BY nombre")
    tiendas_all=fetchall("SELECT tiendas.nombre as t_nombre, regiones.nombre as r_nombre FROM tiendas JOIN regiones ON tiendas.region_id=regiones.id")
    if request.method=='POST':
        try:
            # 13 valores = 13 %s - YA CORREGIDO
            execute("INSERT INTO movimientos (tipo,nombre_tienda,cantidad,cantidad_corregida,observaciones,fecha,usuario,region,tamano,color,proveedor,tipo_ucrs,placa) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",(
                request.form['tipo'], request.form['nombre_tienda'], int(request.form['cantidad'] or 0), int(request.form['cantidad'] or 0),
                request.form.get('observaciones',''), datetime.now().isoformat(sep=' ',timespec='minutes'), session['user'],
                request.form.get('region',''), request.form.get('tamano',''), request.form.get('color',''), request.form.get('proveedor',''), request.form.get('tipo_ucrs',''), request.form.get('placa','')
            ))
            return redirect('/')
        except Exception as e:
            return f"<h3>Error al guardar</h3><p>{e}</p><pre>{__import__('traceback').format_exc()}</pre><a href='/'>Volver</a>", 500

    import json
    tiendas_json=[{"region":safe(t,'r_nombre'),"tienda":safe(t,'t_nombre')} for t in tiendas_all]
    opciones=''.join([f"<option value='{safe(r,'nombre')}'>{safe(r,'nombre')}</option>" for r in regiones])
    html=f"""<div class='card p-4 mx-auto shadow' style='max-width:540px'><h5 style='color:#E70A29'>{'CEDI → TIENDA' if tipo=='egreso' else 'TIENDA → CEDI'}</h5>
    <form method='post'><input type='hidden' name='tipo' value='{tipo}'>
    <select id='regionSelect' name='region' class='form-select mb-2' required><option value=''>Región</option>{opciones}</select>
    <select id='tiendaSelect' name='nombre_tienda' class='form-select mb-3' required><option value=''>Tienda</option></select>
    <div class='row g-2 mb-2'><div class='col-4'><input name='cantidad' type='number' class='form-control' placeholder='Cant' required></div><div class='col-8'><input name='placa' class='form-control' placeholder='Placa' required></div></div>
    <div class='row g-2 mb-2'><div class='col-4'><select name='tipo_ucrs' class='form-select' required><option value=''>Tipo</option><option>Canastillas</option><option>Estibas</option></select></div><div class='col-4'><select name='tamano' class='form-select' required><option value=''>Tamaño</option><option>Grande</option><option>Mediana</option><option>Pequeña</option></select></div><div class='col-4'><input name='color' class='form-control' placeholder='Color'></div></div>
    <div class='row g-2 mb-3'><div class='col-6'><select name='proveedor' class='form-select' required><option value=''>Prov</option><option>D1</option><option>Proveedor</option></select></div><div class='col-6'><input name='observaciones' class='form-control' placeholder='Obs'></div></div>
    <button class='btn w-100' style='background:#E70A29;color:#fff;font-weight:800'>Guardar</button></form></div>
    <script>const tiendas={json.dumps(tiendas_json)};document.getElementById('regionSelect').addEventListener('change',function(){{let r=this.value;let s=document.getElementById('tiendaSelect');s.innerHTML='<option value=>Tienda</option>';tiendas.filter(t=>t.region===r).forEach(t=>{{let o=document.createElement('option');o.value=t.tienda;o.textContent=t.tienda;s.appendChild(o);}})}});</script>"""
    return render_template_string(BASE, content=html)

@app.route('/verificar/<int:id>/<string:tipo>')
@login_required
def verificar(id,tipo): execute("UPDATE movimientos SET verif_tienda=%s WHERE id=%s",(1 if tipo=='tienda' else 0, id)); return redirect('/')
@app.route('/editar_cantidad', methods=['POST'])
@login_required
def editar_cantidad(): execute("UPDATE movimientos SET cantidad_corregida=%s WHERE id=%s",(int(request.form.get('cantidad_corregida') or 0), request.form.get('id'))); return redirect('/')
@app.route('/configuracion', methods=['GET','POST'])
@login_required
def configuracion():
    if session.get('role')!='admin': return "SOLO ADMIN",403
    if request.method=='POST':
        a=request.form.get('action')
        if a=='crear_region': execute("INSERT INTO regiones (nombre) VALUES (%s)",(request.form['nueva_region'].strip(),))
        elif a=='crear_tienda': execute("INSERT INTO tiendas (region_id,nombre) VALUES (%s,%s)",(request.form['region_id'],request.form['nueva_tienda'].strip()))
        elif a=='crear_usuario': execute("INSERT INTO users (username,password,role) VALUES (%s,%s,%s)",(request.form['new_username'],generate_password_hash(request.form['new_password']),request.form['new_role']))
        return redirect('/configuracion')
    regiones=fetchall("SELECT * FROM regiones ORDER BY nombre"); usuarios=fetchall("SELECT id, username, role FROM users ORDER BY username")
    opciones=''.join([f"<option value='{safe(r,'id')}'>{safe(r,'nombre')}</option>" for r in regiones])
    lista_u="".join([f"<span class='badge bg-dark me-1 p-2'>{safe(u,'username')} ({safe(u,'role')})</span>" for u in usuarios])
    html=f"<div class='container' style='max-width:750px'><div class='card p-3 mb-3'><h6>USUARIOS</h6><form method='post' class='row g-2'><input type='hidden' name='action' value='crear_usuario'><div class='col-4'><input name='new_username' class='form-control' required></div><div class='col-3'><input name='new_password' class='form-control' value='1234'></div><div class='col-3'><select name='new_role' class='form-select'><option value='operador'>Operador</option><option value='admin'>Admin</option></select></div><div class='col-2'><button class='btn btn-danger w-100'>Crear</button></div></form><div class='mt-3'>{lista_u}</div></div><div class='card p-3 mb-3'><form method='post' class='row g-2'><input type='hidden' name='action' value='crear_region'><div class='col-8'><input name='nueva_region' class='form-control' placeholder='Nueva Region'></div><div class='col-4'><button class='btn btn-success w-100'>Crear</button></div></form></div><div class='card p-3'><form method='post' class='row g-2'><input type='hidden' name='action' value='crear_tienda'><div class='col-4'><select name='region_id' class='form-select'>{opciones}</select></div><div class='col-5'><input name='nueva_tienda' class='form-control' placeholder='Tienda'></div><div class='col-3'><button class='btn btn-primary w-100'>Crear</button></div></form></div><a href='/' class='btn btn-secondary w-100 mt-3'>Volver</a></div>"
    return render_template_string(BASE, content=html)

with app.app_context(): init_db()
if __name__=='__main__': app.run(host='0.0.0.0', port=5001)
