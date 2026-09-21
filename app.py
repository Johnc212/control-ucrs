# FIX 13 VALORES
def nuevo():
    ...
    if request.method=='POST':
        execute("INSERT INTO movimientos (tipo,nombre_tienda,cantidad,cantidad_corregida,observaciones,fecha,usuario,region,tamano,color,proveedor,tipo_ucrs,placa) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",(request.form['tipo'],request.form['nombre_tienda'],int(request.form['cantidad']),int(request.form['cantidad']),request.form.get('observaciones',''),datetime.now().isoformat(sep=' ',timespec='minutes'),session['user'],request.form.get('region',''),request.form.get('tamano',''),request.form.get('color',''),request.form.get('proveedor',''),request.form.get('tipo_ucrs',''),request.form.get('placa','')))
        return redirect('/')
