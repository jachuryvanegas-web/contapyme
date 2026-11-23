"""
Backend Flask — Facturador DIAN + Siigo
Endpoints para verificación de XML y causación masiva desde Excel.
"""

import os
import json
import sqlite3
import hashlib
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory

# Importar módulos core
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.xml_reader    import leer_xml, LecturaXMLError
from core.siigo_client  import SiigoClient, SiigoError
from core.excel_processor import procesar_excel, ExcelError
from core.siigo_extractor import SiigoExtractor, leer_extracto_banco, leer_certificados_retencion
from core.cierre_contable import CierreContable

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(__file__).parent.parent
DATA_DIR  = Path(os.environ.get('DATA_DIR', BASE_DIR))
DB_PATH   = DATA_DIR / 'facturador.db'
UPLOAD_DIR = DATA_DIR / 'uploads'
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path='')

# ── Base de datos ─────────────────────────────────────────────────────────────

def get_db():
    con = sqlite3.connect(str(DB_PATH))
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA journal_mode=WAL')
    return con

def init_db():
    con = get_db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS empresas (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre      TEXT NOT NULL,
        nit         TEXT NOT NULL UNIQUE,
        creado      TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS configuracion_siigo (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id      INTEGER NOT NULL REFERENCES empresas(id),
        username        TEXT NOT NULL,
        access_key      TEXT NOT NULL,
        partner_id      TEXT DEFAULT '',
        tipo_doc_id     INTEGER DEFAULT 24,
        cuenta_gasto    TEXT DEFAULT '513595',
        activo          INTEGER DEFAULT 1,
        actualizado     TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS facturas_procesadas (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id      INTEGER REFERENCES empresas(id),
        cufe            TEXT UNIQUE,
        numero          TEXT,
        fecha           TEXT,
        nit_proveedor   TEXT,
        nombre_proveedor TEXT,
        subtotal        REAL,
        iva_total       REAL,
        total           REAL,
        retefuente      REAL DEFAULT 0,
        reteiva         REAL DEFAULT 0,
        retica          REAL DEFAULT 0,
        neto_pagar      REAL,
        estado          TEXT DEFAULT 'PENDIENTE',
        siigo_id        TEXT,
        siigo_numero    TEXT,
        origen          TEXT DEFAULT 'XML',
        json_completo   TEXT,
        procesado       TEXT,
        creado          TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS lotes_excel (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id      INTEGER REFERENCES empresas(id),
        nombre_archivo  TEXT,
        total_filas     INTEGER,
        filas_ok        INTEGER,
        filas_error     INTEGER,
        total_monto     REAL,
        estado          TEXT DEFAULT 'PROCESANDO',
        creado          TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS cierres_contables (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_id      INTEGER REFERENCES empresas(id),
        periodo         TEXT NOT NULL,
        estado          TEXT DEFAULT 'BORRADOR',
        datos_siigo     TEXT,
        extracto_banco  TEXT,
        certificados    TEXT,
        estados_fin     TEXT,
        causaciones     TEXT,
        checklist       TEXT,
        creado          TEXT DEFAULT (datetime('now')),
        actualizado     TEXT DEFAULT (datetime('now')),
        UNIQUE(empresa_id, periodo)
    );
    """)
    con.commit()
    con.close()

init_db()

# ── Helpers ───────────────────────────────────────────────────────────────────

def err(msg: str, code: int = 400):
    return jsonify({'ok': False, 'error': msg}), code

def ok(data: dict):
    return jsonify({'ok': True, **data})

def get_siigo_client(empresa_id: int) -> SiigoClient:
    con = get_db()
    cfg = con.execute(
        'SELECT * FROM configuracion_siigo WHERE empresa_id=? AND activo=1',
        (empresa_id,)
    ).fetchone()
    con.close()
    if not cfg:
        raise ValueError('No hay configuración Siigo para esta empresa')
    return SiigoClient(cfg['username'], cfg['access_key'], cfg['partner_id'] or ''), dict(cfg)

# ── Rutas estáticas ───────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(str(BASE_DIR), 'index.html')

# ── API: Empresas ─────────────────────────────────────────────────────────────

@app.route('/api/empresas', methods=['GET'])
def listar_empresas():
    con = get_db()
    rows = con.execute('SELECT * FROM empresas ORDER BY nombre').fetchall()
    con.close()
    return ok({'empresas': [dict(r) for r in rows]})

@app.route('/api/empresas', methods=['POST'])
def crear_empresa():
    d = request.json or {}
    if not d.get('nombre') or not d.get('nit'):
        return err('nombre y nit son obligatorios')
    con = get_db()
    try:
        con.execute('INSERT INTO empresas (nombre, nit) VALUES (?,?)', (d['nombre'], d['nit']))
        con.commit()
        eid = con.execute('SELECT last_insert_rowid()').fetchone()[0]
    except sqlite3.IntegrityError:
        return err('Ya existe una empresa con ese NIT')
    finally:
        con.close()
    return ok({'id': eid, 'mensaje': 'Empresa creada'})

# ── API: Configuración Siigo ──────────────────────────────────────────────────

@app.route('/api/siigo/config', methods=['POST'])
def guardar_config_siigo():
    d = request.json or {}
    eid = d.get('empresa_id')
    if not eid or not d.get('username') or not d.get('access_key'):
        return err('empresa_id, username y access_key son obligatorios')

    # Verificar credenciales antes de guardar
    cliente = SiigoClient(d['username'], d['access_key'], d.get('partner_id', ''))
    prueba = cliente.probar_conexion()
    if not prueba['ok']:
        return err(f'Credenciales Siigo inválidas: {prueba["error"]}')

    con = get_db()
    con.execute('''
        INSERT INTO configuracion_siigo (empresa_id, username, access_key, partner_id,
                                         tipo_doc_id, cuenta_gasto)
        VALUES (?,?,?,?,?,?)
        ON CONFLICT(empresa_id) DO UPDATE SET
            username=excluded.username, access_key=excluded.access_key,
            partner_id=excluded.partner_id, tipo_doc_id=excluded.tipo_doc_id,
            cuenta_gasto=excluded.cuenta_gasto, activo=1,
            actualizado=datetime('now')
    ''', (eid, d['username'], d['access_key'], d.get('partner_id', ''),
          d.get('tipo_doc_id', 24), d.get('cuenta_gasto', '513595')))
    con.commit()
    con.close()
    return ok({'mensaje': 'Configuración Siigo guardada', 'siigo_usuario': prueba.get('usuario')})

@app.route('/api/siigo/probar', methods=['POST'])
def probar_siigo():
    d = request.json or {}
    if not d.get('username') or not d.get('access_key'):
        return err('username y access_key requeridos')
    cliente = SiigoClient(d['username'], d['access_key'], d.get('partner_id', ''))
    result = cliente.probar_conexion()
    if result['ok']:
        return ok(result)
    return err(result['error'])

# ── API: Verificar / subir XML ────────────────────────────────────────────────

@app.route('/api/xml/verificar', methods=['POST'])
def verificar_xml():
    """
    Recibe un archivo XML, lo parsea y retorna todos los valores fiscales.
    No requiere Siigo configurado — solo verificación.
    """
    if 'xml' not in request.files:
        return err('Archivo XML no recibido (campo: xml)')

    archivo = request.files['xml']
    xml_bytes = archivo.read()

    if len(xml_bytes) == 0:
        return err('El archivo XML está vacío')

    try:
        datos = leer_xml(xml_bytes)
    except LecturaXMLError as e:
        return err(str(e))
    except Exception as e:
        return err(f'Error inesperado al leer XML: {e}')

    return ok({'datos': datos})

@app.route('/api/xml/causar', methods=['POST'])
def causar_xml():
    """
    Recibe XML + empresa_id, verifica y causa en Siigo.
    """
    empresa_id = request.form.get('empresa_id')
    if not empresa_id:
        return err('empresa_id requerido')

    if 'xml' not in request.files:
        return err('Archivo XML no recibido')

    archivo = request.files['xml']
    xml_bytes = archivo.read()

    # 1. Parsear XML
    try:
        datos = leer_xml(xml_bytes)
    except LecturaXMLError as e:
        return err(str(e))

    if not datos['valido'] and datos['errores']:
        # Advertencia pero continuamos si tiene CUFE y total
        pass

    cufe = datos['cabecera'].get('cufe')

    # 2. Verificar duplicado
    con = get_db()
    if cufe:
        dup = con.execute('SELECT id, estado FROM facturas_procesadas WHERE cufe=?', (cufe,)).fetchone()
        if dup:
            con.close()
            return err(f'Factura ya procesada (ID={dup["id"]}, estado={dup["estado"]}). CUFE: {cufe}')

    # 3. Obtener cliente Siigo
    try:
        cliente, cfg = get_siigo_client(int(empresa_id))
    except ValueError as e:
        con.close()
        return err(str(e))

    # 4. Causar en Siigo
    try:
        resultado = cliente.causar_factura_compra(datos, {
            'tipo_documento_id': cfg['tipo_doc_id'],
            'cuenta_gasto':      cfg['cuenta_gasto'],
        })
    except SiigoError as e:
        con.close()
        return err(f'Error Siigo: {e}')

    # 5. Guardar en BD
    rf = datos['resumen_fiscal']
    em = datos['emisor']
    cab = datos['cabecera']
    con.execute('''
        INSERT INTO facturas_procesadas
            (empresa_id, cufe, numero, fecha, nit_proveedor, nombre_proveedor,
             subtotal, iva_total, total, retefuente, reteiva, retica, neto_pagar,
             estado, siigo_id, siigo_numero, origen, json_completo, procesado)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'CAUSADA',?,?,?,?,datetime('now'))
    ''', (
        empresa_id, cufe, cab.get('numero'), cab.get('fecha_emision'),
        em.get('nit'), em.get('razon_social'),
        rf['subtotal'], rf['total_iva'], datos['totales']['total_pagar'],
        rf['retefuente'], rf['reteiva'], rf['retica'], rf['neto_causar'],
        resultado.get('siigo_id'), resultado.get('siigo_numero'),
        'XML', json.dumps(datos, ensure_ascii=False, default=str),
    ))
    con.commit()
    con.close()

    return ok({
        'mensaje':      'Factura causada exitosamente en Siigo',
        'siigo':        resultado,
        'resumen':      rf,
        'advertencias': datos.get('errores', []),
    })

# ── API: Carga masiva Excel ───────────────────────────────────────────────────

@app.route('/api/excel/procesar', methods=['POST'])
def procesar_excel_endpoint():
    """
    Recibe un archivo .xlsx, lo procesa y retorna las filas detectadas.
    No causa aún — permite revisar antes de confirmar.
    """
    if 'excel' not in request.files:
        return err('Archivo Excel no recibido (campo: excel)')

    archivo = request.files['excel']
    excel_bytes = archivo.read()
    hoja = request.form.get('hoja')

    try:
        resultado = procesar_excel(excel_bytes, hoja or None)
    except ExcelError as e:
        return err(str(e))
    except Exception as e:
        return err(f'Error al procesar Excel: {e}')

    return ok({
        'facturas':  resultado['facturas'],
        'errores':   resultado['errores'],
        'resumen':   resultado['resumen'],
        'hojas':     resultado['hojas'],
        'columnas':  resultado['columnas_detectadas'],
    })

@app.route('/api/excel/causar', methods=['POST'])
def causar_excel_masivo():
    """
    Recibe JSON con lista de facturas (resultado de /api/excel/procesar)
    y las causa todas en Siigo una por una.
    """
    d = request.json or {}
    empresa_id = d.get('empresa_id')
    facturas   = d.get('facturas', [])

    if not empresa_id:
        return err('empresa_id requerido')
    if not facturas:
        return err('No hay facturas para causar')

    try:
        cliente, cfg = get_siigo_client(int(empresa_id))
    except ValueError as e:
        return err(str(e))

    con = get_db()
    resultados = []
    errores    = []

    for factura in facturas:
        try:
            # Verificar duplicado por número + NIT
            dup = con.execute(
                'SELECT id FROM facturas_procesadas WHERE numero=? AND nit_proveedor=? AND empresa_id=?',
                (factura.get('numero'), factura.get('nit'), empresa_id)
            ).fetchone()
            if dup:
                errores.append({
                    'fila': factura.get('fila'),
                    'numero': factura.get('numero'),
                    'error': 'Duplicado — ya fue causada',
                })
                continue

            # Construir estructura compatible con causar_factura_compra
            datos_xml_simulado = _excel_a_estructura_xml(factura)
            resultado_siigo = cliente.causar_factura_compra(datos_xml_simulado, {
                'tipo_documento_id': cfg['tipo_doc_id'],
                'cuenta_gasto':      factura.get('cuenta_gasto') or cfg['cuenta_gasto'],
            })

            # Guardar
            con.execute('''
                INSERT INTO facturas_procesadas
                    (empresa_id, numero, fecha, nit_proveedor, nombre_proveedor,
                     subtotal, iva_total, total, retefuente, reteiva, retica, neto_pagar,
                     estado, siigo_id, siigo_numero, origen, procesado)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,'CAUSADA',?,?,'EXCEL',datetime('now'))
            ''', (
                empresa_id,
                factura.get('numero'), factura.get('fecha'),
                factura.get('nit'), factura.get('nombre'),
                factura.get('subtotal', 0), factura.get('iva_total', 0), factura.get('total', 0),
                factura.get('retefuente', 0), factura.get('reteiva', 0), factura.get('retica', 0),
                factura.get('neto_pagar', 0),
                resultado_siigo.get('siigo_id'), resultado_siigo.get('siigo_numero'),
            ))
            con.commit()

            resultados.append({
                'fila':          factura.get('fila'),
                'numero':        factura.get('numero'),
                'nit':           factura.get('nit'),
                'siigo_id':      resultado_siigo.get('siigo_id'),
                'siigo_numero':  resultado_siigo.get('siigo_numero'),
                'exito':         True,
            })

        except (SiigoError, Exception) as e:
            errores.append({
                'fila':   factura.get('fila'),
                'numero': factura.get('numero'),
                'error':  str(e),
            })

    con.close()

    return ok({
        'causadas':  len(resultados),
        'errores':   len(errores),
        'detalle_ok':  resultados,
        'detalle_err': errores,
    })

# ── API: Historial ────────────────────────────────────────────────────────────

@app.route('/api/facturas', methods=['GET'])
def listar_facturas():
    empresa_id = request.args.get('empresa_id')
    estado     = request.args.get('estado')
    origen     = request.args.get('origen')
    limit      = int(request.args.get('limit', 100))

    q = 'SELECT * FROM facturas_procesadas WHERE 1=1'
    p = []
    if empresa_id:
        q += ' AND empresa_id=?'; p.append(empresa_id)
    if estado:
        q += ' AND estado=?'; p.append(estado)
    if origen:
        q += ' AND origen=?'; p.append(origen)
    q += ' ORDER BY creado DESC LIMIT ?'; p.append(limit)

    con = get_db()
    rows = con.execute(q, p).fetchall()
    con.close()
    return ok({'facturas': [dict(r) for r in rows], 'total': len(rows)})

@app.route('/api/facturas/<int:fid>', methods=['GET'])
def detalle_factura(fid: int):
    con = get_db()
    row = con.execute('SELECT * FROM facturas_procesadas WHERE id=?', (fid,)).fetchone()
    con.close()
    if not row:
        return err('Factura no encontrada', 404)
    d = dict(row)
    if d.get('json_completo'):
        d['json_completo'] = json.loads(d['json_completo'])
    return ok({'factura': d})

@app.route('/api/resumen', methods=['GET'])
def resumen_dashboard():
    empresa_id = request.args.get('empresa_id')
    q_base = 'SELECT * FROM facturas_procesadas'
    p = []
    if empresa_id:
        q_base += ' WHERE empresa_id=?'
        p.append(empresa_id)

    con = get_db()
    rows = [dict(r) for r in con.execute(q_base, p).fetchall()]
    con.close()

    total_causadas  = sum(1 for r in rows if r['estado'] == 'CAUSADA')
    total_pendientes = sum(1 for r in rows if r['estado'] == 'PENDIENTE')
    suma_iva        = sum(r['iva_total']   or 0 for r in rows)
    suma_retefuente = sum(r['retefuente']  or 0 for r in rows)
    suma_reteiva    = sum(r['reteiva']     or 0 for r in rows)
    suma_retica     = sum(r['retica']      or 0 for r in rows)
    suma_total      = sum(r['total']       or 0 for r in rows)

    return ok({
        'total_facturas':  len(rows),
        'causadas':        total_causadas,
        'pendientes':      total_pendientes,
        'suma_total':      suma_total,
        'suma_iva':        suma_iva,
        'suma_retefuente': suma_retefuente,
        'suma_reteiva':    suma_reteiva,
        'suma_retica':     suma_retica,
        'suma_retenciones': suma_retefuente + suma_reteiva + suma_retica,
    })

# ── Helpers internos ──────────────────────────────────────────────────────────

def _excel_a_estructura_xml(factura: dict) -> dict:
    """Convierte fila de Excel al dict que espera causar_factura_compra."""
    return {
        'cabecera': {
            'cufe':          None,
            'numero':        factura.get('numero', ''),
            'fecha_emision': factura.get('fecha', ''),
            'hora_emision':  None,
            'fecha_vencimiento': factura.get('fecha', ''),
            'moneda':        'COP',
            'medio_pago':    '1',
            'forma_pago':    'CREDITO',
        },
        'emisor': {
            'nit':           factura.get('nit', ''),
            'razon_social':  factura.get('nombre', ''),
            'tipo_id':       'NIT',
            'email':         '',
            'telefono':      '',
            'direccion':     '',
            'ciudad':        '',
            'departamento':  '',
        },
        'receptor':  {},
        'totales': {
            'subtotal':              factura.get('subtotal', 0),
            'descuentos':            0,
            'cargos':                0,
            'base_impuestos':        factura.get('subtotal', 0),
            'total_con_impuestos':   factura.get('total', 0),
            'anticipo':              0,
            'total_pagar':           factura.get('total', 0),
        },
        'impuestos':    factura.get('impuestos', []),
        'retenciones':  factura.get('retenciones', []),
        'lineas': [{
            'id':              '1',
            'descripcion':     factura.get('concepto', 'Compra'),
            'codigo':          'COMPRA',
            'unidad':          'UND',
            'cantidad':        1,
            'precio_unitario': factura.get('subtotal', 0),
            'descuento':       0,
            'valor_linea':     factura.get('subtotal', 0),
            'iva_tarifa':      19.0 if factura.get('iva19', 0) > 0 else 0.0,
            'iva_base':        factura.get('subtotal', 0),
            'iva_valor':       factura.get('iva_total', 0),
        }],
        'resumen_fiscal': {
            'subtotal':        factura.get('subtotal', 0),
            'iva_19':          factura.get('iva19', 0),
            'iva_5':           factura.get('iva5', 0),
            'base_exenta':     0,
            'inc':             0,
            'ica':             0,
            'total_iva':       factura.get('iva_total', 0),
            'retefuente':      factura.get('retefuente', 0),
            'reteiva':         factura.get('reteiva', 0),
            'retica':          factura.get('retica', 0),
            'total_retenciones': factura.get('total_retenciones', 0),
            'total_pagar':     factura.get('total', 0),
            'neto_causar':     factura.get('neto_pagar', 0),
        },
    }


# ── API: Cierre Contable Mensual ──────────────────────────────────────────────

@app.route('/api/cierre/iniciar', methods=['POST'])
def cierre_iniciar():
    """
    Paso 1: Extrae datos de Siigo para el período dado y crea el registro de cierre.
    Body: { empresa_id, periodo: 'YYYY-MM' }
    """
    d = request.json or {}
    empresa_id = d.get('empresa_id')
    periodo    = d.get('periodo')  # 'YYYY-MM'

    if not empresa_id or not periodo:
        return err('empresa_id y periodo son obligatorios')
    if len(periodo) != 7 or periodo[4] != '-':
        return err('periodo debe tener formato YYYY-MM')

    try:
        cliente, cfg = get_siigo_client(int(empresa_id))
    except ValueError as e:
        return err(str(e))

    # Extraer datos de Siigo
    extractor = SiigoExtractor(cliente)
    try:
        datos_siigo = extractor.extraer_todo(periodo)
    except Exception as e:
        datos_siigo = {'periodo': periodo, 'errores': [str(e)], 'extraido': {}, 'manual_requerido': []}

    # Obtener empresa
    con = get_db()
    empresa = con.execute('SELECT * FROM empresas WHERE id=?', (empresa_id,)).fetchone()
    if not empresa:
        con.close()
        return err('Empresa no encontrada')

    # Crear o actualizar cierre
    con.execute('''
        INSERT INTO cierres_contables (empresa_id, periodo, estado, datos_siigo, actualizado)
        VALUES (?,?,'BORRADOR',?,datetime('now'))
        ON CONFLICT(empresa_id, periodo) DO UPDATE SET
            datos_siigo=excluded.datos_siigo,
            estado='BORRADOR',
            actualizado=datetime('now')
    ''', (empresa_id, periodo, json.dumps(datos_siigo, ensure_ascii=False, default=str)))
    con.commit()
    cierre_id = con.execute(
        'SELECT id FROM cierres_contables WHERE empresa_id=? AND periodo=?',
        (empresa_id, periodo)
    ).fetchone()['id']
    con.close()

    return ok({
        'cierre_id':       cierre_id,
        'periodo':         periodo,
        'empresa':         dict(empresa),
        'datos_siigo':     datos_siigo,
        'manual_requerido': datos_siigo.get('manual_requerido', []),
        'mensaje':         f'Cierre {periodo} iniciado. Sube los archivos manuales requeridos.',
    })


@app.route('/api/cierre/<int:cierre_id>/banco', methods=['POST'])
def cierre_banco(cierre_id: int):
    """
    Paso 2: Sube el extracto bancario del período.
    """
    con = get_db()
    cierre = con.execute('SELECT * FROM cierres_contables WHERE id=?', (cierre_id,)).fetchone()
    con.close()
    if not cierre:
        return err('Cierre no encontrado', 404)

    if 'banco' not in request.files:
        return err('Archivo banco no recibido (campo: banco)')

    archivo = request.files['banco']
    cuenta_banco = request.form.get('cuenta_banco', '1110')

    try:
        extracto = leer_extracto_banco(archivo.read(), request.form.get('hoja'))
    except Exception as e:
        return err(f'Error al leer extracto: {e}')

    if extracto.get('error'):
        return err(extracto['error'])

    con = get_db()
    con.execute('''
        UPDATE cierres_contables
        SET extracto_banco=?, actualizado=datetime('now')
        WHERE id=?
    ''', (json.dumps({'cuenta': cuenta_banco, **extracto}, ensure_ascii=False, default=str), cierre_id))
    con.commit()
    con.close()

    return ok({
        'mensaje':        f'{extracto["total"]} movimientos cargados',
        'total_debitos':  extracto['total_debitos'],
        'total_creditos': extracto['total_creditos'],
        'saldo_final':    extracto['saldo_final'],
        'movimientos':    extracto['movimientos'][:10],  # preview
    })


@app.route('/api/cierre/<int:cierre_id>/retenciones', methods=['POST'])
def cierre_retenciones(cierre_id: int):
    """
    Paso 3 (opcional): Sube certificados de retención recibidos de clientes.
    """
    if 'retenciones' not in request.files:
        return err('Archivo no recibido (campo: retenciones)')

    try:
        certs = leer_certificados_retencion(request.files['retenciones'].read())
    except Exception as e:
        return err(f'Error al leer certificados: {e}')

    con = get_db()
    con.execute('''
        UPDATE cierres_contables
        SET certificados=?, actualizado=datetime('now')
        WHERE id=?
    ''', (json.dumps(certs, ensure_ascii=False, default=str), cierre_id))
    con.commit()
    con.close()

    total = sum(c.get('valor', 0) for c in certs)
    return ok({'mensaje': f'{len(certs)} certificados cargados', 'total': total, 'detalle': certs[:5]})


@app.route('/api/cierre/<int:cierre_id>/generar', methods=['POST'])
def cierre_generar(cierre_id: int):
    """
    Paso 4: Genera las causaciones y estados financieros del cierre.
    Puede recibir body con saldos manuales si Siigo no los tiene.
    """
    con = get_db()
    cierre = con.execute('SELECT * FROM cierres_contables WHERE id=?', (cierre_id,)).fetchone()
    if not cierre:
        con.close()
        return err('Cierre no encontrado', 404)

    empresa = con.execute('SELECT * FROM empresas WHERE id=?', (cierre['empresa_id'],)).fetchone()
    con.close()

    datos_siigo = json.loads(cierre['datos_siigo'] or '{}')
    extracto    = json.loads(cierre['extracto_banco'] or '{}')
    certs       = json.loads(cierre['certificados'] or '[]')

    # Crear objeto cierre contable
    cc = CierreContable(dict(empresa), cierre['periodo'])

    # Cargar saldos de Siigo
    saldos = datos_siigo.get('extraido', {}).get('saldos', [])
    if saldos:
        cc.cargar_saldos_siigo(saldos)

    # Cargar extracto bancario
    if extracto.get('movimientos'):
        cc.cargar_extracto_banco(extracto['movimientos'], extracto.get('cuenta', '1110'))

    # Cargar retenciones de certificados
    if certs:
        for cert in certs:
            cuenta = '1355'  # ReteFuente a favor (activo)
            cc.saldos[cuenta] = cc.saldos.get(cuenta, 0) + cert.get('valor', 0)

    # Generar estados financieros
    try:
        estados = cc.generar_estados_financieros()
    except Exception as e:
        return err(f'Error generando estados: {e}')

    # Guardar
    con = get_db()
    con.execute('''
        UPDATE cierres_contables
        SET estados_fin=?, causaciones=?, checklist=?, estado='GENERADO', actualizado=datetime('now')
        WHERE id=?
    ''', (
        json.dumps(estados, ensure_ascii=False, default=str),
        json.dumps(estados.get('causaciones', []), ensure_ascii=False, default=str),
        json.dumps(estados.get('checklist', []), ensure_ascii=False, default=str),
        cierre_id
    ))
    con.commit()
    con.close()

    return ok({
        'mensaje':     f'Cierre {cierre["periodo"]} generado exitosamente',
        'cierre_id':   cierre_id,
        'estados':     estados,
    })


@app.route('/api/cierre/<int:cierre_id>/causar', methods=['POST'])
def cierre_causar_provisiones(cierre_id: int):
    """
    Paso 5: Causa las provisiones/ajustes generados en Siigo.
    """
    con = get_db()
    cierre = con.execute('SELECT * FROM cierres_contables WHERE id=?', (cierre_id,)).fetchone()
    if not cierre:
        con.close()
        return err('Cierre no encontrado', 404)

    causaciones = json.loads(cierre['causaciones'] or '[]')
    if not causaciones:
        con.close()
        return err('No hay causaciones generadas. Ejecuta /generar primero.')

    try:
        cliente, cfg = get_siigo_client(cierre['empresa_id'])
    except ValueError as e:
        con.close()
        return err(str(e))

    resultados = []
    errores    = []

    for caus in causaciones:
        try:
            # Construir comprobante contable para Siigo
            items = []
            for mov in caus.get('movimientos', []):
                items.append({
                    'account': {'code': mov['cuenta']},
                    'customer': None,
                    'description': caus.get('concepto', ''),
                    'debit':  mov.get('db', mov.get('debito', 0)),
                    'credit': mov.get('cr', mov.get('credito', 0)),
                })

            payload = {
                'document': {'id': cfg.get('tipo_doc_id', 24)},
                'date': f'{cierre["periodo"]}-{_ultimo_dia(cierre["periodo"])}',
                'description': caus.get('concepto', ''),
                'observations': f'Cierre {cierre["periodo"]} - {caus.get("tipo", "")}',
                'items': items,
            }

            resp = cliente._post('/v1/journal-vouchers', payload)
            resultados.append({
                'tipo':     caus.get('tipo'),
                'concepto': caus.get('concepto'),
                'valor':    caus.get('valor_total', 0),
                'siigo_id': resp.get('id'),
                'exito':    True,
            })
        except Exception as e:
            errores.append({'tipo': caus.get('tipo'), 'error': str(e)})

    con = get_db()
    con.execute('''
        UPDATE cierres_contables SET estado=?, actualizado=datetime('now') WHERE id=?
    ''', ('CAUSADO' if not errores else 'PARCIAL', cierre_id))
    con.commit()
    con.close()

    return ok({
        'causadas': len(resultados),
        'errores':  len(errores),
        'detalle':  resultados,
        'fallos':   errores,
    })


@app.route('/api/cierre/<int:cierre_id>', methods=['GET'])
def cierre_detalle(cierre_id: int):
    """Retorna todo el detalle del cierre."""
    con = get_db()
    cierre  = con.execute('SELECT * FROM cierres_contables WHERE id=?', (cierre_id,)).fetchone()
    empresa = None
    if cierre:
        empresa = con.execute('SELECT * FROM empresas WHERE id=?', (cierre['empresa_id'],)).fetchone()
    con.close()
    if not cierre:
        return err('Cierre no encontrado', 404)

    d = dict(cierre)
    for campo in ('datos_siigo', 'extracto_banco', 'certificados', 'estados_fin', 'causaciones', 'checklist'):
        if d.get(campo):
            try:
                d[campo] = json.loads(d[campo])
            except Exception:
                pass

    return ok({'cierre': d, 'empresa': dict(empresa) if empresa else {}})


@app.route('/api/cierres', methods=['GET'])
def listar_cierres():
    empresa_id = request.args.get('empresa_id')
    con = get_db()
    q = 'SELECT id, empresa_id, periodo, estado, creado, actualizado FROM cierres_contables'
    p = []
    if empresa_id:
        q += ' WHERE empresa_id=?'; p.append(empresa_id)
    q += ' ORDER BY periodo DESC'
    rows = con.execute(q, p).fetchall()
    con.close()
    return ok({'cierres': [dict(r) for r in rows]})


def _ultimo_dia(periodo: str) -> str:
    """Retorna el último día del mes como string '28'/'30'/'31'."""
    import calendar as _cal
    año, mes = periodo.split('-')
    return str(_cal.monthrange(int(año), int(mes))[1])


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(debug=True, host='0.0.0.0', port=port)
