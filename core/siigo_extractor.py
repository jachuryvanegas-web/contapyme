"""
Extractor Siigo para cierre contable mensual.
Descarga automáticamente de Siigo lo que se puede:
  ✅ Balance de prueba (saldos de cuentas)
  ✅ Comprobantes del período
  ✅ Facturas emitidas y recibidas
  ✅ Pagos registrados

Lo que el usuario sube manualmente en Excel:
  📋 Extracto bancario (descarga del banco, no de Siigo)
  📋 Certificados de retención recibidos de clientes
  📋 Nómina (si no está en Siigo)
  📋 Activos fijos y vida útil (primera vez)
"""

from datetime import date
from typing import Optional
import calendar

from core.siigo_client import SiigoClient, SiigoError


class SiigoExtractor:
    """
    Extrae todos los informes necesarios para el cierre mensual desde Siigo.
    """

    def __init__(self, cliente: SiigoClient):
        self.cliente = cliente

    def extraer_todo(self, periodo: str) -> dict:
        """
        periodo: 'YYYY-MM'
        Extrae todo lo necesario del período y retorna un dict consolidado.
        """
        año, mes  = periodo.split('-')
        ultimo    = calendar.monthrange(int(año), int(mes))[1]
        fecha_ini = f'{año}-{mes}-01'
        fecha_fin = f'{año}-{mes}-{ultimo:02d}'

        resultado = {
            'periodo':     periodo,
            'fecha_ini':   fecha_ini,
            'fecha_fin':   fecha_fin,
            'extraido':    {},
            'errores':     [],
            'advertencias': [],
        }

        # 1. Balance de prueba / saldos de cuentas
        try:
            saldos = self._extraer_balance_prueba(fecha_ini, fecha_fin)
            resultado['extraido']['saldos'] = saldos
            resultado['extraido']['total_cuentas'] = len(saldos)
        except Exception as e:
            resultado['errores'].append(f'Balance de prueba: {e}')
            resultado['extraido']['saldos'] = []

        # 2. Comprobantes del período
        try:
            comprobantes = self._extraer_comprobantes(fecha_ini, fecha_fin)
            resultado['extraido']['comprobantes'] = comprobantes
            resultado['extraido']['total_comprobantes'] = len(comprobantes)
        except Exception as e:
            resultado['errores'].append(f'Comprobantes: {e}')
            resultado['extraido']['comprobantes'] = []

        # 3. Facturas de venta del período
        try:
            facturas_venta = self.cliente.listar_facturas(fecha_ini, fecha_fin)
            resultado['extraido']['facturas_venta'] = facturas_venta
            resultado['extraido']['total_fv'] = len(facturas_venta)
        except Exception as e:
            resultado['advertencias'].append(f'Facturas venta: {e}')
            resultado['extraido']['facturas_venta'] = []

        # 4. Pagos / recibos de caja
        try:
            pagos = self._extraer_pagos(fecha_ini, fecha_fin)
            resultado['extraido']['pagos'] = pagos
        except Exception as e:
            resultado['advertencias'].append(f'Pagos: {e}')
            resultado['extraido']['pagos'] = []

        # 5. Lo que falta (manual)
        resultado['manual_requerido'] = self._que_falta_manual(resultado['extraido'])

        return resultado

    def _extraer_balance_prueba(self, fecha_ini: str, fecha_fin: str) -> list:
        """
        Obtiene el balance de prueba usando POST /v1/test-balance-report (Siigo Premium).
        """
        año, mes_ini = fecha_ini[:7].split('-')
        _, mes_fin   = fecha_fin[:7].split('-')
        try:
            data = self.cliente._post('/v1/test-balance-report', {
                'year':                  int(año),
                'month_start':           int(mes_ini),
                'month_end':             int(mes_fin),
                'includes_tax_difference': False,
            })
            # Puede venir como lista directa o como {'results': [...]}
            if isinstance(data, list):
                return data
            return data.get('results', data.get('data', []))
        except SiigoError as e:
            raise Exception(f'No se pudo obtener el balance de prueba: {e}')

    def _consolidar_saldos(self, movimientos: list) -> list:
        """Agrupa movimientos por cuenta para armar el balance de prueba."""
        cuentas = {}
        for mov in movimientos:
            for item in mov.get('items', []):
                cuenta = str(item.get('account', {}).get('code', '') or
                             item.get('account_code', ''))
                if not cuenta:
                    continue
                if cuenta not in cuentas:
                    cuentas[cuenta] = {
                        'account_code': cuenta,
                        'account_name': item.get('account', {}).get('name', '') or cuenta,
                        'debit':  0.0,
                        'credit': 0.0,
                        'balance': 0.0,
                    }
                cuentas[cuenta]['debit']  += float(item.get('debit') or 0)
                cuentas[cuenta]['credit'] += float(item.get('credit') or 0)

        for c in cuentas.values():
            c['balance'] = c['debit'] - c['credit']

        return list(cuentas.values())

    def _extraer_comprobantes(self, fecha_ini: str, fecha_fin: str) -> list:
        """Obtiene comprobantes contables usando GET /v1/journals (Siigo Premium)."""
        try:
            data = self.cliente._get('/v1/journals', params={
                'date_start': fecha_ini,
                'date_end':   fecha_fin,
                'page_size':  200,
            })
            return data.get('results', [])
        except SiigoError as e:
            raise Exception(str(e))

    def _extraer_pagos(self, fecha_ini: str, fecha_fin: str) -> list:
        try:
            data = self.cliente._get('/v1/payments', params={
                'date_start': fecha_ini,
                'date_end':   fecha_fin,
                'page_size':  100,
            })
            return data.get('results', [])
        except SiigoError:
            return []

    def _que_falta_manual(self, extraido: dict) -> list:
        """
        Retorna lista de lo que el usuario DEBE subir manualmente en Excel.
        """
        faltantes = [
            {
                'item':        'Extracto bancario',
                'descripcion': 'Descarga el extracto del banco (Bancolombia, Davivienda, etc.) en Excel o CSV',
                'plantilla':   'extracto_banco',
                'obligatorio': True,
                'columnas':    ['Fecha', 'Descripcion', 'Debito', 'Credito', 'Saldo', 'Referencia'],
            },
            {
                'item':        'Certificados de retención de clientes',
                'descripcion': 'Los certificados de ReteFuente que te practicaron tus clientes (si no los tienes en XML)',
                'plantilla':   'certificados_retencion',
                'obligatorio': False,
                'columnas':    ['NIT cliente', 'Nombre cliente', 'Periodo', 'Concepto', 'Base', 'Tarifa', 'Valor retencion'],
            },
            {
                'item':        'Activos fijos (primera vez o nuevas adquisiciones)',
                'descripcion': 'Tabla de activos con valor de compra, fecha, vida útil y depreciación acumulada',
                'plantilla':   'activos_fijos',
                'obligatorio': False,
                'columnas':    ['Cuenta', 'Descripcion', 'Fecha compra', 'Valor', 'Vida util años', 'Dep acumulada', 'Dep mes'],
            },
        ]

        # Si no hay comprobantes de nómina, es obligatorio
        tiene_nomina = any(
            str(c.get('document_type', '') or '').upper() in ('NOM', 'NOMINA', 'PL')
            for c in extraido.get('comprobantes', [])
        )
        if not tiene_nomina:
            faltantes.insert(1, {
                'item':        'Nómina del período',
                'descripcion': 'Si la nómina no está registrada en Siigo, sube el resumen de liquidación',
                'plantilla':   'nomina',
                'obligatorio': True,
                'columnas':    ['Empleado', 'CC', 'Cargo', 'Salario', 'Dias', 'Devengado', 'Deducciones', 'Neto', 'Aportes empresa'],
            })

        return faltantes


# ── Lector de Excel del banco (plantilla banco) ───────────────────────────────

def leer_extracto_banco(excel_bytes: bytes, hoja: str = None) -> dict:
    """
    Lee el extracto bancario en Excel.
    Compatible con formatos de Bancolombia, Davivienda, Banco de Bogotá (CSV convertido a XLSX).
    """
    import io
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    ws = wb[hoja] if hoja and hoja in wb.sheetnames else wb.active

    # Detectar fila cabecera (buscar palabras clave)
    CLAVES = ['fecha', 'descripcion', 'descripción', 'debito', 'débito',
              'credito', 'crédito', 'saldo', 'valor']
    fila_cab = 1
    mapa = {}

    for ridx, fila in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
        hits = 0
        tmp  = {}
        for cidx, celda in enumerate(fila):
            if celda is None:
                continue
            txt = str(celda).lower().strip()
            for clave in CLAVES:
                if clave in txt:
                    hits += 1
                    if 'fech' in txt:    tmp['fecha'] = cidx
                    elif 'descri' in txt: tmp['descripcion'] = cidx
                    elif 'debit' in txt or 'retiro' in txt or 'cargo' in txt: tmp['debito'] = cidx
                    elif 'credi' in txt or 'abono' in txt or 'consig' in txt: tmp['credito'] = cidx
                    elif 'saldo' in txt: tmp['saldo'] = cidx
                    elif 'ref' in txt or 'nro' in txt: tmp['referencia'] = cidx
        if hits >= 3:
            fila_cab = ridx
            mapa = tmp
            break

    if not mapa:
        return {'error': 'No se detectó formato de extracto bancario. Verifica las columnas.', 'movimientos': []}

    movimientos = []
    for fila in ws.iter_rows(min_row=fila_cab + 1, values_only=True):
        if all(v is None for v in fila):
            continue

        def g(col):
            idx = mapa.get(col)
            return fila[idx] if idx is not None and idx < len(fila) else None

        fecha = g('fecha')
        if fecha is None:
            continue

        # Parsear fecha
        from core.excel_processor import _parsear_fecha
        fecha_str = _parsear_fecha(fecha)

        debito  = _num(g('debito'))
        credito = _num(g('credito'))
        saldo   = _num(g('saldo'))

        movimientos.append({
            'fecha':       fecha_str,
            'descripcion': str(g('descripcion') or '').strip()[:200],
            'debito':      debito,
            'credito':     credito,
            'saldo':       saldo,
            'referencia':  str(g('referencia') or '').strip()[:50],
        })

    saldo_final = movimientos[-1]['saldo'] if movimientos else 0
    total_debitos  = sum(m['debito']  for m in movimientos)
    total_creditos = sum(m['credito'] for m in movimientos)

    return {
        'movimientos':    movimientos,
        'total':          len(movimientos),
        'total_debitos':  total_debitos,
        'total_creditos': total_creditos,
        'saldo_final':    saldo_final,
        'columnas':       list(mapa.keys()),
        'error':          None,
    }


def leer_certificados_retencion(excel_bytes: bytes) -> list:
    """Lee certificados de retención subidos manualmente en Excel."""
    import io, openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    certs = []
    for fila in rows:
        if len(fila) < 5 or all(v is None for v in fila):
            continue
        certs.append({
            'nit_cliente':  str(fila[0] or '').strip(),
            'nombre':       str(fila[1] or '').strip(),
            'periodo':      str(fila[2] or '').strip(),
            'concepto':     str(fila[3] or '').strip(),
            'base':         _num(fila[4]),
            'tarifa':       _num(fila[5]) if len(fila) > 5 else 0,
            'valor':        _num(fila[6]) if len(fila) > 6 else 0,
        })
    return certs


def _num(v) -> float:
    try:
        return float(str(v or 0).replace(',', '').replace('$', '').strip() or 0)
    except (ValueError, TypeError):
        return 0.0
