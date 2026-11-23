"""
Procesador de Excel para causación masiva en Siigo.
Lee un archivo .xlsx con columnas de facturas y genera la causación automática.

Columnas esperadas (nombre de columna flexible, se detecta por similitud):
  - NIT / Identificación proveedor
  - Nombre / Razón social
  - Número factura
  - Fecha
  - Subtotal / Valor bruto
  - IVA 19% / IVA
  - IVA 5%
  - Total
  - ReteFuente / Retención fuente
  - ReteIVA / Retención IVA
  - ReteICA / Retención ICA
  - Concepto / Descripción
"""

import re
from decimal import Decimal, InvalidOperation
from datetime import datetime
from typing import Optional, List, Dict, Any

import openpyxl
from openpyxl import load_workbook


class ExcelError(Exception):
    pass


# Mapeo flexible de columnas: variantes → nombre canónico
COLUMNAS_MAPA = {
    'nit':          ['nit', 'identificacion', 'identificación', 'documento', 'nit proveedor',
                     'nit/cc', 'rut'],
    'nombre':       ['nombre', 'razón social', 'razon social', 'proveedor', 'empresa', 'tercero'],
    'numero':       ['numero', 'número', 'factura', 'no. factura', 'nro factura', 'fact',
                     'documento', 'no factura', 'numero factura', 'número factura'],
    'fecha':        ['fecha', 'fecha factura', 'fecha emisión', 'fecha emision', 'date'],
    'subtotal':     ['subtotal', 'valor bruto', 'base', 'base iva', 'valor sin iva',
                     'importe', 'valor neto'],
    'iva19':        ['iva 19', 'iva19', 'iva 19%', '19%', 'iva al 19', 'valor iva 19'],
    'iva5':         ['iva 5', 'iva5', 'iva 5%', '5%', 'iva al 5', 'valor iva 5'],
    'iva':          ['iva', 'valor iva', 'impuesto iva', 'total iva'],
    'total':        ['total', 'valor total', 'total factura', 'total a pagar', 'importe total'],
    'retefuente':   ['retefuente', 'rete fuente', 'retención fuente', 'retencion fuente',
                     'ret fuente', 'retencion en la fuente', 'rf'],
    'reteiva':      ['reteiva', 'rete iva', 'retención iva', 'retencion iva', 'ret iva'],
    'retica':       ['retica', 'rete ica', 'retención ica', 'retencion ica', 'ret ica'],
    'concepto':     ['concepto', 'descripcion', 'descripción', 'detalle', 'glosa',
                     'observaciones', 'notas'],
    'cuenta_gasto': ['cuenta', 'cuenta gasto', 'cuenta contable', 'puc', 'cuenta puc'],
    'centro_costo': ['centro costo', 'centro de costo', 'cc', 'ccosto'],
}


class ProcesadorExcel:
    """
    Lee un archivo .xlsx y extrae las filas de facturas para causación en Siigo.
    """

    def __init__(self, archivo_bytes: bytes, hoja: Optional[str] = None):
        try:
            self.wb = load_workbook(filename=__import__('io').BytesIO(archivo_bytes),
                                    read_only=True, data_only=True)
        except Exception as e:
            raise ExcelError(f'No se pudo abrir el archivo Excel: {e}')

        if hoja:
            if hoja not in self.wb.sheetnames:
                raise ExcelError(f'Hoja "{hoja}" no encontrada. Hojas disponibles: {self.wb.sheetnames}')
            self.ws = self.wb[hoja]
        else:
            self.ws = self.wb.active

        self.mapa_columnas: Dict[str, int] = {}   # nombre_canonico → índice columna (0-based)
        self.errores_globales: List[str] = []

    # ── API pública ───────────────────────────────────────────────────────────

    def procesar(self) -> dict:
        """
        Retorna dict con:
        - facturas: lista de dicts listos para causar en Siigo
        - errores: lista de errores por fila
        - resumen: totales del lote
        """
        self._detectar_cabecera()

        if not self.mapa_columnas:
            raise ExcelError(
                'No se detectaron columnas reconocibles. '
                'El archivo debe tener una fila de encabezados con columnas como: '
                'NIT, Nombre, Número, Fecha, Subtotal, IVA, Total, ReteFuente...'
            )

        filas_datos = list(self.ws.iter_rows(min_row=self._fila_datos, values_only=True))
        facturas = []
        errores  = []

        for idx, fila in enumerate(filas_datos, start=self._fila_datos):
            # Saltar filas completamente vacías
            if all(v is None or str(v).strip() == '' for v in fila):
                continue

            resultado = self._parsear_fila(fila, idx)
            if resultado.get('error'):
                errores.append({'fila': idx, 'error': resultado['error'], 'datos': resultado})
            else:
                facturas.append(resultado)

        resumen = self._calcular_resumen(facturas)

        return {
            'facturas':    facturas,
            'errores':     errores + [{'fila': 0, 'error': e} for e in self.errores_globales],
            'resumen':     resumen,
            'hojas':       self.wb.sheetnames,
            'columnas_detectadas': list(self.mapa_columnas.keys()),
        }

    # ── Detección de cabecera ─────────────────────────────────────────────────

    def _detectar_cabecera(self):
        """Busca la fila de encabezados en las primeras 10 filas."""
        for fila_idx, fila in enumerate(self.ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            mapa = {}
            celdas_texto = 0

            for col_idx, celda in enumerate(fila):
                if celda is None:
                    continue
                texto = str(celda).lower().strip()
                if not texto:
                    continue
                celdas_texto += 1

                for canonico, variantes in COLUMNAS_MAPA.items():
                    for variante in variantes:
                        if variante in texto or texto in variante:
                            if canonico not in mapa:  # primera coincidencia gana
                                mapa[canonico] = col_idx  # 0-based

            # Necesitamos al menos 3 columnas reconocidas para considerar esta como cabecera
            if len(mapa) >= 3:
                self.mapa_columnas = mapa
                self._fila_datos   = fila_idx + 1
                return

        # No se encontró cabecera clara — asumir fila 1
        self._fila_datos = 2
        self.errores_globales.append(
            'No se detectó cabecera clara. Se asume fila 1 como encabezado.'
        )

    # ── Parseo de fila ────────────────────────────────────────────────────────

    def _parsear_fila(self, fila: tuple, numero_fila: int) -> dict:
        def get(col: str):
            idx = self.mapa_columnas.get(col)
            if idx is None or idx >= len(fila):
                return None
            return fila[idx]

        def num(col: str, default=0.0) -> float:
            v = get(col)
            if v is None:
                return float(default)
            try:
                return float(str(v).replace(',', '').replace('$', '').strip() or default)
            except (ValueError, TypeError):
                return float(default)

        def txt(col: str) -> str:
            v = get(col)
            return str(v).strip() if v is not None else ''

        # Campos obligatorios
        nit    = txt('nit')
        total  = num('total')

        if not nit:
            return {'error': f'Fila {numero_fila}: NIT vacío', 'fila': numero_fila}
        if total <= 0:
            return {'error': f'Fila {numero_fila}: Total inválido ({total})', 'fila': numero_fila}

        # IVA — puede venir en columna única o separado por tarifa
        iva19     = num('iva19')
        iva5      = num('iva5')
        iva_total = num('iva')
        if iva_total == 0 and (iva19 > 0 or iva5 > 0):
            iva_total = iva19 + iva5

        # Subtotal: si no existe, calcularlo
        subtotal = num('subtotal')
        if subtotal == 0 and total > 0:
            subtotal = total - iva_total

        # Fecha
        fecha_raw = get('fecha')
        fecha = _parsear_fecha(fecha_raw)

        # Retenciones
        retefuente = num('retefuente')
        reteiva    = num('reteiva')
        retica     = num('retica')
        total_ret  = retefuente + reteiva + retica

        # Neto a pagar (lo que realmente se le pagará al proveedor)
        neto_pagar = total - total_ret

        return {
            'fila':         numero_fila,
            'nit':          _limpiar_nit(nit),
            'nombre':       txt('nombre'),
            'numero':       txt('numero'),
            'fecha':        fecha,
            'concepto':     txt('concepto') or f"Factura {txt('numero')}",
            'cuenta_gasto': txt('cuenta_gasto') or '513595',
            'centro_costo': txt('centro_costo'),
            'subtotal':     subtotal,
            'iva19':        iva19,
            'iva5':         iva5,
            'iva_total':    iva_total,
            'total':        total,
            'retefuente':   retefuente,
            'reteiva':      reteiva,
            'retica':       retica,
            'total_retenciones': total_ret,
            'neto_pagar':   neto_pagar,
            # Estructura para convertir a formato XML/Siigo
            'impuestos': _construir_impuestos(iva19, iva5, subtotal),
            'retenciones': _construir_retenciones(retefuente, reteiva, retica, subtotal, iva_total),
            'error': None,
        }

    # ── Resumen del lote ──────────────────────────────────────────────────────

    def _calcular_resumen(self, facturas: list) -> dict:
        if not facturas:
            return {}
        return {
            'total_facturas': len(facturas),
            'total_subtotal':     sum(f['subtotal']    for f in facturas),
            'total_iva':          sum(f['iva_total']   for f in facturas),
            'total_bruto':        sum(f['total']       for f in facturas),
            'total_retefuente':   sum(f['retefuente']  for f in facturas),
            'total_reteiva':      sum(f['reteiva']     for f in facturas),
            'total_retica':       sum(f['retica']      for f in facturas),
            'total_retenciones':  sum(f['total_retenciones'] for f in facturas),
            'total_neto_pagar':   sum(f['neto_pagar']  for f in facturas),
        }


# ── Helpers ───────────────────────────────────────────────────────────────────

def _limpiar_nit(nit: str) -> str:
    """Extrae solo los dígitos del NIT, sin dígito verificador."""
    s = re.sub(r'[^0-9]', '', str(nit or ''))
    # Si termina con dígito verificador separado por guión, ya fue removido
    return s

def _parsear_fecha(valor) -> str:
    """Convierte varios formatos de fecha a 'YYYY-MM-DD'."""
    if valor is None:
        return datetime.today().strftime('%Y-%m-%d')
    if isinstance(valor, (datetime,)):
        return valor.strftime('%Y-%m-%d')
    if hasattr(valor, 'strftime'):  # date object
        return valor.strftime('%Y-%m-%d')
    s = str(valor).strip()
    # Intentar varios formatos
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d',
                '%d/%m/%y', '%d-%m-%y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    # Si es número (Excel serial date)
    try:
        n = int(float(s))
        base = datetime(1899, 12, 30)
        return (base + __import__('datetime').timedelta(days=n)).strftime('%Y-%m-%d')
    except Exception:
        pass
    return datetime.today().strftime('%Y-%m-%d')

def _construir_impuestos(iva19: float, iva5: float, subtotal: float) -> list:
    imp = []
    if iva19 > 0:
        base19 = subtotal * (iva19 / (subtotal * 0.19)) if subtotal > 0 else iva19 / 0.19
        imp.append({'codigo': '01', 'nombre': 'IVA', 'tarifa': 19.0,
                    'base': round(base19, 2), 'valor': round(iva19, 2)})
    if iva5 > 0:
        base5 = subtotal * (iva5 / (subtotal * 0.05)) if subtotal > 0 else iva5 / 0.05
        imp.append({'codigo': '01', 'nombre': 'IVA', 'tarifa': 5.0,
                    'base': round(base5, 2), 'valor': round(iva5, 2)})
    if not imp and subtotal > 0:
        imp.append({'codigo': '01', 'nombre': 'IVA', 'tarifa': 0.0,
                    'base': subtotal, 'valor': 0.0})
    return imp

def _construir_retenciones(rf: float, rv: float, ri: float, subtotal: float, iva: float) -> list:
    ret = []
    if rf > 0:
        tarifa = round(rf / subtotal * 100, 2) if subtotal > 0 else 0
        ret.append({'tipo': 'ReteFuente', 'base': subtotal,
                    'tarifa': tarifa, 'valor': rf})
    if rv > 0:
        tarifa = round(rv / iva * 100, 2) if iva > 0 else 15.0
        ret.append({'tipo': 'ReteIVA', 'base': iva,
                    'tarifa': tarifa, 'valor': rv})
    if ri > 0:
        tarifa = round(ri / subtotal * 100, 4) if subtotal > 0 else 0
        ret.append({'tipo': 'ReteICA', 'base': subtotal,
                    'tarifa': tarifa, 'valor': ri})
    return ret


def procesar_excel(archivo_bytes: bytes, hoja: Optional[str] = None) -> dict:
    """Función de entrada principal."""
    proc = ProcesadorExcel(archivo_bytes, hoja)
    return proc.procesar()
