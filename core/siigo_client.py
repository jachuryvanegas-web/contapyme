"""
Cliente Siigo API v1
Autenticación OAuth2, creación de proveedores, carga de facturas de compra.
Docs: https://siigoapi.docs.apiary.io/
"""

import json
import requests
from datetime import datetime, timedelta
from typing import Optional


class SiigoError(Exception):
    pass


class SiigoClient:
    """
    Cliente para Siigo API con OAuth2 y manejo de token automático.
    """

    BASE_URL = 'https://api.siigo.com'

    def __init__(self, username: str, access_key: str, partner_id: str = ''):
        self.username    = username
        self.access_key  = access_key
        self.partner_id  = partner_id
        self._token      = None
        self._token_exp  = None
        self._session    = requests.Session()
        self._session.headers.update({
            'Content-Type': 'application/json',
            'Accept':       'application/json',
        })
        if partner_id:
            self._session.headers['Partner-Id'] = partner_id

    # ── Autenticación ─────────────────────────────────────────────────────────

    def autenticar(self) -> str:
        """Obtiene y almacena el token OAuth2 de Siigo."""
        resp = self._session.post(
            f'{self.BASE_URL}/auth',
            json={'username': self.username, 'access_key': self.access_key},
            timeout=30,
        )
        if resp.status_code != 200:
            raise SiigoError(f'Autenticación fallida ({resp.status_code}): {resp.text}')

        data = resp.json()
        self._token     = data.get('access_token')
        expires_in      = data.get('expires_in', 3600)
        self._token_exp = datetime.utcnow() + timedelta(seconds=expires_in - 60)
        self._session.headers['Authorization'] = f'Bearer {self._token}'
        return self._token

    def _asegurar_token(self):
        """Renueva el token si está expirado o no existe."""
        if not self._token or datetime.utcnow() >= self._token_exp:
            self.autenticar()

    def _get(self, path: str, params: dict = None) -> dict:
        self._asegurar_token()
        resp = self._session.get(f'{self.BASE_URL}{path}', params=params, timeout=30)
        if resp.status_code not in (200, 201):
            raise SiigoError(f'GET {path} → {resp.status_code}: {resp.text}')
        return resp.json()

    def _post(self, path: str, body: dict) -> dict:
        self._asegurar_token()
        resp = self._session.post(f'{self.BASE_URL}{path}', json=body, timeout=30)
        if resp.status_code not in (200, 201):
            raise SiigoError(f'POST {path} → {resp.status_code}: {resp.text}')
        return resp.json()

    # ── Terceros / Proveedores ─────────────────────────────────────────────────

    def buscar_proveedor(self, nit: str) -> Optional[dict]:
        """Busca un proveedor en Siigo por NIT. Retorna el primero encontrado o None."""
        try:
            data = self._get('/v1/customers', params={'identification': nit})
            results = data.get('results', [])
            if results:
                return results[0]
        except SiigoError:
            pass
        return None

    def crear_proveedor(self, tercero: dict) -> dict:
        """
        Crea un proveedor en Siigo a partir del dict del emisor del XML.
        Parámetros tercero: nit, razon_social, email, telefono, direccion, ciudad, tipo_id
        """
        # Buscar si ya existe
        existente = self.buscar_proveedor(tercero.get('nit', ''))
        if existente:
            return existente

        tipo_persona = 'Person' if tercero.get('tipo_id') in ('CC', 'CE', 'TI', 'PA') else 'Company'

        body = {
            'type':        'Customer',
            'person_type': tipo_persona,
            'id_type': {
                'code': _mapear_tipo_id_siigo(tercero.get('tipo_id', 'NIT')),
            },
            'identification': _limpiar_nit(tercero.get('nit', '')),
            'name': [tercero.get('razon_social', '')][:1],
            'commercial_name': tercero.get('razon_social', ''),
            'contacts': [
                {
                    'first_name': tercero.get('razon_social', ''),
                    'email':      tercero.get('email', ''),
                    'phone': {
                        'number': (tercero.get('telefono') or '')[:10],
                    },
                }
            ],
            'address': {
                'address':     tercero.get('direccion', ''),
                'city': {
                    'country_code':   'Co',
                    'state_code':     _cod_dpto(tercero.get('departamento', '')),
                    'city_code':      _cod_ciudad(tercero.get('ciudad', '')),
                },
            },
            'fiscal_responsibilities': [{'code': 'R-99-PN'}],
        }

        return self._post('/v1/customers', body)

    # ── Tipos de comprobante ──────────────────────────────────────────────────

    def obtener_tipos_documento(self, tipo: str = 'FV') -> list:
        """
        Retorna los tipos de documento configurados en Siigo.
        tipo: 'FV'=Factura venta, 'FC'=Factura compra, etc.
        """
        mapeo = {'FV': 'FV', 'FC': 'FC', 'NC': 'NC', 'ND': 'ND'}
        data = self._get('/v1/document-types', params={'type': mapeo.get(tipo, tipo)})
        return data.get('results', [])

    # ── Causación de factura de compra ────────────────────────────────────────

    def causar_factura_compra(self, datos_xml: dict, config: dict) -> dict:
        """
        Envía la causación de una factura de compra a Siigo.

        datos_xml: resultado de xml_reader.leer_xml()
        config: {
            'tipo_documento_id': 24,  # ID del tipo FV en Siigo (obtener de obtener_tipos_documento)
            'cuenta_gasto': '513595', # cuenta de gasto por defecto
            'centro_costo_id': None,  # opcional
        }
        """
        cab = datos_xml['cabecera']
        em  = datos_xml['emisor']
        tot = datos_xml['totales']
        rf  = datos_xml['resumen_fiscal']
        imp = datos_xml['impuestos']
        ret = datos_xml['retenciones']
        lin = datos_xml['lineas']

        # 1. Asegurar que el proveedor exista en Siigo
        proveedor = self.crear_proveedor(em)
        proveedor_id = proveedor.get('id') or proveedor.get('code')

        # 2. Construir items de la factura
        items = []
        for linea in lin:
            item = {
                'code':        linea.get('codigo') or 'COMPRA',
                'description': (linea.get('descripcion') or 'Compra')[:255],
                'quantity':    linea['cantidad'],
                'price':       linea['precio_unitario'],
                'discount':    linea.get('descuento', 0),
                'taxes': [],
            }
            # IVA de la línea
            if linea['iva_tarifa'] > 0:
                item['taxes'].append({
                    'id': _id_impuesto_siigo(linea['iva_tarifa']),
                })
            items.append(item)

        # Si no hay líneas detalladas, crear una línea con el subtotal
        if not items:
            items = [{
                'code':        'COMPRA',
                'description': f"Factura {cab.get('numero', '')}",
                'quantity':    1,
                'price':       tot['subtotal'],
                'discount':    0,
                'taxes':       _construir_taxes_siigo(imp),
            }]

        # 3. Construir retenciones (WithholdingTaxes)
        retenciones_siigo = []
        for r in ret:
            tipo_siigo = _tipo_retencion_siigo(r['tipo'])
            if tipo_siigo:
                retenciones_siigo.append({
                    'id': tipo_siigo,
                    'value': r['valor'],
                })

        # 4. Fecha
        fecha = cab.get('fecha_emision') or datetime.today().strftime('%Y-%m-%d')

        # 5. Payload principal
        body = {
            'document': {
                'id': config.get('tipo_documento_id', 24),
            },
            'date': fecha,
            'currency': {
                'code':          cab.get('moneda', 'COP'),
                'exchange_rate': 1,
            },
            'customer': {
                'id': proveedor_id,
                'branch_office': 0,
            },
            'seller': 0,
            'observations': (
                f"Factura electrónica {cab.get('numero','')} "
                f"CUFE: {(cab.get('cufe') or '')[:20]}"
            ),
            'items': items,
            'payments': [
                {
                    'id':    _medio_pago_siigo(cab.get('medio_pago', '')),
                    'value': tot['total_pagar'],
                    'due_date': cab.get('fecha_vencimiento') or fecha,
                }
            ],
        }

        if retenciones_siigo:
            body['withholding_taxes'] = retenciones_siigo

        # 6. Enviar
        resultado = self._post('/v1/invoices', body)

        return {
            'exito':          True,
            'siigo_id':       resultado.get('id'),
            'siigo_numero':   resultado.get('number'),
            'siigo_prefijo':  resultado.get('prefix'),
            'fecha':          fecha,
            'proveedor_id':   proveedor_id,
            'total_causado':  tot['total_pagar'],
        }

    # ── Consulta ──────────────────────────────────────────────────────────────

    def consultar_factura(self, factura_id: str) -> dict:
        return self._get(f'/v1/invoices/{factura_id}')

    def listar_facturas(self, fecha_desde: str, fecha_hasta: str) -> list:
        data = self._get('/v1/invoices', params={
            'date_start': fecha_desde,
            'date_end':   fecha_hasta,
        })
        return data.get('results', [])

    def probar_conexion(self) -> dict:
        """Verifica que las credenciales son correctas."""
        try:
            self.autenticar()
            info = self._get('/v1/users/me')
            return {'ok': True, 'usuario': info.get('username'), 'empresa': info.get('company')}
        except SiigoError as e:
            return {'ok': False, 'error': str(e)}


# ── Helpers de mapeo ──────────────────────────────────────────────────────────

def _limpiar_nit(nit: str) -> str:
    """Elimina puntos, guiones y dígito verificador del NIT."""
    return re.sub(r'[^0-9]', '', nit or '').rstrip('0') if nit else ''

import re

def _limpiar_nit(nit: str) -> str:
    return re.sub(r'[^0-9]', '', nit or '')

def _mapear_tipo_id_siigo(tipo: str) -> str:
    """Mapea tipo de identificación DIAN → código Siigo."""
    mapa = {
        'NIT':  '31',
        'CC':   '13',
        'CE':   '22',
        'TI':   '12',
        'PA':   '41',
        'RUT':  '31',
    }
    return mapa.get((tipo or '').upper(), '31')

def _id_impuesto_siigo(tarifa: float) -> int:
    """Retorna el ID de impuesto en Siigo según tarifa IVA."""
    mapa = {19: 1, 5: 5, 0: 3}
    return mapa.get(int(tarifa), 1)

def _construir_taxes_siigo(impuestos: list) -> list:
    taxes = []
    for imp in impuestos:
        if imp['codigo'] == '01':  # IVA
            taxes.append({'id': _id_impuesto_siigo(imp['tarifa'])})
    return taxes

def _tipo_retencion_siigo(tipo: str) -> Optional[int]:
    """ID de tipo de retención en Siigo."""
    mapa = {
        'ReteFuente': 1,
        'ReteIVA':    5,
        'ReteICA':    6,
    }
    return mapa.get(tipo)

def _medio_pago_siigo(codigo_dian: str) -> int:
    """Mapea código de medio pago DIAN → ID Siigo."""
    mapa = {
        '10': 1,  # Efectivo / Contado
        '20': 2,  # Cheque
        '31': 3,  # Transferencia
        '42': 3,  # Débito bancario
        '48': 3,  # Tarjeta crédito
        '49': 3,  # Tarjeta débito
    }
    return mapa.get(str(codigo_dian), 1)

def _cod_dpto(nombre: str) -> str:
    """Código de departamento colombiano (simplificado)."""
    mapa = {
        'BOGOTA': '11', 'CUNDINAMARCA': '25', 'ANTIOQUIA': '05',
        'VALLE':  '76', 'ATLANTICO':    '08', 'BOLIVAR':   '13',
        'SANTANDER': '68', 'CALDAS': '17', 'RISARALDA': '66',
        'TOLIMA': '73', 'HUILA': '41', 'NARIÑO': '52',
        'CAUCA': '19', 'BOYACA': '15', 'META': '50',
    }
    n = (nombre or '').upper()
    for k, v in mapa.items():
        if k in n:
            return v
    return '11'

def _cod_ciudad(nombre: str) -> str:
    """Código DANE de ciudad (simplificado — devuelve Bogotá por defecto)."""
    mapa = {
        'BOGOTA': '11001', 'MEDELLIN': '05001', 'CALI': '76001',
        'BARRANQUILLA': '08001', 'CARTAGENA': '13001', 'BUCARAMANGA': '68001',
        'MANIZALES': '17001', 'PEREIRA': '66001', 'IBAGUE': '73001',
        'NEIVA': '41001', 'PASTO': '52001', 'POPAYAN': '19001',
        'TUNJA': '15001', 'VILLAVICENCIO': '50001',
    }
    n = (nombre or '').upper()
    for k, v in mapa.items():
        if k in n:
            return v
    return '11001'


Optional = type(None)  # noqa: redefinición local para evitar import circular
try:
    from typing import Optional
except ImportError:
    pass
