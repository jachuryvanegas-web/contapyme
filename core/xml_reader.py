"""
Motor XML UBL 2.1 Colombiano
Parsea facturas electrónicas DIAN y extrae todos los valores fiscales:
IVA por tarifa, ReteFuente, ReteIVA, ReteICA, CUFE, NIT, etc.
"""

import re
from decimal import Decimal
from datetime import datetime
from typing import Optional
from lxml import etree

# ── Namespaces UBL 2.1 / DIAN ────────────────────────────────────────────────
NS = {
    'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
    'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
    'ext': 'urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2',
    'sts': 'dian:gov:co:facturaelectronica:Structures-2-1',
    'xades': 'http://uri.etsi.org/01903/v1.3.2#',
    'ds':   'http://www.w3.org/2000/09/xmldsig#',
    'fe':   'http://www.dian.gov.co/contratos/facturaelectronica/v1',
    'ccts': 'urn:un:unece:uncefact:documentation:2',
}

# ── Cuentas PUC para causación automática ────────────────────────────────────
PUC = {
    'clientes':          '130505',
    'proveedores':       '220505',
    'ingresos_ventas':   '413595',
    'costo_ventas':      '613595',
    'gasto_servicios':   '513595',
    'iva_por_pagar':     '240801',
    'iva_descontable':   '240802',
    'iva_0':             '240803',
    'retefuente_pagar':  '236540',
    'reteiva_pagar':     '236701',
    'retica_pagar':      '236801',
    'retefuente_favor':  '135517',
    'reteiva_favor':     '135516',
    'retica_favor':      '135518',
}

# ── Tarifas IVA válidas en Colombia ──────────────────────────────────────────
TARIFAS_IVA = {0, 5, 19}

def _t(elem, xpath: str, ns=NS) -> Optional[str]:
    """Extrae el texto del primer elemento que coincida con xpath."""
    nodes = elem.xpath(xpath, namespaces=ns)
    if nodes:
        val = nodes[0]
        return val.text.strip() if hasattr(val, 'text') else str(val).strip()
    return None

def _d(elem, xpath: str, ns=NS) -> Decimal:
    """Extrae un valor decimal (0 si no existe)."""
    v = _t(elem, xpath, ns)
    try:
        return Decimal(v or '0')
    except Exception:
        return Decimal('0')

def _attr(elem, xpath: str, attr: str, ns=NS) -> Optional[str]:
    """Extrae un atributo de un elemento."""
    nodes = elem.xpath(xpath, namespaces=ns)
    return nodes[0].get(attr) if nodes else None


class LecturaXMLError(Exception):
    pass


class LectorFacturaXML:
    """
    Parsea una factura electrónica colombiana en formato UBL 2.1
    y devuelve un dict estructurado con todos los datos fiscales.
    """

    def __init__(self, xml_bytes: bytes):
        try:
            self.root = etree.fromstring(xml_bytes)
        except etree.XMLSyntaxError as e:
            raise LecturaXMLError(f"XML inválido: {e}")

        # Detectar tipo de documento
        tag = self.root.tag
        if 'Invoice' in tag:
            self.tipo_doc = 'FV'   # Factura de venta
        elif 'CreditNote' in tag:
            self.tipo_doc = 'NC'   # Nota crédito
        elif 'DebitNote' in tag:
            self.tipo_doc = 'ND'   # Nota débito
        else:
            raise LecturaXMLError(f"Tipo de documento no reconocido: {tag}")

    # ── API pública ───────────────────────────────────────────────────────────

    def parsear(self) -> dict:
        """
        Retorna un dict completo con todos los datos de la factura.
        """
        cabecera   = self._cabecera()
        emisor     = self._emisor()
        receptor   = self._receptor()
        totales    = self._totales()
        impuestos  = self._impuestos()
        retenciones = self._retenciones()
        lineas     = self._lineas()

        # Validaciones básicas
        errores = self._validar(cabecera, totales, impuestos)

        return {
            'valido':       len(errores) == 0,
            'errores':      errores,
            'tipo_doc':     self.tipo_doc,
            'cabecera':     cabecera,
            'emisor':       emisor,
            'receptor':     receptor,
            'totales':      totales,
            'impuestos':    impuestos,
            'retenciones':  retenciones,
            'lineas':       lineas,
            'resumen_fiscal': self._resumen_fiscal(totales, impuestos, retenciones),
            'cuentas_sugeridas': self._sugerir_cuentas(impuestos, retenciones),
        }

    # ── Cabecera ──────────────────────────────────────────────────────────────

    def _cabecera(self) -> dict:
        r = self.root
        # CUFE / CUDE
        cufe = _t(r, '//cbc:UUID') or _t(r, '//sts:CUFE')

        # Número de factura
        numero = _t(r, 'cbc:ID')

        # Fechas
        fecha_emision  = _t(r, 'cbc:IssueDate')
        hora_emision   = _t(r, 'cbc:IssueTime')
        fecha_venc     = _t(r, 'cac:PaymentMeans/cac:PaymentDueDate') or \
                         _t(r, './/cac:InvoicePeriod/cbc:EndDate')

        # Moneda
        moneda = _attr(r, 'cbc:DocumentCurrencyCode', 'listID') or \
                 _t(r, 'cbc:DocumentCurrencyCode') or 'COP'

        # Tipo de operación DIAN
        tipo_operacion = _t(r, 'cbc:InvoiceTypeCode') or \
                         _t(r, 'cbc:CreditNoteTypeCode') or \
                         _t(r, 'cbc:DebitNoteTypeCode')

        # Referencia a factura original (NC/ND)
        factura_ref = _t(r, './/cac:BillingReference/cac:InvoiceDocumentReference/cbc:ID')
        cufe_ref    = _t(r, './/cac:BillingReference/cac:InvoiceDocumentReference/cbc:UUID')

        # Forma de pago
        medio_pago = _t(r, './/cac:PaymentMeans/cbc:PaymentMeansCode')
        forma_pago = _t(r, './/cac:PaymentTerms/cbc:Note') or \
                     (_t(r, './/cac:PaymentMeans/cbc:PaymentMeansCode') == '10' and 'CONTADO' or 'CREDITO')

        return {
            'cufe':          cufe,
            'numero':        numero,
            'fecha_emision': fecha_emision,
            'hora_emision':  hora_emision,
            'fecha_vencimiento': fecha_venc,
            'moneda':        moneda,
            'tipo_operacion': tipo_operacion,
            'factura_referencia': factura_ref,
            'cufe_referencia': cufe_ref,
            'medio_pago':    medio_pago,
            'forma_pago':    'CONTADO' if medio_pago == '10' else 'CREDITO',
        }

    # ── Emisor (proveedor) ────────────────────────────────────────────────────

    def _emisor(self) -> dict:
        path = './/cac:AccountingSupplierParty/cac:Party'
        return self._parsear_tercero(path)

    # ── Receptor (empresa del usuario) ───────────────────────────────────────

    def _receptor(self) -> dict:
        path = './/cac:AccountingCustomerParty/cac:Party'
        return self._parsear_tercero(path)

    def _parsear_tercero(self, path: str) -> dict:
        r = self.root
        nit       = _t(r, f'{path}/cac:PartyTaxScheme/cbc:CompanyID')
        dv        = _attr(r, f'{path}/cac:PartyTaxScheme/cbc:CompanyID', 'schemeID')
        razon     = _t(r, f'{path}/cac:PartyTaxScheme/cbc:RegistrationName') or \
                    _t(r, f'{path}/cac:PartyLegalEntity/cbc:RegistrationName') or \
                    _t(r, f'{path}/cac:PartyName/cbc:Name')
        regimen   = _t(r, f'{path}/cac:PartyTaxScheme/cbc:TaxLevelCode')
        email     = _t(r, f'{path}/cac:Contact/cbc:ElectronicMail')
        telefono  = _t(r, f'{path}/cac:Contact/cbc:Telephone')
        direccion = _t(r, f'{path}/cac:PhysicalLocation/cac:Address/cbc:Line') or \
                    _t(r, f'{path}/cac:PhysicalLocation/cac:Address/cac:AddressLine/cbc:Line')
        ciudad    = _t(r, f'{path}/cac:PhysicalLocation/cac:Address/cbc:CityName')
        dpto      = _t(r, f'{path}/cac:PhysicalLocation/cac:Address/cbc:CountrySubentity')
        tipo_id   = _attr(r, f'{path}/cac:PartyTaxScheme/cbc:CompanyID', 'schemeName') or 'NIT'

        return {
            'nit':       nit,
            'dv':        dv,
            'tipo_id':   tipo_id,
            'razon_social': razon,
            'regimen':   regimen,
            'email':     email,
            'telefono':  telefono,
            'direccion': direccion,
            'ciudad':    ciudad,
            'departamento': dpto,
        }

    # ── Totales ───────────────────────────────────────────────────────────────

    def _totales(self) -> dict:
        r = self.root

        subtotal    = _d(r, './/cac:LegalMonetaryTotal/cbc:LineExtensionAmount')
        descuentos  = _d(r, './/cac:LegalMonetaryTotal/cbc:AllowanceTotalAmount')
        cargos      = _d(r, './/cac:LegalMonetaryTotal/cbc:ChargeTotalAmount')
        base_imptos = _d(r, './/cac:LegalMonetaryTotal/cbc:TaxExclusiveAmount')
        total_impto = _d(r, './/cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount')
        anticipo    = _d(r, './/cac:LegalMonetaryTotal/cbc:PrepaidAmount')
        total_pagar = _d(r, './/cac:LegalMonetaryTotal/cbc:PayableAmount')

        return {
            'subtotal':              float(subtotal),
            'descuentos':            float(descuentos),
            'cargos':                float(cargos),
            'base_impuestos':        float(base_imptos),
            'total_con_impuestos':   float(total_impto),
            'anticipo':              float(anticipo),
            'total_pagar':           float(total_pagar),
        }

    # ── Impuestos (IVA por tarifa) ────────────────────────────────────────────

    def _impuestos(self) -> list:
        """
        Retorna lista de impuestos agrupados por tarifa:
        [{'tarifa': 19, 'base': 1000000, 'valor': 190000, 'codigo': '01'}, ...]
        Código DIAN: 01=IVA, 04=INC, 03=ICA
        """
        r = self.root
        impuestos = []

        for tax in r.xpath('.//cac:TaxTotal/cac:TaxSubtotal', namespaces=NS):
            base   = _d(tax, 'cbc:TaxableAmount')
            valor  = _d(tax, 'cbc:TaxAmount')
            tarifa = _d(tax, 'cac:TaxCategory/cbc:Percent')
            codigo = _t(tax, 'cac:TaxCategory/cac:TaxScheme/cbc:ID') or '01'
            nombre = _t(tax, 'cac:TaxCategory/cac:TaxScheme/cbc:Name') or 'IVA'

            impuestos.append({
                'codigo':  codigo,
                'nombre':  nombre,
                'tarifa':  float(tarifa),
                'base':    float(base),
                'valor':   float(valor),
            })

        return impuestos

    # ── Retenciones ───────────────────────────────────────────────────────────

    def _retenciones(self) -> list:
        """
        Retorna lista de retenciones practicadas (vienen en factura si el comprador las informó).
        Muchas veces no vienen en el XML; se calculan al causar.
        [{'tipo': 'ReteFuente', 'base': 1000000, 'tarifa': 2.5, 'valor': 25000}, ...]
        """
        r = self.root
        retenciones = []

        # Retenciones como WithholdingTaxTotal (formato estándar UBL)
        for wh in r.xpath('.//cac:WithholdingTaxTotal', namespaces=NS):
            total_wh = _d(wh, 'cbc:TaxAmount')
            for sub in wh.xpath('cac:TaxSubtotal', namespaces=NS):
                base   = _d(sub, 'cbc:TaxableAmount')
                valor  = _d(sub, 'cbc:TaxAmount')
                tarifa = _d(sub, 'cac:TaxCategory/cbc:Percent')
                codigo = _t(sub, 'cac:TaxCategory/cac:TaxScheme/cbc:ID') or ''
                nombre = _t(sub, 'cac:TaxCategory/cac:TaxScheme/cbc:Name') or ''

                tipo = _mapear_tipo_retencion(codigo, nombre)
                retenciones.append({
                    'tipo':   tipo,
                    'codigo': codigo,
                    'nombre': nombre,
                    'base':   float(base),
                    'tarifa': float(tarifa),
                    'valor':  float(valor),
                })

        return retenciones

    # ── Líneas / ítems ────────────────────────────────────────────────────────

    def _lineas(self) -> list:
        r = self.root
        lineas = []

        # InvoiceLine o CreditNoteLine o DebitNoteLine
        line_tag = {
            'FV': 'cac:InvoiceLine',
            'NC': 'cac:CreditNoteLine',
            'ND': 'cac:DebitNoteLine',
        }.get(self.tipo_doc, 'cac:InvoiceLine')

        for line in r.xpath(f'.//{line_tag}', namespaces=NS):
            linea_id  = _t(line, 'cbc:ID')
            cantidad  = _d(line, 'cbc:InvoicedQuantity') or _d(line, 'cbc:CreditedQuantity') or _d(line, 'cbc:DebitedQuantity')
            unidad    = _attr(line, 'cbc:InvoicedQuantity', 'unitCode') or \
                        _attr(line, 'cbc:CreditedQuantity', 'unitCode') or 'UND'
            valor_lin = _d(line, 'cbc:LineExtensionAmount')
            descuento = _d(line, './/cac:AllowanceCharge[cbc:ChargeIndicator="false"]/cbc:Amount')

            # Producto
            codigo    = _t(line, 'cac:Item/cac:SellersItemIdentification/cbc:ID') or \
                        _t(line, 'cac:Item/cac:StandardItemIdentification/cbc:ID')
            descripcion = _t(line, 'cac:Item/cbc:Description') or \
                          _t(line, 'cac:Item/cbc:Name')

            # Precio
            precio_unit = _d(line, 'cac:Price/cbc:PriceAmount')

            # IVA de la línea
            iva_tarifa = _d(line, './/cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cbc:Percent')
            iva_valor  = _d(line, './/cac:TaxTotal/cbc:TaxAmount')
            iva_base   = _d(line, './/cac:TaxTotal/cac:TaxSubtotal/cbc:TaxableAmount') or valor_lin

            lineas.append({
                'id':          linea_id,
                'descripcion': descripcion,
                'codigo':      codigo,
                'unidad':      unidad,
                'cantidad':    float(cantidad),
                'precio_unitario': float(precio_unit),
                'descuento':   float(descuento),
                'valor_linea': float(valor_lin),
                'iva_tarifa':  float(iva_tarifa),
                'iva_base':    float(iva_base),
                'iva_valor':   float(iva_valor),
            })

        return lineas

    # ── Resumen fiscal ────────────────────────────────────────────────────────

    def _resumen_fiscal(self, totales: dict, impuestos: list, retenciones: list) -> dict:
        """
        Consolida los valores más importantes para la causación.
        """
        iva_19 = sum(i['valor'] for i in impuestos if i['codigo'] == '01' and i['tarifa'] == 19)
        iva_5  = sum(i['valor'] for i in impuestos if i['codigo'] == '01' and i['tarifa'] == 5)
        iva_0  = sum(i['base']  for i in impuestos if i['codigo'] == '01' and i['tarifa'] == 0)
        inc    = sum(i['valor'] for i in impuestos if i['codigo'] == '04')
        ica    = sum(i['valor'] for i in impuestos if i['codigo'] == '03')
        total_iva = iva_19 + iva_5

        retefuente = sum(r['valor'] for r in retenciones if r['tipo'] == 'ReteFuente')
        reteiva    = sum(r['valor'] for r in retenciones if r['tipo'] == 'ReteIVA')
        retica     = sum(r['valor'] for r in retenciones if r['tipo'] == 'ReteICA')

        return {
            'subtotal':        totales['subtotal'],
            'iva_19':          iva_19,
            'iva_5':           iva_5,
            'base_exenta':     iva_0,
            'inc':             inc,
            'ica':             ica,
            'total_iva':       total_iva,
            'retefuente':      retefuente,
            'reteiva':         reteiva,
            'retica':          retica,
            'total_retenciones': retefuente + reteiva + retica,
            'total_pagar':     totales['total_pagar'],
            'neto_causar':     totales['total_pagar'] - retefuente - reteiva - retica,
        }

    # ── Cuentas sugeridas ─────────────────────────────────────────────────────

    def _sugerir_cuentas(self, impuestos: list, retenciones: list) -> dict:
        """
        Sugiere las cuentas PUC para la causación automática.
        La dirección depende de si la factura es de venta o compra.
        """
        cuentas = {}

        if self.tipo_doc == 'FV':  # Factura de VENTA (nosotros vendemos)
            cuentas['contraparte']    = PUC['clientes']
            cuentas['ingreso']        = PUC['ingresos_ventas']
            for imp in impuestos:
                if imp['codigo'] == '01':
                    cuentas[f'iva_{int(imp["tarifa"])}'] = PUC['iva_por_pagar']
            for ret in retenciones:
                tipo = ret['tipo']
                if tipo == 'ReteFuente': cuentas['retefuente'] = PUC['retefuente_favor']
                elif tipo == 'ReteIVA':  cuentas['reteiva']    = PUC['reteiva_favor']
                elif tipo == 'ReteICA':  cuentas['retica']     = PUC['retica_favor']

        else:  # Factura de COMPRA (nosotros compramos) — uso más frecuente con XML proveedor
            cuentas['contraparte']    = PUC['proveedores']
            cuentas['gasto']          = PUC['gasto_servicios']
            for imp in impuestos:
                if imp['codigo'] == '01':
                    cuentas[f'iva_{int(imp["tarifa"])}'] = PUC['iva_descontable']
            for ret in retenciones:
                tipo = ret['tipo']
                if tipo == 'ReteFuente': cuentas['retefuente'] = PUC['retefuente_pagar']
                elif tipo == 'ReteIVA':  cuentas['reteiva']    = PUC['reteiva_pagar']
                elif tipo == 'ReteICA':  cuentas['retica']     = PUC['retica_pagar']

        return cuentas

    # ── Validaciones ──────────────────────────────────────────────────────────

    def _validar(self, cabecera: dict, totales: dict, impuestos: list) -> list:
        errores = []

        if not cabecera.get('cufe'):
            errores.append('CUFE ausente — la factura puede no estar autorizada por DIAN')

        if not cabecera.get('numero'):
            errores.append('Número de factura no encontrado')

        if totales['total_pagar'] <= 0:
            errores.append(f'Total a pagar inválido: {totales["total_pagar"]}')

        # Verificar cuadre: subtotal + iva ≈ total_con_impuestos
        suma_iva = sum(i['valor'] for i in impuestos if i['codigo'] == '01')
        esperado = totales['subtotal'] + suma_iva
        diferencia = abs(esperado - totales['total_con_impuestos'])
        if diferencia > 1:  # tolerancia $1 por redondeos
            errores.append(
                f'Descuadre: subtotal({totales["subtotal"]:.0f}) + '
                f'IVA({suma_iva:.0f}) = {esperado:.0f} ≠ '
                f'total_con_impuestos({totales["total_con_impuestos"]:.0f})'
            )

        # Validar tarifas IVA conocidas
        for imp in impuestos:
            if imp['codigo'] == '01' and imp['tarifa'] not in TARIFAS_IVA:
                errores.append(f'Tarifa IVA desconocida: {imp["tarifa"]}%')

        return errores


# ── Helpers ───────────────────────────────────────────────────────────────────

def _mapear_tipo_retencion(codigo: str, nombre: str) -> str:
    """Mapea código DIAN / texto a tipo de retención normalizado."""
    nombre_u = (nombre or '').upper()
    codigo_u = (codigo or '').upper()
    if '06' in codigo_u or 'FUENTE' in nombre_u or 'RETEFTE' in nombre_u:
        return 'ReteFuente'
    if '07' in codigo_u or 'RETEIVA' in nombre_u or 'RETENCIÓN IVA' in nombre_u:
        return 'ReteIVA'
    if '03' in codigo_u or 'ICA' in nombre_u or 'RETICA' in nombre_u:
        return 'ReteICA'
    return nombre or codigo


def leer_xml(xml_bytes: bytes) -> dict:
    """
    Función de entrada principal.
    Recibe los bytes del XML y retorna el dict con todos los datos parseados.
    """
    lector = LectorFacturaXML(xml_bytes)
    return lector.parsear()


# ── Tests inline ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import json, sys

    if len(sys.argv) < 2:
        print("Uso: python xml_reader.py factura.xml")
        sys.exit(1)

    with open(sys.argv[1], 'rb') as f:
        resultado = leer_xml(f.read())

    print(json.dumps(resultado, indent=2, ensure_ascii=False, default=str))
