"""
Módulo de Cierre Contable Mensual
==================================
1. Extrae de Siigo: movimientos, saldos de cuentas, comprobantes del período
2. Recibe Excel de banco (extracto) y retenciones recibidas
3. Genera causaciones faltantes: ajustes, provisiones, depreciación
4. Produce estados financieros con notas NIIF

Qué baja de Siigo automáticamente (con credenciales):
  - Saldos de cuentas (balance de prueba)
  - Comprobantes del período
  - Facturas emitidas y recibidas
  - Movimientos de bancos registrados

Qué sube el usuario en Excel manualmente:
  - Extracto bancario (CSV/XLSX del banco)
  - Certificados de retención recibidos (si no los tiene en XML)
  - Nómina del período (si no está en Siigo)
  - Activos fijos y tabla de depreciación
"""

from datetime import datetime, date
from decimal import Decimal
from typing import Optional
import calendar


# ── Cuentas PUC relevantes para cierre ────────────────────────────────────────
PUC_CIERRE = {
    # Activo
    '1105': 'Caja',
    '1110': 'Bancos',
    '1120': 'Cuentas de ahorro',
    '1305': 'Clientes',
    '1310': 'Clientes - sector financiero',
    '1330': 'Anticipos y avances',
    '1355': 'Anticipo de impuestos',
    '1360': 'Reclamaciones',
    '1365': 'Cuentas por cobrar a trabajadores',
    '1370': 'Deudores varios',
    '1380': 'Deudas de difícil cobro',
    '1399': 'Provisión deudores',
    '1424': 'Inventario de mercancías',
    '1465': 'Inventario materias primas',
    '1499': 'Provisión inventarios',
    '1504': 'Terrenos',
    '1516': 'Construcciones y edificaciones',
    '1520': 'Maquinaria y equipo',
    '1524': 'Equipo de oficina',
    '1528': 'Equipo de computación',
    '1540': 'Flota y equipo de transporte',
    '1592': 'Depreciación acumulada propiedades',
    '1910': 'Diferidos',
    '1970': 'Intangibles',
    '1992': 'Amortización acumulada',
    # Pasivo
    '2105': 'Obligaciones financieras CP',
    '2205': 'Proveedores nacionales',
    '2335': 'Costos y gastos por pagar',
    '2360': 'Dividendos por pagar',
    '2365': 'Retención en la fuente',
    '2367': 'Impuesto a las ventas retenido',
    '2368': 'Industria y comercio retenido',
    '2370': 'Retenciones y aportes de nómina',
    '2380': 'Acreedores',
    '2404': 'IVA por pagar',
    '2408': 'IVA descontable',
    '2505': 'Obligaciones financieras LP',
    '2610': 'Para obligaciones laborales',
    '2615': 'Para costos y gastos',
    '2685': 'Otros pasivos estimados',
    # Patrimonio
    '3105': 'Capital suscrito y pagado',
    '3305': 'Reserva legal',
    '3405': 'Resultado ejercicios anteriores',
    '3605': 'Utilidad del ejercicio',
    '3610': 'Pérdida del ejercicio',
    # Ingresos
    '4135': 'Comercio al por mayor y menor',
    '4175': 'Servicios',
    '4205': 'Financieros',
    '4210': 'Dividendos',
    '4245': 'Recuperaciones',
    '4295': 'Otros ingresos',
    # Costos
    '6135': 'Comercio al por mayor y menor',
    '6175': 'Servicios',
    # Gastos
    '5105': 'Gastos de personal',
    '5110': 'Honorarios',
    '5115': 'Impuestos',
    '5120': 'Arrendamientos',
    '5125': 'Contribuciones y afiliaciones',
    '5130': 'Seguros',
    '5135': 'Servicios',
    '5140': 'Gastos legales',
    '5145': 'Mantenimiento y reparaciones',
    '5150': 'Adecuación e instalación',
    '5155': 'Gastos de viaje',
    '5160': 'Depreciaciones',
    '5165': 'Amortizaciones',
    '5195': 'Diversos',
    '5305': 'Financieros',
    '5315': 'Gastos extraordinarios',
    '5905': 'Impuesto de renta',
    # Cuentas de orden
    '8105': 'Bienes y valores entregados en garantía',
    '9105': 'Bienes y valores recibidos en garantía',
}

# ── Grupos de estados financieros ─────────────────────────────────────────────
GRUPOS_ESF = {
    'activo_corriente':     ['11', '12', '13', '14', '19'],
    'activo_no_corriente':  ['15', '16', '17', '18'],
    'pasivo_corriente':     ['21', '22', '23', '24', '25'],
    'pasivo_no_corriente':  ['26', '27', '28', '29'],
    'patrimonio':           ['31', '32', '33', '34', '35', '36', '37'],
    'ingresos':             ['41', '42', '43', '44'],
    'costos':               ['61', '62', '63', '64'],
    'gastos_operacion':     ['51', '52'],
    'gastos_no_operacion':  ['53', '54', '55'],
    'impuesto_renta':       ['59'],
}

NOTAS_NIIF = {
    'activo_corriente': {
        'titulo': 'Nota 1 — Activos corrientes',
        'politica': (
            'Los activos corrientes se miden al costo o valor razonable según '
            'la naturaleza del activo, de conformidad con la NIIF para las PYMES '
            'Sección 7 (Efectivo) y Sección 11 (Instrumentos financieros básicos). '
            'Las cuentas por cobrar se reconocen al costo amortizado menos deterioro.'
        ),
    },
    'bancos': {
        'titulo': 'Nota 2 — Efectivo y equivalentes de efectivo',
        'politica': (
            'Comprende el efectivo disponible en caja y saldos en cuentas bancarias. '
            'Se presentan al valor nominal. Las diferencias entre el saldo contable y '
            'el extracto bancario se registran como partidas conciliatorias hasta su '
            'aclaración. NIIF PYMES Sección 7.'
        ),
    },
    'clientes': {
        'titulo': 'Nota 3 — Cuentas por cobrar comerciales',
        'politica': (
            'Se reconocen inicialmente al precio de la transacción. '
            'Se evalúa deterioro individualmente para saldos mayores a 3 meses. '
            'La provisión se calcula por el método simplificado de pérdida esperada. '
            'NIIF PYMES Sección 11.'
        ),
    },
    'iva': {
        'titulo': 'Nota 4 — IVA por pagar / IVA descontable',
        'politica': (
            'El IVA generado en ventas se acumula en la cuenta 240801 y el IVA '
            'descontable de compras en 240802. La compensación se realiza '
            'bimestralmente (o cuatrimestral para régimen ordinario PYME). '
            'El saldo neto es la obligación a declarar.'
        ),
    },
    'retenciones': {
        'titulo': 'Nota 5 — Retenciones y anticipos de impuestos',
        'politica': (
            'Las retenciones practicadas a terceros (cuentas 2365, 2367, 2368) '
            'se declaran mensualmente. Las retenciones que nos practican '
            '(cuentas 1355xx) se cruzan contra el impuesto de renta anual. '
            'Se concilia mensualmente con los certificados recibidos.'
        ),
    },
    'nomina': {
        'titulo': 'Nota 6 — Obligaciones laborales y provisiones',
        'politica': (
            'Las obligaciones con empleados se reconocen mensualmente. '
            'Las provisiones de prima, cesantías, intereses sobre cesantías y '
            'vacaciones se acumulan mes a mes al 1/12 del salario con auxilio. '
            'NIIF PYMES Sección 28.'
        ),
    },
    'propiedad_planta': {
        'titulo': 'Nota 7 — Propiedades, planta y equipo',
        'politica': (
            'Se miden al costo menos depreciación acumulada y pérdidas por deterioro. '
            'Depreciación por línea recta sobre vida útil fiscal: equipos de cómputo '
            '5 años, maquinaria 10 años, vehículos 5 años, muebles 10 años. '
            'NIIF PYMES Sección 17.'
        ),
    },
    'ingresos': {
        'titulo': 'Nota 8 — Ingresos de actividades ordinarias',
        'politica': (
            'Los ingresos se reconocen cuando el control del bien o servicio se '
            'transfiere al cliente, al precio de la transacción neto de descuentos. '
            'NIIF PYMES Sección 23.'
        ),
    },
    'impuesto_renta': {
        'titulo': 'Nota 9 — Impuesto sobre la renta',
        'politica': (
            'La provisión de impuesto de renta se estima sobre la utilidad fiscal '
            'a la tarifa vigente del 35% (persona jurídica) o tabla progresiva '
            '(persona natural). Se registra mensualmente el 1/12 de la provisión anual.'
        ),
    },
}


class CierreContable:
    """
    Motor de cierre contable mensual.
    Recibe los datos extraídos de Siigo y los subidos en Excel,
    genera causaciones faltantes y los estados financieros.
    """

    def __init__(self, empresa: dict, periodo: str):
        """
        empresa: dict con datos de la empresa (nit, nombre, regimen, etc.)
        periodo: 'YYYY-MM' ej. '2024-03'
        """
        self.empresa  = empresa
        self.periodo  = periodo
        año, mes      = periodo.split('-')
        self.año      = int(año)
        self.mes      = int(mes)
        self.nombre_mes = ['', 'Enero','Febrero','Marzo','Abril','Mayo','Junio',
                           'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][self.mes]
        ultimo_dia    = calendar.monthrange(self.año, self.mes)[1]
        self.fecha_inicio = date(self.año, self.mes, 1)
        self.fecha_fin    = date(self.año, self.mes, ultimo_dia)
        self.fecha_str    = self.fecha_fin.strftime('%Y-%m-%d')

        # Datos que se irán llenando
        self.saldos: dict        = {}   # cuenta → saldo
        self.movimientos: list   = []   # lista de movimientos del período
        self.partidas_banco: list = []  # partidas conciliatorias
        self.causaciones: list   = []   # causaciones generadas
        self.alertas: list       = []
        self.checklist: list     = []

    # ── Carga de datos desde Siigo ────────────────────────────────────────────

    def cargar_saldos_siigo(self, saldos_raw: list):
        """
        saldos_raw: lista de dicts {cuenta, nombre, debito, credito, saldo}
        tal como los devuelve Siigo en el balance de prueba.
        """
        for s in saldos_raw:
            cuenta = str(s.get('account_code') or s.get('cuenta', '')).strip()
            if not cuenta:
                continue
            # Siigo puede dar débito/crédito separados o saldo neto
            debito  = Decimal(str(s.get('debit') or s.get('debito') or 0))
            credito = Decimal(str(s.get('credit') or s.get('credito') or 0))
            saldo   = Decimal(str(s.get('balance') or s.get('saldo') or 0))
            if saldo == 0:
                saldo = debito - credito

            self.saldos[cuenta] = {
                'nombre':  s.get('account_name') or s.get('nombre') or PUC_CIERRE.get(cuenta[:4], cuenta),
                'debito':  float(debito),
                'credito': float(credito),
                'saldo':   float(saldo),
            }

    def cargar_movimientos_siigo(self, movs_raw: list):
        """movs_raw: lista de movimientos contables del período desde Siigo."""
        self.movimientos = movs_raw

    # ── Carga de Excel del banco ──────────────────────────────────────────────

    def cargar_extracto_banco(self, filas_excel: list, cuenta_banco: str = '1110'):
        """
        filas_excel: resultado de excel_processor.procesar_excel() adaptado para banco.
        Cada fila: {fecha, descripcion, debito, credito, saldo, referencia}
        Genera la conciliación bancaria.
        """
        saldo_banco_final = Decimal('0')
        saldo_contable    = Decimal(str(self.saldos.get(cuenta_banco, {}).get('saldo', 0)))
        partidas          = []

        for mov in filas_excel:
            if not mov.get('fecha'):
                continue
            debito  = Decimal(str(mov.get('debito') or mov.get('credito_banco') or 0))
            credito = Decimal(str(mov.get('credito') or mov.get('debito_banco') or 0))
            saldo_banco_final = Decimal(str(mov.get('saldo') or 0))

            # Buscar si ya está contabilizado
            ref = str(mov.get('referencia') or mov.get('descripcion') or '')
            ya_contabilizado = any(
                ref.lower() in str(m.get('description', '')).lower()
                for m in self.movimientos
            )

            if not ya_contabilizado:
                partidas.append({
                    'fecha':       mov.get('fecha'),
                    'descripcion': mov.get('descripcion') or mov.get('description', ''),
                    'debito':      float(debito),
                    'credito':     float(credito),
                    'tipo':        'EN BANCO / NO EN LIBROS',
                    'accion':      'Registrar en contabilidad',
                })

        diferencia = float(saldo_banco_final) - float(saldo_contable)
        self.partidas_banco = partidas

        self.checklist.append({
            'item':    'Conciliación bancaria',
            'cuenta':  cuenta_banco,
            'estado':  '✅ OK' if abs(diferencia) < 100 else f'⚠️ Diferencia: ${diferencia:,.0f}',
            'detalle': f'Saldo banco: ${float(saldo_banco_final):,.0f} | Saldo libros: ${float(saldo_contable):,.0f}',
            'partidas_conciliatorias': len(partidas),
        })

        return {
            'saldo_banco':    float(saldo_banco_final),
            'saldo_contable': float(saldo_contable),
            'diferencia':     diferencia,
            'partidas':       partidas,
        }

    # ── Causaciones automáticas ───────────────────────────────────────────────

    def generar_causaciones_cierre(self) -> list:
        """
        Genera todas las causaciones de cierre necesarias:
        1. Provisión nómina (si hay empleados pero no hay comprobante de nómina)
        2. Provisión prestaciones sociales
        3. Depreciación del mes
        4. Provisión impuesto de renta
        5. Provisión ICA
        6. Ajuste IVA (compensación descontable vs generado)
        7. Causación intereses financieros
        """
        causaciones = []

        # 1. Provisión prestaciones sociales
        prov = self._provision_prestaciones()
        if prov:
            causaciones.append(prov)

        # 2. Depreciación del mes
        dep = self._depreciacion_mes()
        if dep:
            causaciones.append(dep)

        # 3. Provisión impuesto de renta (1/12 mensual)
        prov_renta = self._provision_renta()
        if prov_renta:
            causaciones.append(prov_renta)

        # 4. Ajuste IVA — compensar descontable vs generado
        ajuste_iva = self._ajuste_iva()
        if ajuste_iva:
            causaciones.append(ajuste_iva)

        # 5. Partidas bancarias no contabilizadas
        for pb in self.partidas_banco[:5]:  # máximo 5 en automático
            if pb['debito'] > 0:
                causaciones.append({
                    'tipo':        'BANCO_NO_CONTABILIZADO',
                    'descripcion': pb['descripcion'],
                    'fecha':       pb['fecha'],
                    'movimientos': [
                        {'cuenta': '5195', 'db': pb['debito'], 'cr': 0,
                         'descripcion': pb['descripcion']},
                        {'cuenta': '1110', 'db': 0, 'cr': pb['debito'],
                         'descripcion': pb['descripcion']},
                    ],
                    'nota': 'Movimiento en extracto bancario no registrado en libros',
                })

        # Normalizar: agregar valor_total como alias de total en cada causación
        for c in causaciones:
            if 'valor_total' not in c:
                c['valor_total'] = c.get('total', sum(
                    m.get('db', 0) for m in c.get('movimientos', [])
                ))
            # Asegurar que 'concepto' exista (alias de descripcion)
            if 'concepto' not in c:
                c['concepto'] = c.get('descripcion', c.get('tipo', ''))

        self.causaciones = causaciones
        return causaciones

    def _provision_prestaciones(self) -> Optional[dict]:
        """Provisión mensual: prima 1/12, cesantías 1/12, int. cesantías 1%, vacaciones 1/24."""
        gasto_nomina = sum(
            v['saldo'] for k, v in self.saldos.items()
            if k.startswith('5105') and v['saldo'] > 0
        )
        if gasto_nomina <= 0:
            return None

        prima      = round(gasto_nomina / 12)
        cesantias  = round(gasto_nomina / 12)
        int_ces    = round(cesantias * 0.01)
        vacaciones = round(gasto_nomina / 24)
        total      = prima + cesantias + int_ces + vacaciones

        return {
            'tipo':        'PROVISION_PRESTACIONES',
            'descripcion': f'Provisión prestaciones sociales {self.nombre_mes} {self.año}',
            'fecha':       self.fecha_str,
            'movimientos': [
                {'cuenta': '5105', 'db': total, 'cr': 0,
                 'descripcion': 'Provisión prestaciones sociales'},
                {'cuenta': '261005', 'db': 0, 'cr': prima,
                 'descripcion': f'Provisión prima — {self.nombre_mes}'},
                {'cuenta': '261010', 'db': 0, 'cr': cesantias,
                 'descripcion': f'Provisión cesantías — {self.nombre_mes}'},
                {'cuenta': '261015', 'db': 0, 'cr': int_ces,
                 'descripcion': f'Intereses sobre cesantías — {self.nombre_mes}'},
                {'cuenta': '261020', 'db': 0, 'cr': vacaciones,
                 'descripcion': f'Provisión vacaciones — {self.nombre_mes}'},
            ],
            'total': total,
            'nota': f'Prima={prima:,.0f} | Cesantías={cesantias:,.0f} | Int={int_ces:,.0f} | Vac={vacaciones:,.0f}',
        }

    def _depreciacion_mes(self) -> Optional[dict]:
        """Depreciación mensual de propiedades planta y equipo."""
        TASAS = {
            '1520': (10, '5160'),  # Maquinaria: 10 años
            '1524': (10, '5160'),  # Equipo oficina: 10 años
            '1528': (5,  '5160'),  # Cómputo: 5 años
            '1540': (5,  '5160'),  # Flota: 5 años
            '1516': (20, '5160'),  # Edificios: 20 años
        }
        movs = []
        total = 0

        for cuenta, (años, gasto_cuenta) in TASAS.items():
            saldo_bruto = sum(
                v['saldo'] for k, v in self.saldos.items()
                if k.startswith(cuenta) and v['saldo'] > 0
            )
            if saldo_bruto <= 0:
                continue

            dep_mes = round(saldo_bruto / (años * 12))
            if dep_mes <= 0:
                continue

            total += dep_mes
            movs.append({'cuenta': gasto_cuenta, 'db': dep_mes, 'cr': 0,
                         'descripcion': f'Depreciación {PUC_CIERRE.get(cuenta, cuenta)}'})
            movs.append({'cuenta': '159240', 'db': 0, 'cr': dep_mes,
                         'descripcion': f'Depreciación acum. {PUC_CIERRE.get(cuenta, cuenta)}'})

        if not movs:
            return None

        return {
            'tipo':        'DEPRECIACION',
            'descripcion': f'Depreciación propiedades planta y equipo {self.nombre_mes} {self.año}',
            'fecha':       self.fecha_str,
            'movimientos': movs,
            'total':       total,
            'nota':        'Calculada por línea recta sobre vida útil fiscal',
        }

    def _provision_renta(self) -> Optional[dict]:
        """Provisión mensual de impuesto de renta (1/12 de la estimación anual)."""
        ingresos = sum(
            abs(v['saldo']) for k, v in self.saldos.items()
            if k.startswith(('41', '42', '43', '44'))
        )
        costos_gastos = sum(
            abs(v['saldo']) for k, v in self.saldos.items()
            if k.startswith(('51', '52', '53', '54', '61', '62', '63'))
        )
        utilidad_gravable = ingresos - costos_gastos
        if utilidad_gravable <= 0:
            return None

        tarifa = 0.35  # 35% personas jurídicas
        impuesto_estimado_año = utilidad_gravable * 12 * tarifa  # proyectar al año
        provision_mes = round(impuesto_estimado_año / 12)

        return {
            'tipo':        'PROVISION_RENTA',
            'descripcion': f'Provisión impuesto de renta {self.nombre_mes} {self.año}',
            'fecha':       self.fecha_str,
            'movimientos': [
                {'cuenta': '5905', 'db': provision_mes, 'cr': 0,
                 'descripcion': 'Provisión impuesto de renta'},
                {'cuenta': '261510', 'db': 0, 'cr': provision_mes,
                 'descripcion': 'Provisión impuesto de renta por pagar'},
            ],
            'total': provision_mes,
            'nota': f'Tarifa 35% sobre utilidad proyectada. Revisar con el contador.',
        }

    def _ajuste_iva(self) -> Optional[dict]:
        """Compensación IVA generado vs descontable."""
        iva_generado   = abs(sum(v['saldo'] for k, v in self.saldos.items()
                                 if k.startswith('240801') and v['saldo'] < 0))
        iva_descontable = sum(v['saldo'] for k, v in self.saldos.items()
                              if k.startswith('240802') and v['saldo'] > 0)

        neto = iva_generado - iva_descontable
        if abs(neto) < 1000:
            return None

        self.checklist.append({
            'item':    'IVA del período',
            'estado':  '📋 Pendiente declaración' if neto > 0 else '💚 Saldo a favor',
            'detalle': f'Generado: ${iva_generado:,.0f} | Descontable: ${iva_descontable:,.0f} | Neto: ${abs(neto):,.0f}',
        })
        return None  # El ajuste de IVA no genera asiento — solo se informa

    # ── Estados financieros ───────────────────────────────────────────────────

    def generar_estados_financieros(self) -> dict:
        """
        Genera el ESF (Balance), ERI (P&G) y flujo de efectivo básico.
        Con notas NIIF para PYMES.
        """
        # Generar causaciones si aún no se han generado
        if not self.causaciones:
            self.generar_causaciones_cierre()

        esf = self._balance_general()
        eri = self._estado_resultados()
        checklist = self._checklist_cierre()
        retenciones = self._resumen_retenciones()

        return {
            'empresa':       self.empresa,
            'periodo':       self.periodo,
            'fecha_corte':   self.fecha_str,
            'esf':           esf,
            'eri':           eri,
            'retenciones':   retenciones,
            'checklist':     checklist + self.checklist,
            'causaciones':   self.causaciones,
            'notas_niif':    self._generar_notas(esf, eri),
            'alertas':       self.alertas,
        }

    def _balance_general(self) -> dict:
        def grupo(prefijos, es_activo=True):
            """
            es_activo=True  → saldos deudores  (cuentas 1x): positivos = bien
            es_activo=False → saldos acreedores (cuentas 2x-3x): en Siigo
                              llegan negativos; los presentamos en positivo.
            """
            items = []
            total = 0
            for k, v in sorted(self.saldos.items()):
                for p in prefijos:
                    if k.startswith(p):
                        saldo = v['saldo'] if es_activo else abs(v['saldo'])
                        items.append({
                            'cuenta': k,
                            'nombre': v['nombre'],
                            'saldo':  saldo,
                        })
                        total += saldo
                        break
            return {'items': items, 'total': total}

        ac  = grupo(['11', '12', '13', '14'], es_activo=True)
        anc = grupo(['15', '16', '17', '18'], es_activo=True)
        pc  = grupo(['21', '22', '23', '24', '25'], es_activo=False)
        pnc = grupo(['26', '27', '28', '29'],       es_activo=False)
        pat = grupo(['31', '32', '33', '34', '35', '36'], es_activo=False)

        total_activo  = ac['total'] + anc['total']
        total_pasivo  = pc['total'] + pnc['total']
        total_pat     = pat['total']
        ecuacion_ok   = abs(total_activo - (total_pasivo + total_pat)) < 1000

        if not ecuacion_ok:
            self.alertas.append(
                f'⚠️ ECUACIÓN PATRIMONIAL DESCUADRADA: Activo={total_activo:,.0f} '
                f'vs Pasivo+Patrimonio={total_pasivo + total_pat:,.0f}'
            )

        return {
            'activo_corriente':     ac,
            'activo_no_corriente':  anc,
            'total_activo':         total_activo,
            'pasivo_corriente':     pc,
            'pasivo_no_corriente':  pnc,
            'total_pasivo':         total_pasivo,
            'patrimonio':           pat,
            'total_patrimonio':     total_pat,
            'total_pasivo_pat':     total_pasivo + total_pat,
            'ecuacion_ok':          ecuacion_ok,
        }

    def _estado_resultados(self) -> dict:
        def suma(prefijos):
            items = []
            total = 0
            for k, v in sorted(self.saldos.items()):
                for p in prefijos:
                    if k.startswith(p):
                        items.append({
                            'cuenta': k,
                            'nombre': v['nombre'],
                            'valor':  abs(v['saldo']),
                        })
                        total += abs(v['saldo'])
                        break
            return {'items': items, 'total': total}

        ingresos_op  = suma(['41', '42'])
        ingresos_no  = suma(['43', '44'])
        costos        = suma(['61', '62', '63'])
        gastos_op     = suma(['51', '52'])
        gastos_no_op  = suma(['53', '54', '55'])
        imp_renta     = suma(['59'])

        utilidad_bruta  = ingresos_op['total'] - costos['total']
        uaii            = utilidad_bruta - gastos_op['total']
        utilidad_antes  = uaii + ingresos_no['total'] - gastos_no_op['total']
        utilidad_neta   = utilidad_antes - imp_renta['total']

        # EBITDA: UAII + depreciaciones y amortizaciones
        dep_amort = sum(abs(v['saldo']) for k, v in self.saldos.items()
                        if k.startswith(('5160', '5165')))
        ebitda = uaii + dep_amort

        return {
            'ingresos_operacionales':     ingresos_op,
            'ingresos_no_operacionales':  ingresos_no,
            'costo_ventas':               costos,
            'gastos_operacionales':       gastos_op,
            'gastos_no_operacionales':    gastos_no_op,
            'impuesto_renta':             imp_renta,
            'utilidad_bruta':             utilidad_bruta,
            'uaii':                       uaii,
            'utilidad_antes_impuestos':   utilidad_antes,
            'utilidad_neta':              utilidad_neta,
            'ebitda':                     ebitda,
            'margen_bruto':               round(utilidad_bruta / ingresos_op['total'] * 100, 2) if ingresos_op['total'] else 0,
            'margen_neto':                round(utilidad_neta / ingresos_op['total'] * 100, 2) if ingresos_op['total'] else 0,
        }

    def _resumen_retenciones(self) -> dict:
        """Consolida las retenciones a declarar y las retenciones a favor."""
        retefuente_pagar  = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('2365'))
        reteiva_pagar     = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('2367'))
        retica_pagar      = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('2368'))
        ret_nomina_pagar  = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('2370'))
        retefuente_favor  = sum(v['saldo'] for k, v in self.saldos.items() if k.startswith('1355'))
        iva_generado      = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('240801'))
        iva_descontable   = sum(v['saldo'] for k, v in self.saldos.items() if k.startswith('240802'))

        self.checklist.append({
            'item':    'ReteFuente a declarar (Form. 350)',
            'estado':  '📋 Verificar' if retefuente_pagar > 0 else '✅ Sin obligación',
            'detalle': f'${retefuente_pagar:,.0f}',
        })
        self.checklist.append({
            'item':    'ReteIVA a declarar',
            'estado':  '📋 Verificar' if reteiva_pagar > 0 else '✅ Sin obligación',
            'detalle': f'${reteiva_pagar:,.0f}',
        })
        self.checklist.append({
            'item':    'ReteICA a declarar',
            'estado':  '📋 Verificar' if retica_pagar > 0 else '✅ Sin obligación',
            'detalle': f'${retica_pagar:,.0f}',
        })

        return {
            'retefuente_pagar':   retefuente_pagar,
            'reteiva_pagar':      reteiva_pagar,
            'retica_pagar':       retica_pagar,
            'ret_nomina_pagar':   ret_nomina_pagar,
            'total_a_declarar':   retefuente_pagar + reteiva_pagar + retica_pagar + ret_nomina_pagar,
            'retefuente_a_favor': retefuente_favor,
            'iva_generado':       iva_generado,
            'iva_descontable':    iva_descontable,
            'iva_neto':           iva_generado - iva_descontable,
        }

    def _checklist_cierre(self) -> list:
        items = []

        # Verificar saldo bancos > 0
        saldo_bancos = sum(v['saldo'] for k, v in self.saldos.items() if k.startswith('11'))
        items.append({
            'item':   'Efectivo y bancos',
            'estado': '✅ OK' if saldo_bancos >= 0 else '⚠️ Saldo negativo',
            'detalle': f'${saldo_bancos:,.0f}',
        })

        # Clientes sin vencer
        saldo_clientes = sum(v['saldo'] for k, v in self.saldos.items() if k.startswith('13'))
        items.append({
            'item':   'Cartera clientes',
            'estado': '📋 Revisar antigüedad' if saldo_clientes > 0 else '✅ Sin cartera',
            'detalle': f'${saldo_clientes:,.0f}',
        })

        # Proveedores
        saldo_prov = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('22'))
        items.append({
            'item':   'Obligaciones con proveedores',
            'estado': '📋 Verificar vencimientos' if saldo_prov > 0 else '✅ Al día',
            'detalle': f'${saldo_prov:,.0f}',
        })

        # Nómina
        prov_laborales = sum(abs(v['saldo']) for k, v in self.saldos.items() if k.startswith('261'))
        items.append({
            'item':   'Provisiones laborales',
            'estado': '✅ Provisionado' if prov_laborales > 0 else '⚠️ Sin provisión',
            'detalle': f'${prov_laborales:,.0f}',
        })

        return items

    def _generar_notas(self, esf: dict, eri: dict) -> list:
        notas = []
        for clave, nota in NOTAS_NIIF.items():
            notas.append({
                'titulo':   nota['titulo'],
                'politica': nota['politica'],
                'periodo':  f'{self.nombre_mes} {self.año}',
            })
        return notas
