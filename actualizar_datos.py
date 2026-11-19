#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Monitor de Cashflows ONs - actualización con una sola base.

Uso normal: py actualizar_datos.py
Validar sin tocar data.js: py actualizar_datos.py --validar
Generar sin API (solo excepcional): py actualizar_datos.py --sin-api

Formato salida compatible con el HTML existente: window `const jsonData` en data.js.
Las cotizaciones en tiempo real siguen viniendo de los endpoints Data912 originales.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import requests
import urllib3

# Mantenemos la configuración SSL de la implementación anterior para no alterar
# la conectividad de Data912 en los entornos donde el certificado falla.
# Si se solucionó en el entorno: DATA912_SSL_VERIFY=1 activa la verificación.
VERIFY_SSL = os.getenv('DATA912_SSL_VERIFY', '0').strip() == '1'
if not VERIFY_SSL:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

API_URLS = (
    'https://data912.com/live/arg_corp',
    'https://data912.com/live/arg_bonds',
)
# MAPEO ORIGINAL PRESERVADO: algunos símbolos de mercado difieren del ticker base.
MAPEO_ESPECIAL = {
    'BPO27': ['BPJ27D', 'BPO27D', 'BPO2D', 'BPI27D'],
    'BPO28': ['BPJ28D', 'BPO28D'],
    'BPOA7': ['BPA7D', 'BPOA7D'],
    'BPOB7': ['BPB7D', 'BPOB7D'],
    'BPOC7': ['BPC7D', 'BPOC7D'],
    'BPOD7': ['BPD7D', 'BPOD7D'],
    'BPO30': ['BPJ30D', 'BPO30D', 'BPY26D'],
    'AO29': ['AO29D', 'AO29'],
    'AL29': ['AL29D'], 'AL30': ['AL30D'],
    'AL35': ['AL35D'], 'AE38': ['AE38D'], 'GD29': ['GD29D'],
    'GD30': ['GD30D'], 'GD35': ['GD35D'], 'GD38': ['GD38D'],
    'GD41': ['GD41D'], 'GD46': ['GD46D'],
}

RENOMBRES = {
    'Fecha Cupón': 'Fecha de pago',
    'VR Cartera': 'VN residual previo',
    'Renta Efect.': 'Interes c/100 VN',
    'Amortización % c/100 VN': 'Amortizacion c/100 VN',
    'Total': 'Flujo total c/100 VN',
}
FLUJO_COLUMNAS = ['Ticker', 'Fecha Cupón', 'VR Cartera', 'Renta Efect.',
                  'Amortización % c/100 VN', 'Obs. Prox. Pago', 'Total']
UNIVERSO_REQUERIDOS = ['Ticker', 'Ley', 'Emisor', 'Tipo de instrumento']


def texto(x: Any) -> str:
    if x is None or pd.isna(x):
        return ''
    return str(x).strip()


def ticker(x: Any) -> str:
    return texto(x).upper()


def numero(x: Any, defecto=None):
    if x is None or pd.isna(x):
        return defecto
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return float(x)
    s = str(x).strip().replace(' ', '').replace('\u00a0', '')
    if not s or s.upper() in {'N/D', '-', 'NAN', 'NONE'}:
        return defecto
    s = s.replace('%', '')
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return defecto


def fecha_excel(v: Any):
    """Fecha segura: datetime, serial Excel, ISO, dd/mm/yyyy, DD-Mon-YY.

    Nunca invierte día y mes de un texto dd/mm/yyyy.
    """
    if v is None or pd.isna(v) or str(v).strip() == '':
        return None
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 <= v <= 100000:
            return (pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(v))).normalize()
        return None
    s = str(v).strip()
    # Versiones de Excel exportadas como serial en texto
    if re.fullmatch(r'\d{5}(?:\.0+)?', s):
        return fecha_excel(float(s))
    for f in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y', '%d/%m/%y',
              '%d-%m-%Y', '%d-%m-%y', '%d-%b-%y', '%d-%b-%Y',
              '%d/%b/%Y', '%d/%b/%y'):
        try:
            return pd.Timestamp(datetime.strptime(s, f)).normalize()
        except ValueError:
            pass
    d = pd.to_datetime(s, errors='coerce', dayfirst=True)
    return d.normalize() if pd.notna(d) else None


def tasa_porcentaje(valor: Any) -> str:
    if isinstance(valor, str) and '%' in valor:
        return valor.strip()
    x = numero(valor)
    if x is None:
        return 'N/D'
    return f'{x*100 if abs(x)<1 else x:.2f}%'


def paridad_porcentaje(valor: Any) -> str:
    if isinstance(valor, str) and ('%' in valor or 'USD' in valor.upper()):
        return valor.strip()
    x = numero(valor)
    if x is None:
        return 'N/D'
    return f'{x*100 if abs(x)<2 else x:.2f}%'


def dm_texto(valor: Any) -> str:
    x = numero(valor)
    return f'{x:.2f}' if x is not None else 'N/D'


def calcular_tir(precio, flujos_futuros):
    """Misma matemática del procesador original; solo explicita la convergencia."""
    if not flujos_futuros or precio is None or precio <= 0 or precio > 200:
        return None
    low, high = -0.99, 10.0
    try:
        for _ in range(100):
            mid = (low + high) / 2
            vp = sum(monto / ((1 + mid) ** (dias / 365.0))
                     for dias, monto in flujos_futuros)
            if vp > precio:
                low = mid
            else:
                high = mid
        return mid
    except (OverflowError, ZeroDivisionError, ValueError):
        return None


def calcular_dm(tir, precio, flujos_futuros):
    if not flujos_futuros or precio is None or precio <= 0 or precio > 200 or tir is None:
        return None
    try:
        macaulay = sum((dias / 365.0) * monto / ((1 + tir) ** (dias / 365.0))
                       for dias, monto in flujos_futuros) / precio
        return macaulay / (1 + tir)
    except (OverflowError, ZeroDivisionError, ValueError):
        return None


def obtener_precios_usd_data912(tickers_universo, getter=requests.get):
    """Conserva endpoints, campo 'c'/bid, mapeos y denominación USD de Data912."""
    precios = {}
    respuesta_ok = 0
    errores = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
        'Accept': 'application/json',
    }
    raw_items = []
    for url in API_URLS:
        try:
            print(f'Consultando Data912: {url}')
            res = getter(url, headers=headers, timeout=15, verify=VERIFY_SSL)
            res.raise_for_status()
            items = res.json()
            if not isinstance(items, list):
                raise ValueError('La API no devolvió una lista')
            raw_items.extend(items)
            respuesta_ok += 1
            print(f'  OK: {len(items)} cotizaciones disponibles')
        except Exception as exc:
            msg = f'{url}: {str(exc)[:180]}'
            errores.append(msg)
            print(f'  AVISO: {msg}')

    inverso = {var.upper(): base for base, variantes in MAPEO_ESPECIAL.items()
               for var in variantes}
    for item in raw_items:
        if not isinstance(item, dict):
            continue
        tk_data = ticker(item.get('symbol'))
        precio = numero(item.get('c'))
        if precio is None or precio <= 0:
            precio = numero(item.get('px_bid'))
        if not precio or not 0 < precio <= 200:
            continue  # el mismo resguardo anti-precios en pesos del script anterior
        base = inverso.get(tk_data)
        if base and base in tickers_universo:
            precios[base] = precio
            continue
        if tk_data.endswith('D'):
            if tk_data in tickers_universo:
                precios[tk_data] = precio
            elif tk_data[:-1] + 'O' in tickers_universo:
                precios[tk_data[:-1] + 'O'] = precio
            elif tk_data[:-1] in tickers_universo:
                precios[tk_data[:-1]] = precio
    return precios, respuesta_ok, errores


def ubicaciones(base: Path):
    unificado = base / 'Base_Cashflows.xlsx'
    if unificado.is_file():
        return unificado, unificado, True
    flujos = base / 'flujos.xlsx'
    universo = base / 'cashflows_ON_USD_ley_ARG_NY_universo_solo_USD_2.xlsx'
    if flujos.is_file() and universo.is_file():
        print('Modo compatible: se encontraron los 2 Excel antiguos.')
        return flujos, universo, False
    raise FileNotFoundError(
        'No se encontró Base_Cashflows.xlsx ni el par de archivos originales. '
        'Colocalos con actualizar_datos.py.'
    )


def leer_datos(base: Path):
    ruta_f, ruta_u, combinado = ubicaciones(base)
    df_f = pd.read_excel(ruta_f, sheet_name='Flujos' if combinado else 0)
    df_u = pd.read_excel(ruta_u, sheet_name='Universo' if combinado else 'Universo')
    for key in FLUJO_COLUMNAS:
        if key not in df_f.columns:
            raise ValueError(f'Falta columna de Flujos: {key}')
    for key in UNIVERSO_REQUERIDOS:
        if key not in df_u.columns:
            raise ValueError(f'Falta columna de Universo: {key}')
    df_f = df_f.copy()
    df_u = df_u.copy()
    df_f['Ticker'] = df_f['Ticker'].map(ticker)
    df_u['Ticker'] = df_u['Ticker'].map(ticker)
    df_f = df_f[df_f['Ticker'] != ''].copy()
    df_u = df_u[df_u['Ticker'] != ''].copy()
    return df_f, df_u, combinado


def validar_flujos(df_f):
    validos = []
    fechas_invalidas = []
    importes_invalidos = []
    inconsistencias = []
    duplicados = []
    claves = Counter()
    for num, row in enumerate(df_f.to_dict('records'), start=2):
        tk = ticker(row.get('Ticker'))
        f = fecha_excel(row.get('Fecha Cupón'))
        if f is None:
            fechas_invalidas.append((tk, num, texto(row.get('Fecha Cupón'))))
            continue
        conv = {}
        for col in ('VR Cartera', 'Renta Efect.', 'Amortización % c/100 VN', 'Total'):
            x = numero(row.get(col))
            if x is None:
                importes_invalidos.append((tk, num, col, texto(row.get(col))))
            conv[col] = x
        if any(v is None for v in conv.values()):
            continue
        if abs(conv['Renta Efect.'] + conv['Amortización % c/100 VN'] - conv['Total']) > 0.06:
            inconsistencias.append((tk, num, f.strftime('%d/%m/%Y')))
        claves[tk, f.strftime('%Y-%m-%d')] += 1
        validos.append({
            'Ticker': tk,
            'Fecha de pago': f.strftime('%d/%m/%Y'),
            'VN residual previo': conv['VR Cartera'],
            'Interes c/100 VN': conv['Renta Efect.'],
            'Amortizacion c/100 VN': conv['Amortización % c/100 VN'],
            'Obs. Prox. Pago': texto(row.get('Obs. Prox. Pago')),
            'Flujo total c/100 VN': conv['Total'],
            'Fecha_ISO': f.strftime('%Y-%m-%d'),
            'Año_Mes': f.strftime('%Y-%m'),
            'Mes_Num': f.month,
            '_fecha': f,
        })
    duplicados = [(k[0], k[1], v) for k,v in claves.items() if v > 1]
    return validos, fechas_invalidas, importes_invalidos, inconsistencias, duplicados


def cargar_universo(df_u):
    uni = {}
    duplicados = []
    excluidos = set()
    monedas_bloqueadas = set()
    por_verificar = set()
    for row in df_u.to_dict('records'):
        tk = ticker(row.get('Ticker'))
        include = texto(row.get('Incluir Dashboard', 'SI')).upper()
        if tk in uni or tk in excluidos:
            duplicados.append(tk)
            continue  # PRIMER registro gana: no pisar la ficha por un duplicado posterior
        if include not in {'SI', 'SÍ', 'S', '1', 'TRUE', 'YES', 'X', ''}:
            excluidos.add(tk)
            continue
        if 'Moneda Pago' in df_u.columns:
            moneda = texto(row.get('Moneda Pago')).upper()
            # Reincorporaciones antiguas sin validar: excepcionalmente no borrarlas
            # de golpe; permanecen visibles y figuran en el reporte a auditar.
            recuperado = texto(row.get('Origen ficha')).startswith('Recuperado de data.js')
            if moneda not in {'USD', 'US$', 'DOLARES', 'DÓLARES'}:
                if recuperado:
                    por_verificar.add(tk)
                else:
                    excluidos.add(tk)
                    monedas_bloqueadas.add(tk)
                    continue
        uni[tk] = row
    return uni, duplicados, excluidos, monedas_bloqueadas, por_verificar


def construir(base: Path, validar: bool, sin_api: bool, permite_sin_api: bool):
    print('Cargando Excel maestro...')
    df_f, df_u, combinado = leer_datos(base)
    catalogo, uni_duplicados, excluidos, monedas_bloqueadas, por_verificar = cargar_universo(df_u)
    flujos, fechas_malas, numeros_malos, totales_distintos, fechas_duplicadas = validar_flujos(df_f)
    por_ticker = defaultdict(list)
    for f in flujos:
        por_ticker[f['Ticker']].append(f)
    pendientes = sorted(set(por_ticker) - set(catalogo) - excluidos)
    sin_flujo = sorted(set(catalogo) - set(por_ticker))

    print(f'  Base: {"unificada" if combinado else "legacy 2 Excel"}')
    print(f'  Flujos válidos: {len(flujos)} / {len(df_f)}')
    print(f'  Activos incluidos: {len(catalogo)} / con pagos: {len(set(catalogo)&set(por_ticker))}')
    print(f'  Activos de Flujos pendientes de catalogar: {len(pendientes)}')
    print(f'  Activos incluidos sin Flujos: {len(sin_flujo)}')
    print(f'  Duplicados de ticker en Universo: {len(uni_duplicados)}')
    print(f'  Moneda pendiente de revisar en catálogo anterior: {len(por_verificar)}')
    print(f'  Fechas inválidas: {len(fechas_malas)}; importes inválidos: {len(numeros_malos)}')
    print(f'  Fechas repetidas ticker/día: {len(fechas_duplicadas)} (sin eliminar)')

    reporte = []
    for tk in pendientes:
        reporte.append([tk, 'PENDIENTE', len(por_ticker[tk]), 'Existe en Flujos pero falta ficha en Universo (no se publica)'])
    for tk in sin_flujo:
        reporte.append([tk, 'SIN FLUJOS', 0, 'Ficha incluida sin cronograma de pagos'])
    for tk in sorted(monedas_bloqueadas):
        reporte.append([tk, 'MONEDA NO CONFIRMADA', len(por_ticker.get(tk, [])),
                        'No se publica: confirmar pago en USD en Universo'])
    for tk in sorted(por_verificar):
        reporte.append([tk, 'FICHA LEGACY A AUDITAR', len(por_ticker.get(tk, [])),
                        'Se conservó del dashboard previo; revisar ley, moneda y emisor'])
    for tk in uni_duplicados:
        reporte.append([tk, 'UNIVERSO DUPLICADO', len(por_ticker.get(tk, [])), 'Se usa la primera ficha; revisar la duplicada'])
    for tk, f, n in fechas_duplicadas:
        reporte.append([tk, 'FECHA REPETIDA', n, f'{f}: se mantienen las {n} filas, sin borrar nada'])
    for tk, r, val in fechas_malas:
        reporte.append([tk, 'FECHA INVALIDA', 0, f'Fila Excel {r}: {val}'])
    for tk, r, col, val in numeros_malos:
        reporte.append([tk, 'NUMERO INVALIDO', 0, f'Fila Excel {r} / {col}: {val}'])
    for tk, r, fecha in totales_distintos:
        reporte.append([tk, 'TOTAL A REVISAR', 1, f'Fila Excel {r} / {fecha}: total no coincide con renta+amortización'])
    with (base / 'reporte_actualizacion.csv').open('w', encoding='utf-8-sig', newline='') as fh:
        writer = csv.writer(fh, delimiter=';')
        writer.writerow(['Ticker', 'Estado', 'Filas', 'Detalle'])
        writer.writerows(reporte)

    if fechas_malas or numeros_malos:
        raise ValueError('Hay fechas o importes ilegibles en Flujos; ver reporte_actualizacion.csv. Se protegió data.js.')
    if validar:
        print('Control completado; data.js y data.json no fueron modificados.')
        return {'pendientes': len(pendientes), 'sin_flujos': len(sin_flujo)}

    # Todo el procesamiento de flujos es independiente de la cotización actual:
    # que falle Data912 no debe borrar cronogramas ni inventar precios.
    precios, endpoints_ok, errores = ({}, 0, []) if sin_api else obtener_precios_usd_data912(set(catalogo))
    if not sin_api and (not endpoints_ok or len(precios) == 0) and not permite_sin_api:
        raise RuntimeError('Data912 no entregó precios utilizables: data.js anterior preservado. '
                           'Usar --sin-api solo si necesitás generar la base sin cotizaciones actuales.')

    hoy = pd.Timestamp(datetime.now().date())
    flujos_futuros = defaultdict(list)
    for f in flujos:
        if f['Ticker'] not in catalogo:
            continue
        dias = (f['_fecha'] - hoy).days
        if dias >= 0 and f['Flujo total c/100 VN'] > 0:
            flujos_futuros[f['Ticker']].append((dias, f['Flujo total c/100 VN']))

    universo_dict = {}
    for tk, row in catalogo.items():
        venc = fecha_excel(row.get('Fecha Vencimiento'))
        if venc is None and por_ticker.get(tk):
            venc = max(f['_fecha'] for f in por_ticker[tk])
        cupon = numero(row.get('Tasa Cupon Vigente %'))
        # No convertir automáticamente pagos parciales del año en una tasa contractual.
        if cupon is None:
            cupon = 0.0  # sin dato: mantener compatibilidad de tipos con el HTML
        ley = texto(row.get('Ley')) or 'Sin verificar'
        emisor = texto(row.get('Emisor')) or 'Emisor a verificar'
        tipo = texto(row.get('Tipo de instrumento')) or 'Sin clasificar'
        ytm = tasa_porcentaje(row.get('YTM'))
        dm = dm_texto(row.get('DM'))
        paridad = paridad_porcentaje(row.get('Paridad'))
        if tk in precios and tk in flujos_futuros:
            precio = precios[tk]
            tir_math = calcular_tir(precio, flujos_futuros[tk])
            dm_math = calcular_dm(tir_math, precio, flujos_futuros[tk])
            if tir_math is not None:
                ytm = f'{tir_math * 100:.2f}%'
            if dm_math is not None:
                dm = f'{dm_math:.2f}'
            # El endpoint original entrega cotizaciones USD; preservamos escala.
            paridad = f'{precio:.2f}%'
        universo_dict[tk] = {
            'Cupon': float(cupon),
            'Vencimiento': venc.strftime('%d/%m/%Y') if venc is not None else '',
            'VencAño': int(venc.year) if venc is not None else None,
            'Emisor': emisor,
            'Ley': ley,
            'Tipo_Inst': tipo,
            'YTM': ytm,
            'DM': dm,
            'Paridad': paridad,
        }

    # Mantenemos íntegros los importes, incluso duplicados históricos a auditar.
    cashflows = []
    for f in flujos:
        if f['Ticker'] not in universo_dict:
            continue
        item = {k: v for k, v in f.items() if k != '_fecha'}
        item['Ley'] = universo_dict[f['Ticker']]['Ley']
        cashflows.append(item)
    cashflows.sort(key=lambda f: (f['Ticker'], f['Fecha_ISO']))

    if sin_api:
        fuente = 'Referencias del Excel (sin consulta a Data912)'
        api_estado = 'sin_consulta'
    elif endpoints_ok == 2:
        fuente = 'Data912 API (precios actuales)'
        api_estado = 'ok'
    elif endpoints_ok == 1:
        fuente = 'Data912 API (respuesta parcial: 1 de 2 endpoints)'
        api_estado = 'parcial'
    else:
        fuente = 'Referencias del Excel; Data912 no respondió'
        api_estado = 'sin_datos'

    datos = {
        'metadata': {
            'ultima_actualizacion': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'total_activos': len(universo_dict),
            'fuente': fuente,
            'api_estado': api_estado,
            'precios_api': len(precios),
            'endpoints_api_ok': endpoints_ok,
            'activos_con_flujos': len(set(catalogo) & set(por_ticker)),
            'total_flujos': len(cashflows),
            'activos_pendientes_revision': len(pendientes),
            'activos_sin_flujos': len(sin_flujo),
        },
        'cashflows': cashflows,
        'universo': universo_dict,
    }
    # Validar serialización antes de escribir. Usar archivos temporales evita
    # dejar data.js truncado si ocurre un error al generar JSON.
    contenido = json.dumps(datos, ensure_ascii=False, separators=(',', ':'), allow_nan=False)
    js = 'const jsonData = ' + contenido + ';\n'
    temp_js = base / 'data.js.tmp'
    temp_json = base / 'data.json.tmp'
    temp_js.write_text(js, encoding='utf-8')
    temp_json.write_text(contenido + '\n', encoding='utf-8')
    os.replace(temp_js, base / 'data.js')
    os.replace(temp_json, base / 'data.json')
    print(f'OK: data.js y data.json actualizados. {len(universo_dict)} activos, {len(cashflows)} pagos.')
    print(f'Data912: {api_estado} / precios: {len(precios)} / reporte_actualizacion.csv generado')
    if errores:
        print('Avisos de Data912: ' + '; '.join(errores))
    return datos['metadata']


def pausar_si_interactivo():
    """Mantiene abierta la consola de Windows al abrir el .py con doble clic.

    No bloquea ejecuciones automáticas, GitHub Actions, tareas sin consola,
    ni procesos que importen este módulo (p. ej. prueba_compatibilidad.py).
    """
    if not sys.stdin or not sys.stdin.isatty():
        return
    if os.environ.get('CI', '').lower() in ('1', 'true', 'yes'):
        return
    if os.environ.get('GITHUB_ACTIONS', '').lower() == 'true':
        return
    try:
        input('\nPresioná ENTER para cerrar esta ventana...')
    except (EOFError, KeyboardInterrupt):
        pass


def main(argv=None):
    parser = argparse.ArgumentParser(description='Actualizar dashboard Cashflows sin perder la API Data912')
    parser.add_argument('--validar', action='store_true', help='Controlar Excel y generar reporte, sin tocar los datos JS')
    parser.add_argument('--sin-api', action='store_true', help='Generar sin precios API: solo para contingencias')
    parser.add_argument('--permitir-sin-api', action='store_true', help='No bloquear si los 2 endpoints fallan')
    parser.add_argument('--directorio', type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument('--no-pausar', action='store_true', help='No esperar ENTER; útil para .bat y automatizaciones')
    args = parser.parse_args(argv)
    codigo = 0
    try:
        resumen = construir(args.directorio.resolve(), args.validar, args.sin_api, args.permitir_sin_api)
        print('\n' + '=' * 58)
        if args.validar:
            print('CONTROL COMPLETADO CORRECTAMENTE - no se modificaron los datos JS')
            print(f'Pendientes de catalogar: {resumen.get("pendientes", 0)}')
            print(f'Activos incluidos sin flujos: {resumen.get("sin_flujos", 0)}')
        else:
            print('ACTUALIZACION FINALIZADA CORRECTAMENTE')
            print(f'Activos en dashboard: {resumen["total_activos"]}')
            print(f'Activos con pagos: {resumen["activos_con_flujos"]}')
            print(f'Registros de flujos: {resumen["total_flujos"]}')
            print(f'Data912: {resumen["api_estado"]} - {resumen["precios_api"]} precios')
            print(f'Endpoints con respuesta: {resumen["endpoints_api_ok"]} de 2')
            print(f'Ultima actualizacion: {resumen["ultima_actualizacion"]}')
            print('Archivos generados: data.js, data.json y reporte_actualizacion.csv')
        print('=' * 58)
    except Exception as exc:
        codigo = 1
        print('\n' + '=' * 58, file=sys.stderr)
        print(f'ERROR: {exc}', file=sys.stderr)
        print('No se pudo completar la actualizacion. Revisá los mensajes anteriores.', file=sys.stderr)
        print('=' * 58, file=sys.stderr)
    finally:
        if not args.no_pausar:
            pausar_si_interactivo()
    return codigo


if __name__ == '__main__':
    sys.exit(main())
