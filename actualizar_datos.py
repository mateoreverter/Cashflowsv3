import os
import pandas as pd
import json
import traceback
import requests
import urllib3
from datetime import datetime, timedelta

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# MATEMÁTICA FINANCIERA (CALCULADORA DE BONOS)
# =========================================================
def calcular_tir(precio, flujos_futuros):
    if not flujos_futuros or precio <= 0 or precio > 200:
        return None
    low, high = -0.99, 10.0
    for _ in range(100):
        mid = (low + high) / 2
        vp = sum([monto / ((1 + mid) ** (dias / 365.0)) for dias, monto in flujos_futuros])
        if vp > precio: low = mid
        else: high = mid
    return mid

def calcular_dm(tir, precio, flujos_futuros):
    if not flujos_futuros or precio <= 0 or precio > 200 or tir is None:
        return None
    macaulay = sum([(dias/365.0) * monto / ((1 + tir) ** (dias / 365.0)) for dias, monto in flujos_futuros]) / precio
    dm = macaulay / (1 + tir)
    return dm

# =========================================================
# CONEXIÓN A DATA912 CON MAPEO MÚLTIPLE
# =========================================================
def obtener_precios_usd_data912(tickers_universo):
    print("\nConectando a Data912 para descargar precios en vivo...")
    precios_usd = {}
    
    # Diccionario revertido: Solo Bopreales
    MAPEO_ESPECIAL = {
        "BPO27": ["BPJ27D", "BPO27D", "BPO2D", "BPI27D"], 
        "BPO28": ["BPJ28D", "BPO28D"],
        "BPOA7": ["BPA7D", "BPOA7D"],
        "BPOB7": ["BPB7D", "BPOB7D"],
        "BPOC7": ["BPC7D", "BPOC7D"],
        "BPOD7": ["BPD7D", "BPOD7D"],
        "BPO30": ["BPJ30D", "BPO30D", "BPY26D"]
    }
    
    urls_api = [
        "https://data912.com/live/arg_corp",
        "https://data912.com/live/arg_bonds"
    ]
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json"
    }
    
    raw_items = []
    for url in urls_api:
        try:
            print(f" -> Consultando {url}...")
            response = requests.get(url, headers=headers, timeout=15, verify=False)
            if response.status_code == 200:
                raw_items.extend(response.json())
        except Exception as e:
            print(f"   [!] Falló la conexión a {url}: {e}")
            
    for item in raw_items:
        ticker_data = str(item.get('symbol', '')).strip().upper()
        
        precio = item.get('c', 0)
        if precio == 0:
            precio = item.get('px_bid', 0)
            
        if precio > 0:
            if precio > 200:
                continue
                
            match_especial = False
            for base_key, variantes in MAPEO_ESPECIAL.items():
                if ticker_data in variantes:
                    precios_usd[base_key] = float(precio)
                    match_especial = True
                    break
                    
            if match_especial:
                continue
            
            if ticker_data.endswith('D'):
                base_ticker = None
                if ticker_data in tickers_universo:
                    base_ticker = ticker_data
                elif ticker_data[:-1] + 'O' in tickers_universo: 
                    base_ticker = ticker_data[:-1] + 'O'
                elif ticker_data[:-1] in tickers_universo: 
                    base_ticker = ticker_data[:-1]
                
                if base_ticker:
                    precios_usd[base_ticker] = float(precio)
                        
    print(f"✓ ¡Éxito! Se rescataron precios limpios en USD para {len(precios_usd)} activos.\n")
    return precios_usd

def formatear_fecha(fecha_dt):
    meses = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}
    if pd.isnull(fecha_dt): return ''
    return f"{fecha_dt.strftime('%d')}-{meses[fecha_dt.month]}-{fecha_dt.strftime('%y')}"

# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================
def main():
    directorio_base = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    
    archivo_flujos = os.path.join(directorio_base, 'flujos.xlsx')
    archivo_universo = os.path.join(directorio_base, 'cashflows_ON_USD_ley_ARG_NY_universo_solo_USD_2.xlsx')
    archivo_salida_js = os.path.join(directorio_base, 'data.js')

    print("Cargando archivos Excel...")
    df_flujos = pd.read_excel(archivo_flujos)
    
    xls = pd.ExcelFile(archivo_universo)
    hoja_universo = 'Universo' if 'Universo' in xls.sheet_names else xls.sheet_names[0]
    df_universo = pd.read_excel(archivo_universo, sheet_name=hoja_universo)

    tickers_universo_set = set(df_universo['Ticker'].dropna().astype(str).str.strip().str.upper())
    
    print("Mapeando flujos futuros para calcular TIR...")
    hoy = datetime.now()
    flujos_futuros_dict = {}
    
    df_flujos['Fecha_DT_temp'] = pd.to_datetime(df_flujos['Fecha Cupón'], errors='coerce')
    
    for _, row in df_flujos.iterrows():
        if pd.isna(row['Ticker']) or pd.isna(row['Fecha_DT_temp']): continue
        
        tk = str(row['Ticker']).strip().upper()
        fecha_pago = row['Fecha_DT_temp']
        monto = float(row.get('Total', 0.0)) if pd.notnull(row.get('Total')) else 0.0
        
        if fecha_pago >= hoy and monto > 0:
            dias = (fecha_pago - hoy).days
            if tk not in flujos_futuros_dict: flujos_futuros_dict[tk] = []
            flujos_futuros_dict[tk].append((dias, monto))

    flujos_año_actual = df_flujos[df_flujos['Fecha_DT_temp'].dt.year == hoy.year]
    cupones_calculados = flujos_año_actual.groupby('Ticker')['Renta Efect.'].sum().to_dict()
    vencimientos_calculados = df_flujos.groupby('Ticker')['Fecha_DT_temp'].max().to_dict()

    precios_usd = obtener_precios_usd_data912(tickers_universo_set.union(set(df_flujos['Ticker'].dropna().str.strip().str.upper())))

    print("Calculando Métricas y armando Universo...")
    universo_dict = {}
    mapa_leyes = {}

    def procesar_activo(tk, row=None):
        if row is not None:
            venc_str = str(row.get('Fecha Vencimiento', '')) if pd.notnull(row.get('Fecha Vencimiento')) else ''
            emisor_str = str(row.get('Emisor', '')) if pd.notnull(row.get('Emisor')) else ''
            ley_str = str(row.get('Ley', '')) if pd.notnull(row.get('Ley')) else 'Ley ARG'
            tipo_inst = str(row.get('Tipo de instrumento', 'ON')).strip() if pd.notnull(row.get('Tipo de instrumento')) else 'ON'
            cupon_excel = row.get('Tasa Cupon Vigente %', 0.0)
            
            ytm_excel = str(row.get('YTM', 'N/D')) if pd.notnull(row.get('YTM')) else 'N/D'
            dm_excel = str(row.get('DM', 'N/D')) if pd.notnull(row.get('DM')) else 'N/D'
            paridad_excel = str(row.get('Paridad', 'N/D')) if pd.notnull(row.get('Paridad')) else 'N/D'
        else:
            fecha_venc_dt = vencimientos_calculados.get(tk)
            venc_str = formatear_fecha(fecha_venc_dt)
            emisor_str = 'Otros'
            ley_str = 'Ley ARG'
            tipo_inst = 'Otros'
            cupon_excel = 0.0
            ytm_excel, dm_excel, paridad_excel = 'N/D', 'N/D', 'N/D'

        if pd.isna(cupon_excel) or cupon_excel == 0.0:
            cupon_excel = cupones_calculados.get(tk, 0.0)
            
        mapa_leyes[tk] = ley_str
        
        venc_year = None
        try:
            venc_dt = pd.to_datetime(venc_str, errors='coerce')
            if pd.notnull(venc_dt): venc_year = int(venc_dt.year)
        except: pass

        if tk in precios_usd and tk in flujos_futuros_dict:
            precio_hoy = precios_usd[tk]
            flujos = flujos_futuros_dict[tk]
            
            tir_math = calcular_tir(precio_hoy, flujos)
            dm_math = calcular_dm(tir_math, precio_hoy, flujos)
            
            ytm_val = f"{tir_math*100:.2f}%" if tir_math is not None else ytm_excel
            dm_val = f"{dm_math:.2f}" if dm_math is not None else dm_excel
            paridad_val = f"{precio_hoy:.2f}%" 
            
        else:
            ytm_val = ytm_excel
            dm_val = dm_excel
            paridad_val = paridad_excel

            if ytm_val != 'N/D' and '%' not in ytm_val:
                try: ytm_val = f"{float(ytm_val)*100:.2f}%" if float(ytm_val) < 1 else f"{float(ytm_val):.2f}%"
                except: pass
            if paridad_val != 'N/D' and '%' not in paridad_val and 'USD' not in paridad_val:
                try: paridad_val = f"{float(paridad_val)*100:.2f}%" if float(paridad_val) < 2 else f"{float(paridad_val):.2f}%"
                except: pass

        universo_dict[tk] = {
            'Cupon': float(cupon_excel),
            'Vencimiento': venc_str if venc_str != 'nan' else '',
            'VencAño': venc_year,
            'Emisor': emisor_str if emisor_str != 'nan' else '',
            'Ley': ley_str,
            'Tipo_Inst': tipo_inst,
            'YTM': ytm_val,
            'DM': dm_val,
            'Paridad': paridad_val
        }

    for _, row in df_universo.iterrows():
        if pd.notna(row.get('Ticker')):
            procesar_activo(str(row['Ticker']).strip().upper(), row)

    for tk in df_flujos['Ticker'].dropna().unique():
        tk = str(tk).strip().upper()
        if tk not in universo_dict:
            procesar_activo(tk)

    print("\nEjecutando limpieza de bonos en pesos y basura...")
    prefijos_basura = ('TX', 'TC', 'T2', 'T3', 'T4', 'T5', 'T6', 'DI', 'PR', 'PA', 'TV', 'TO', 'CU', 'NO')
    tickers_a_eliminar = []
    
    for tk, data in universo_dict.items():
        if tk.startswith(prefijos_basura):
            tickers_a_eliminar.append(tk)
            continue
            
        if data['Emisor'] == 'Otros' and data['Paridad'] == 'N/D':
            tickers_a_eliminar.append(tk)

    for tk in tickers_a_eliminar:
        if tk in universo_dict:
            del universo_dict[tk]
            
    print(f"-> Se eliminaron {len(tickers_a_eliminar)} instrumentos que no corresponden a la curva Hard Dollar.\n")

    print("Formateando hoja de flujos...")
    df_flujos = df_flujos.drop(columns=['Fecha_DT_temp'], errors='ignore')

    rename_map = {
        'Fecha Cupón': 'Fecha de pago',
        'VR Cartera': 'VN residual previo',
        'Renta Efect.': 'Interes c/100 VN',
        'Amortización % c/100 VN': 'Amortizacion c/100 VN',
        'Total': 'Flujo total c/100 VN'
    }
    df_cf_clean = df_flujos.rename(columns=rename_map)
    df_cf_clean = df_cf_clean.dropna(subset=['Ticker'])
    df_cf_clean['Ticker'] = df_cf_clean['Ticker'].astype(str).str.strip().str.upper()
    
    df_cf_clean = df_cf_clean[df_cf_clean['Ticker'].isin(universo_dict.keys())]
    
    if 'Ley' not in df_cf_clean.columns:
        df_cf_clean['Ley'] = df_cf_clean['Ticker'].map(mapa_leyes).fillna('Ley ARG')

    if 'Fecha de pago' in df_cf_clean.columns:
        df_cf_clean['Fecha_DT'] = pd.to_datetime(df_cf_clean['Fecha de pago'], errors='coerce')
        df_cf_clean['Fecha de pago'] = df_cf_clean['Fecha_DT'].dt.strftime('%d/%m/%Y')
        df_cf_clean['Fecha_ISO'] = df_cf_clean['Fecha_DT'].dt.strftime('%Y-%m-%d')
        df_cf_clean['Año_Mes'] = df_cf_clean['Fecha_DT'].dt.strftime('%Y-%m')
        df_cf_clean['Mes_Num'] = df_cf_clean['Fecha_DT'].dt.month
        df_cf_clean = df_cf_clean.drop(columns=['Fecha_DT'])

    df_cf_clean = df_cf_clean.fillna('')

    hora_arg = datetime.utcnow() - timedelta(hours=3)
    metadata = {
        "ultima_actualizacion": hora_arg.strftime("%d/%m/%Y %H:%M:%S"),
        "total_activos": len(universo_dict),
        "fuente": "Data912 API (vía GitHub Actions)"
    }

    output_data = {
        "metadata": metadata,
        "cashflows": df_cf_clean.to_dict(orient='records'),
        "universo": universo_dict
    }

    print(f"Guardando data.js en: {archivo_salida_js}...")
    with open(archivo_salida_js, "w", encoding="utf-8") as f:
        f.write("const jsonData = ")
        json.dump(output_data, f, ensure_ascii=False, indent=2)
        f.write(";")

    print("¡Éxito total! Archivo JS generado listo para usar de forma local.")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("\n--- Ocurrió un ERROR ---")
        traceback.print_exc()