#%pip install openpyxl
import pandas as pd
import openpyxl
import numpy as np

# Cargar todas las hojas del Excel
file = "/Volumes/workspace/default/lab8-accidentes/datos.xlsx"
xls = pd.ExcelFile(file)

import pandas as pd
import numpy as np

# ============================================================================
# CONFIGURACIÓN DE RANGOS POR DIMENSIÓN
# ============================================================================

RANGOS_CONFIG = {
    # Departamento (22)
    ("Departamento", "Anio"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 6},
    ("Departamento", "Mes"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 13},
    ("Departamento", "Weekday"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 8},
    ("Departamento", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 10},
    ("Departamento", "Tipo_Vehiculo"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 20},
    ("Departamento", "Hora"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 25},
    ("Departamento", "Zona"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 27},
    ("Departamento", "Sexo"): {"fila_inicio": 1, "fila_fin": 23, "col_inicio": 0, "col_fin": 6},
    
    # Mes (12)
    ("Mes", "Anio"): {"fila_inicio": 1, "fila_fin": 13, "col_inicio": 0, "col_fin": 6},
    ("Mes", "Weekday"): {"fila_inicio": 1, "fila_fin": 13, "col_inicio": 0, "col_fin": 8},
    ("Mes", "Departamento"): {"fila_inicio": 1, "fila_fin": 13, "col_inicio": 0, "col_fin": 23},
    ("Mes", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 13, "col_inicio": 0, "col_fin": 10},
    
    # Año (5)
    ("Anio", "Departamento"): {"fila_inicio": 1, "fila_fin": 6, "col_inicio": 0, "col_fin": 23},
    ("Anio", "Mes"): {"fila_inicio": 1, "fila_fin": 6, "col_inicio": 0, "col_fin": 13},
    ("Anio", "Weekday"): {"fila_inicio": 1, "fila_fin": 6, "col_inicio": 0, "col_fin": 8},
    
    # Weekday (7)
    ("Weekday", "Departamento"): {"fila_inicio": 1, "fila_fin": 8, "col_inicio": 0, "col_fin": 23},
    ("Weekday", "Mes"): {"fila_inicio": 1, "fila_fin": 8, "col_inicio": 0, "col_fin": 13},
    ("Weekday", "Anio"): {"fila_inicio": 1, "fila_fin": 8, "col_inicio": 0, "col_fin": 6},
    ("Weekday", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 8, "col_inicio": 0, "col_fin": 10},
    ("Weekday", "Sexo"): {"fila_inicio": 1, "fila_fin": 8, "col_inicio": 0, "col_fin": 3},
    
    # Hora (24)
    ("Hora", "Weekday"): {"fila_inicio": 1, "fila_fin": 25, "col_inicio": 0, "col_fin": 8},
    ("Hora", "Mes"): {"fila_inicio": 1, "fila_fin": 25, "col_inicio": 0, "col_fin": 13},
    ("Hora", "Anio"): {"fila_inicio": 1, "fila_fin": 25, "col_inicio": 0, "col_fin": 6},
    ("Hora", "Departamento"): {"fila_inicio": 1, "fila_fin": 25, "col_inicio": 0, "col_fin": 23},
    ("Hora", "Hora"): {"fila_inicio": 1, "fila_fin": 25, "col_inicio": 0, "col_fin": 25},
    ("Hora", "Zona"): {"fila_inicio": 1, "fila_fin": 25, "col_inicio": 0, "col_fin": 27},
    
    # Zona (26: 1-25 + Ignorada)
    ("Zona", "Hora"): {"fila_inicio": 1, "fila_fin": 27, "col_inicio": 0, "col_fin": 25},
    ("Zona", "Weekday"): {"fila_inicio": 1, "fila_fin": 27, "col_inicio": 0, "col_fin": 8},
    ("Zona", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 27, "col_inicio": 0, "col_fin": 10},
    ("Zona", "Tipo_Vehiculo"): {"fila_inicio": 1, "fila_fin": 27, "col_inicio": 0, "col_fin": 20},
    ("Zona", "Sexo"): {"fila_inicio": 1, "fila_fin": 27, "col_inicio": 0, "col_fin": 3},
    
    # Día del mes (31)
    ("Dia_Mes", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 32, "col_inicio": 0, "col_fin": 10},
    
    # Tipo de Accidente (9)
    ("Tipo_Accidente", "Departamento"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 23},
    ("Tipo_Accidente", "Tipo_Vehiculo"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 20},
    ("Tipo_Accidente", "Color_Vehiculo"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 17},
    ("Tipo_Accidente", "Modelo_Vehiculo"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 8},
    ("Tipo_Accidente", "Mes"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 13},
    ("Tipo_Accidente", "Weekday"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 8},
    ("Tipo_Accidente", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 10},
    ("Tipo_Accidente", "Sexo"): {"fila_inicio": 1, "fila_fin": 10, "col_inicio": 0, "col_fin": 3},
    
    # Tipo de Vehículo
    ("Tipo_Vehiculo", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 20, "col_inicio": 0, "col_fin": 10},
    ("Tipo_Vehiculo", "Hora"): {"fila_inicio": 1, "fila_fin": 20, "col_inicio": 0, "col_fin": 25},
    ("Tipo_Vehiculo", "Sexo"): {"fila_inicio": 1, "fila_fin": 20, "col_inicio": 0, "col_fin": 3},
    
    # Color Vehículo
    ("Color_Vehiculo", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 17, "col_inicio": 0, "col_fin": 10},
    
    # Edad (16 rangos)
    ("Edad", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 17, "col_inicio": 0, "col_fin": 10},
    ("Edad", "Weekday"): {"fila_inicio": 1, "fila_fin": 17, "col_inicio": 0, "col_fin": 8},
    ("Edad", "Hora"): {"fila_inicio": 1, "fila_fin": 17, "col_inicio": 0, "col_fin": 25},
    ("Edad", "Sexo"): {"fila_inicio": 1, "fila_fin": 17, "col_inicio": 0, "col_fin": 3},
    
    # Modelo Vehículo (6 rangos)
    ("Modelo_Vehiculo", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 7, "col_inicio": 0, "col_fin": 10},
    
    # Sexo (2)
    ("Sexo", "Weekday"): {"fila_inicio": 1, "fila_fin": 3, "col_inicio": 0, "col_fin": 8},
    ("Sexo", "Tipo_Accidente"): {"fila_inicio": 1, "fila_fin": 3, "col_inicio": 0, "col_fin": 10},
    ("Sexo", "Departamento"): {"fila_inicio": 1, "fila_fin": 3, "col_inicio": 0, "col_fin": 23},
}

# ============================================================================
# MAPEO DE MÉTRICAS POR CUADRO
# ============================================================================

METRICAS_POR_CUADRO = {
    **{i: "accidentes_totales" for i in range(1, 17)},
    **{i: "vehiculos_involucrados_accidentes" for i in range(17, 29)},
    **{i: "victimas_totales" for i in range(29, 31)},
    **{i: "lesionados_totales" for i in range(31, 47)},
    **{i: "fallecidos_totales" for i in range(47, 63)},
    63: "tasa_victimas_involucradas",
    64: "tasa_lesionados",
    65: "tasa_fallecidos",
}

# Hojas estratificadas
HOJAS_COMPLEJAS = []
HOJAS_ESTRATIFICADAS = [37, 53, 54]

# ============================================================================
# FUNCIONES DE DETECCIÓN
# ============================================================================

def detectar_dimension_fila(df_preview, debug=False):
    """Detecta qué dimensión está en las filas"""
    
    primer_valor = None
    for i in range(len(df_preview)):
        val = str(df_preview.iloc[i, 0]).strip().lower()
        if val not in ['total', 'nan', ''] and not pd.isna(df_preview.iloc[i, 0]):
            primer_valor = val
            break
    
    if debug and primer_valor:
        print(f"    Primer valor real: '{primer_valor}'")
    
    departamentos = ['guatemala', 'alta verapaz', 'baja verapaz', 'chimaltenango', 'chiquimula', 
                     'peten', 'petén', 'el progreso', 'quiche', 'quiché', 'escuintla', 'huehuetenango',
                     'izabal', 'jalapa', 'jutiapa', 'quetzaltenango', 'retalhuleu', 'sacatepequez',
                     'sacatepéquez', 'san marcos', 'santa rosa', 'solola', 'sololá', 'suchitepequez',
                     'suchitepéquez', 'totonicapan', 'totonicapán', 'zacapa']
    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
             'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    dias = ['lunes', 'martes', 'miercoles', 'miércoles', 'jueves', 'viernes', 
            'sabado', 'sábado', 'domingo']
    tipos_accidente = ['colisión', 'colision', 'atropello', 'derrape', 'choque', 'vuelco', 
                       'embarrancó', 'embarranco', 'encunetó', 'encuneto', 'caída', 'caida']
    tipos_vehiculo = ['motocicleta', 'automóvil', 'automovil', 'pick up', 'camioneta', 
                      'camión', 'camion', 'cabezal', 'bus', 'microbús', 'microbus']
    colores = ['negro', 'blanco', 'rojo', 'gris', 'azul', 'verde', 'corinto', 
               'amarillo', 'beige', 'café', 'cafe', 'anaranjado']
    sexos_validos = ['hombre', 'mujer', 'ignorado']
    modelos_vehiculo = ['1970', '1980', '1990', '2000', '2010', '2020']
    
    if primer_valor:
        if any(dept in primer_valor for dept in departamentos):
            return "Departamento"
        elif any(mes in primer_valor for mes in meses):
            return "Mes"
        elif any(dia in primer_valor for dia in dias):
            return "Weekday"
        elif ':' in primer_valor and 'a' in primer_valor:
            return "Hora"
        elif any(tipo in primer_valor for tipo in tipos_accidente):
            return "Tipo_Accidente"
        elif any(tipo in primer_valor for tipo in tipos_vehiculo):
            return "Tipo_Vehiculo"
        elif any(color in primer_valor for color in colores):
            return "Color_Vehiculo"
        elif any(año in primer_valor for año in ['2019', '2020', '2021', '2022', '2023', '2024']):
            return "Anio"
        elif primer_valor in sexos_validos:
            return "Sexo"
        elif any(modelo in primer_valor for modelo in modelos_vehiculo):
            return "Modelo_Vehiculo"
        elif 'menor de' in primer_valor or ('años' in primer_valor and '-' in primer_valor):
            return "Edad"
        elif primer_valor.isdigit() or (len(primer_valor) <= 3 and primer_valor[0].isdigit()):
            num = int(''.join(filter(str.isdigit, primer_valor))) if primer_valor[0].isdigit() else 0
            if 1 <= num <= 25:
                return "Zona"
            elif 1 <= num <= 31:
                return "Dia_Mes"
    
    valores_unicos = df_preview.iloc[:, 0].dropna()
    valores_unicos = valores_unicos[valores_unicos.astype(str).str.strip().str.lower() != 'total']
    num_valores = len(valores_unicos)
    
    if debug:
        print(f"    Número de valores únicos: {num_valores}")
    
    if num_valores >= 30 and num_valores <= 32:
        return "Dia_Mes"
    elif num_valores >= 25 and num_valores <= 27:
        primer_val = str(valores_unicos.iloc[0]).strip() if len(valores_unicos) > 0 else ""
        return "Hora" if ':' in primer_val else "Zona"
    elif num_valores >= 23 and num_valores <= 25:
        return "Hora"
    elif num_valores >= 20 and num_valores <= 23:
        return "Departamento"
    elif num_valores >= 15 and num_valores <= 17:
        return "Edad"
    elif num_valores >= 11 and num_valores <= 13:
        return "Mes"
    elif num_valores >= 8 and num_valores <= 10:
        return "Tipo_Accidente"
    elif num_valores >= 6 and num_valores <= 8:
        return "Weekday"
    elif num_valores >= 4 and num_valores <= 6:
        return "Anio"
    elif num_valores == 2 or num_valores == 3:
        vals = valores_unicos.str.lower().str.strip().tolist()
        if all(v in sexos_validos for v in vals):
            return "Sexo"
    elif num_valores == 6 or num_valores == 7:
        return "Modelo_Vehiculo"
    
    return None

def detectar_dimension_columna(columnas, debug=False):
    """Detecta qué dimensión está en las columnas"""
    
    valores_str = ' '.join([str(col).lower().strip() for col in columnas[1:15]])
    
    if debug:
        print(f"    String columnas: '{valores_str[:100]}'")
    
    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio']
    dias = ['lunes', 'martes', 'miercoles', 'miércoles', 'jueves', 'viernes']
    tipos_accidente = ['colisión', 'colision', 'atropello', 'derrape', 'choque', 'vuelco']
    tipos_vehiculo = ['motocicleta', 'automóvil', 'automovil', 'pick up', 'camioneta']
    colores = ['negro', 'blanco', 'rojo', 'gris', 'azul', 'verde']
    departamentos = ['guatemala', 'alta verapaz', 'quetzaltenango', 'peten', 'petén']
    sexos = ['hombre', 'mujer', 'ignorado']
    modelos_vehiculo = ['1970', '1980', '1990', '2000', '2010', '2020']
    
    if any(str(año) in valores_str for año in ['2019', '2020', '2021', '2022', '2023', '2024']):
        return "Anio"
    elif any(mes in valores_str for mes in meses):
        return "Mes"
    elif any(dia in valores_str for dia in dias):
        return "Weekday"
    elif any(tipo in valores_str for tipo in tipos_accidente):
        return "Tipo_Accidente"
    elif any(tipo in valores_str for tipo in tipos_vehiculo):
        return "Tipo_Vehiculo"
    elif any(color in valores_str for color in colores):
        return "Color_Vehiculo"
    elif any(dept in valores_str for dept in departamentos):
        return "Departamento"
    elif any(sexo in valores_str for sexo in sexos):
        return "Sexo"
    elif any(modelo in valores_str for modelo in modelos_vehiculo):
        return "Modelo_Vehiculo"
    elif '00:00' in valores_str or '01:00' in valores_str or (':' in valores_str and 'a' in valores_str):
        return "Hora"
    elif '1980' in valores_str or '1990' in valores_str or '2000' in valores_str:
        return "Modelo_Vehiculo"
    
    return None

# ============================================================================
# FUNCIONES PARA HOJAS COMPLEJAS (22, 23, 25, 26)
# ============================================================================

def procesar_hoja_22_23(file, sheet_name, header_row=9, debug=False):
    """Procesa hojas 22 y 23: Departamento × Condicion_Conductor × Sexo"""
    try:
        num_cuadro = int(sheet_name.replace('cuadro', '').strip())
        metrica = METRICAS_POR_CUADRO.get(num_cuadro, "vehiculos_involucrados_accidentes")
        
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, nrows=30)
        
        datos = []
        
        for i in range(1, min(23, len(df))):
            dept = str(df.iloc[i, 0]).strip()
            if not dept or dept.lower() == 'total':
                continue
            
            col_start = None
            for j, col in enumerate(df.columns):
                if 'no ebrio' in str(col).lower():
                    col_start = j
                    break
            
            if col_start is None:
                col_start = 2
            
            # No ebrio
            if col_start + 1 < len(df.columns):
                val_hombre = df.iloc[i, col_start + 1]
                if pd.notna(val_hombre):
                    datos.append({
                        'Departamento': dept,
                        'Condicion_Conductor': 'No ebrio',
                        'Sexo': 'Hombre',
                        metrica: float(val_hombre)
                    })
            
            if col_start + 2 < len(df.columns):
                val_mujer = df.iloc[i, col_start + 2]
                if pd.notna(val_mujer):
                    datos.append({
                        'Departamento': dept,
                        'Condicion_Conductor': 'No ebrio',
                        'Sexo': 'Mujer',
                        metrica: float(val_mujer)
                    })
            
            # Ebrio
            col_ebrio = None
            for j, col in enumerate(df.columns):
                if 'ebrio' in str(col).lower() and 'no ebrio' not in str(col).lower():
                    col_ebrio = j
                    break
            
            if col_ebrio and col_ebrio + 1 < len(df.columns):
                val_hombre = df.iloc[i, col_ebrio + 1]
                if pd.notna(val_hombre):
                    datos.append({
                        'Departamento': dept,
                        'Condicion_Conductor': 'Ebrio',
                        'Sexo': 'Hombre',
                        metrica: float(val_hombre)
                    })
            
            if col_ebrio and col_ebrio + 2 < len(df.columns):
                val_mujer = df.iloc[i, col_ebrio + 2]
                if pd.notna(val_mujer):
                    datos.append({
                        'Departamento': dept,
                        'Condicion_Conductor': 'Ebrio',
                        'Sexo': 'Mujer',
                        metrica: float(val_mujer)
                    })
        
        df_largo = pd.DataFrame(datos)
        
        todas_dims = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                      "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                      "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
        for dim in todas_dims:
            if dim not in df_largo.columns:
                df_largo[dim] = np.nan
        
        df_largo["hoja_origen"] = sheet_name
        return df_largo
        
    except Exception as e:
        raise ValueError(f"Error procesando hoja {num_cuadro}: {str(e)}")


def procesar_hoja_25_26(file, sheet_name, header_row=9, debug=False):
    """Procesa hojas 25 y 26: Hora/Edad × Condicion_Conductor × Sexo"""
    try:
        num_cuadro = int(sheet_name.replace('cuadro', '').strip())
        metrica = METRICAS_POR_CUADRO.get(num_cuadro, "vehiculos_involucrados_accidentes")
        
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, nrows=35)
        
        primer_val = str(df.iloc[1, 0]).strip().lower()
        dim_fila = "Hora" if ':' in primer_val else "Edad"
        max_filas = 25 if dim_fila == "Hora" else 17
        
        datos = []
        
        for i in range(1, min(max_filas, len(df))):
            val_fila = str(df.iloc[i, 0]).strip()
            if not val_fila or val_fila.lower() == 'total':
                continue
            
            col_start = None
            for j, col in enumerate(df.columns):
                if 'no ebrio' in str(col).lower():
                    col_start = j
                    break
            
            if col_start is None:
                col_start = 2
            
            # No ebrio
            if col_start + 1 < len(df.columns):
                val_hombre = df.iloc[i, col_start + 1]
                if pd.notna(val_hombre):
                    datos.append({
                        dim_fila: val_fila,
                        'Condicion_Conductor': 'No ebrio',
                        'Sexo': 'Hombre',
                        metrica: float(val_hombre)
                    })
            
            if col_start + 2 < len(df.columns):
                val_mujer = df.iloc[i, col_start + 2]
                if pd.notna(val_mujer):
                    datos.append({
                        dim_fila: val_fila,
                        'Condicion_Conductor': 'No ebrio',
                        'Sexo': 'Mujer',
                        metrica: float(val_mujer)
                    })
            
            # Ebrio
            col_ebrio = None
            for j, col in enumerate(df.columns):
                if 'ebrio' in str(col).lower() and 'no ebrio' not in str(col).lower():
                    col_ebrio = j
                    break
            
            if col_ebrio and col_ebrio + 1 < len(df.columns):
                val_hombre = df.iloc[i, col_ebrio + 1]
                if pd.notna(val_hombre):
                    datos.append({
                        dim_fila: val_fila,
                        'Condicion_Conductor': 'Ebrio',
                        'Sexo': 'Hombre',
                        metrica: float(val_hombre)
                    })
            
            if col_ebrio and col_ebrio + 2 < len(df.columns):
                val_mujer = df.iloc[i, col_ebrio + 2]
                if pd.notna(val_mujer):
                    datos.append({
                        dim_fila: val_fila,
                        'Condicion_Conductor': 'Ebrio',
                        'Sexo': 'Mujer',
                        metrica: float(val_mujer)
                    })
        
        df_largo = pd.DataFrame(datos)
        
        todas_dims = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                      "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                      "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
        for dim in todas_dims:
            if dim not in df_largo.columns:
                df_largo[dim] = np.nan
        
        df_largo["hoja_origen"] = sheet_name
        return df_largo
        
    except Exception as e:
        raise ValueError(f"Error procesando hoja {num_cuadro}: {str(e)}")

# ============================================================================
# FUNCIONES PARA HOJAS ESPECIALES
# ============================================================================

def procesar_hoja_14(file, sheet_name, header_row=9, debug=False):
    """Procesa hoja 14: Hora × Zona"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, nrows=30)
        
        col_inicio_zonas = None
        for i, col in enumerate(df.columns):
            if str(col).strip().isdigit() and int(str(col).strip()) == 1:
                col_inicio_zonas = i
                break
        
        if col_inicio_zonas is None:
            for i, col in enumerate(df.columns):
                if 'zona' in str(col).lower() or (str(col).strip().isdigit() and 1 <= int(str(col).strip()) <= 25):
                    col_inicio_zonas = i
                    break
        
        if col_inicio_zonas is None:
            col_inicio_zonas = 1
        
        df_data = df.iloc[1:25, col_inicio_zonas-1:col_inicio_zonas+26]
        
        zonas = [f"Zona {i}" for i in range(1, 26)] + ["Ignorada"]
        if len(df_data.columns) == 27:
            df_data.columns = ["Hora"] + zonas
        else:
            available_cols = min(26, len(df_data.columns) - 1)
            zonas = [f"Zona {i}" for i in range(1, available_cols + 1)]
            if len(df_data.columns) > available_cols + 1:
                zonas.append("Ignorada")
            df_data.columns = ["Hora"] + zonas[:len(df_data.columns)-1]
        
        value_vars = [col for col in df_data.columns if col != "Hora"]
        df_largo = df_data.melt(id_vars=["Hora"], value_vars=value_vars, 
                               var_name="Zona", value_name="accidentes_totales")
        
        df_largo = df_largo[df_largo["Hora"].notna()]
        df_largo["accidentes_totales"] = pd.to_numeric(df_largo["accidentes_totales"], errors='coerce')
        
        todas_dims = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                      "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                      "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
        for dim in todas_dims:
            if dim not in df_largo.columns:
                df_largo[dim] = np.nan
        
        df_largo["hoja_origen"] = sheet_name
        return df_largo
        
    except Exception as e:
        raise ValueError(f"Error procesando hoja 14: {str(e)}")

def procesar_hoja_20(file, sheet_name, header_row=9, debug=False):
    """Procesa hoja 20: Modelo_Vehiculo × Tipo_Accidente"""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, nrows=15)
        
        df_data = df.iloc[1:7, 0:11]
        
        tipos_accidente = ['Colisión', 'Atropello', 'Derrape', 'Choque', 'Vuelco', 
                          'Embarrancó', 'Encunetó', 'Caída', 'Otro', 'Total']
        df_data.columns = ["Modelo_Vehiculo"] + tipos_accidente[:-1] + ["Total"]
        
        df_largo = df_data.melt(id_vars=["Modelo_Vehiculo"], value_vars=tipos_accidente[:-1], 
                               var_name="Tipo_Accidente", value_name="vehiculos_involucrados_accidentes")
        
        df_largo = df_largo[df_largo["Modelo_Vehiculo"].notna()]
        df_largo["vehiculos_involucrados_accidentes"] = pd.to_numeric(
            df_largo["vehiculos_involucrados_accidentes"], errors='coerce'
        )
        
        todas_dims = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                      "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                      "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
        for dim in todas_dims:
            if dim not in df_largo.columns:
                df_largo[dim] = np.nan
        
        df_largo["hoja_origen"] = sheet_name
        return df_largo
        
    except Exception as e:
        raise ValueError(f"Error procesando hoja 20: {str(e)}")

def procesar_hoja_estratificada_optimizada(file, sheet_name, header_row=9, debug=False):
    """Procesa hojas estratificadas"""
    try:
        num_cuadro = int(sheet_name.replace('cuadro', '').strip())
        metrica = METRICAS_POR_CUADRO.get(num_cuadro, "valor")
        
        if num_cuadro == 37:
            nrows = 30
        elif num_cuadro == 53:
            nrows = 30  
        elif num_cuadro == 54:
            nrows = 50
        else:
            nrows = 40
            
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, nrows=nrows)
        
        if num_cuadro in [37, 53]:
            datos = []
            current_tipo = ""
            
            for i in range(1, min(28, len(df))):
                valor_fila = str(df.iloc[i, 0]).strip()
                
                if not valor_fila or valor_fila.lower() == 'total':
                    continue
                    
                if any(tipo in valor_fila.lower() for tipo in ['colisión', 'atropello', 'derrape', 'choque', 'vuelco', 'embarrancó', 'encunetó', 'caída', 'otro']):
                    current_tipo = valor_fila
                elif any(sexo in valor_fila.lower() for sexo in ['hombre', 'mujer', 'ignorado']):
                    for j in range(1, 8):
                        if j < len(df.columns):
                            valor_celda = df.iloc[i, j]
                            if pd.notna(valor_celda):
                                try:
                                    datos.append({
                                        'Tipo_Accidente': current_tipo,
                                        'Sexo': valor_fila,
                                        'Weekday': ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'][j-1],
                                        metrica: float(valor_celda)
                                    })
                                except (ValueError, TypeError):
                                    continue
            
            df_largo = pd.DataFrame(datos)
            
        elif num_cuadro == 54:
            datos = []
            current_edad = ""
            
            for i in range(1, min(49, len(df))):
                valor_fila = str(df.iloc[i, 0]).strip()
                
                if not valor_fila or valor_fila.lower() == 'total':
                    continue
                    
                if 'menor' in valor_fila.lower() or 'años' in valor_fila.lower() or '-' in valor_fila:
                    current_edad = valor_fila
                elif any(sexo in valor_fila.lower() for sexo in ['hombre', 'mujer', 'ignorado']):
                    for j in range(1, 8):
                        if j < len(df.columns):
                            valor_celda = df.iloc[i, j]
                            if pd.notna(valor_celda):
                                try:
                                    datos.append({
                                        'Edad': current_edad,
                                        'Sexo': valor_fila,
                                        'Weekday': ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'][j-1],
                                        metrica: float(valor_celda)
                                    })
                                except (ValueError, TypeError):
                                    continue
            
            df_largo = pd.DataFrame(datos)
        
        else:
            raise ValueError(f"Hoja estratificada {num_cuadro} no reconocida")
        
        todas_dims = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                      "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                      "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
        for dim in todas_dims:
            if dim not in df_largo.columns:
                df_largo[dim] = np.nan
        
        df_largo["hoja_origen"] = sheet_name
        return df_largo
        
    except Exception as e:
        raise ValueError(f"Error procesando hoja estratificada {num_cuadro}: {str(e)}")

# ============================================================================
# FUNCIÓN DE LIMPIEZA
# ============================================================================

def limpiar_dataset_maestro(df):
    """Limpia valores inválidos del dataset maestro"""
    
    valores_validos = {
        'Sexo': ['Hombre', 'Mujer', 'Ignorado'],
        'Condicion_Conductor': ['Ebrio', 'No ebrio', 'Ignorado'],
        'Departamento': ['Guatemala', 'El Progreso', 'Sacatepéquez', 'Chimaltenango', 'Escuintla',
                        'Santa Rosa', 'Sololá', 'Totonicapán', 'Quetzaltenango', 'Suchitepéquez',
                        'Retalhuleu', 'San Marcos', 'Huehuetenango', 'Quiché', 'Baja Verapaz',
                        'Alta Verapaz', 'Petén', 'Izabal', 'Zacapa', 'Chiquimula', 'Jalapa', 'Jutiapa'],
        'Mes': ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 
                'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
        'Weekday': ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    }
    
    # for col in ['Sexo', 'Condicion_Conductor', 'Departamento', 'Mes', 'Weekday']:
    #     if col in df.columns:
    #         df[col] = df[col].astype(str).str.strip().str.title()

    for col in ['Sexo', 'Condicion_Conductor', 'Departamento', 'Mes', 'Weekday']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(x).strip().title() if pd.notna(x) else np.nan)
    
    for col, validos in valores_validos.items():
        if col in df.columns:
            #mask = ~df[col].isin(validos + ['Nan', 'None'])
            mask = df[col].notna() & ~df[col].isin(validos)
            if mask.any():
                print(f"  Limpiando {mask.sum()} valores inválidos en '{col}'")
                df.loc[mask, col] = np.nan
    
    dimensiones = ['Departamento', 'Mes', 'Anio', 'Weekday', 'Hora', 'Zona', 
                   'Dia_Mes', 'Tipo_Accidente', 'Tipo_Vehiculo', 'Color_Vehiculo', 
                   'Modelo_Vehiculo', 'Edad', 'Sexo', 'Condicion_Conductor']
    
    mask_todas_nan = df[dimensiones].isna().all(axis=1)
    if mask_todas_nan.any():
        print(f"  Eliminando {mask_todas_nan.sum()} filas sin dimensiones válidas")
        df = df[~mask_todas_nan]
    
    antes = len(df)
    df = df.drop_duplicates()
    if len(df) < antes:
        print(f"  Eliminados {antes - len(df)} duplicados exactos")
    
    return df

# ============================================================================
# FUNCIÓN PRINCIPAL DE TRANSFORMACIÓN
# ============================================================================

def transformar_hoja_auto(file, sheet_name, header_row=9, debug=False):
    """Transforma una hoja detectando automáticamente sus dimensiones"""
    
    try:
        num_cuadro = int(sheet_name.replace('cuadro', '').strip())
    except:
        raise ValueError(f"No se pudo extraer número de cuadro de '{sheet_name}'")
    
    if num_cuadro == 14:
        if debug:
            print(f"  Procesando hoja especial 14 (Hora × Zona)")
        return procesar_hoja_14(file, sheet_name, header_row, debug)
    elif num_cuadro == 20:
        if debug:
            print(f"  Procesando hoja especial 20 (Modelo_Vehiculo × Tipo_Accidente)")
        return procesar_hoja_20(file, sheet_name, header_row, debug)
    elif num_cuadro in [22, 23]:
        if debug:
            print(f"  Procesando hoja compleja {num_cuadro} (Dept × Condicion × Sexo)")
        return procesar_hoja_22_23(file, sheet_name, header_row, debug)
    elif num_cuadro in [25, 26]:
        if debug:
            print(f"  Procesando hoja compleja {num_cuadro} (Hora/Edad × Condicion × Sexo)")
        return procesar_hoja_25_26(file, sheet_name, header_row, debug)
    elif num_cuadro in HOJAS_ESTRATIFICADAS:
        if debug:
            print(f"  Procesando hoja estratificada {num_cuadro}")
        return procesar_hoja_estratificada_optimizada(file, sheet_name, header_row, debug)
    
    df_preview = pd.read_excel(file, sheet_name=sheet_name, header=header_row, nrows=15)
    
    dim_fila = detectar_dimension_fila(df_preview, debug)
    columnas = [str(col).strip() for col in df_preview.columns]
    dim_columna = detectar_dimension_columna(columnas, debug)
    
    metrica = METRICAS_POR_CUADRO.get(num_cuadro, "valor")
    
    if dim_fila is None or dim_columna is None:
        if debug:
            print(f"  No detectado: fila={dim_fila}, col={dim_columna}")
        raise ValueError(f"No se detectaron dimensiones. fila={dim_fila}, col={dim_columna}")
    
    if debug:
        print(f"  {dim_fila} × {dim_columna} -> {metrica}")
    else:
        print(f"  {dim_fila} × {dim_columna} -> {metrica}")
    
    config_key = (dim_fila, dim_columna)
    if config_key not in RANGOS_CONFIG:
        rangos = {
            "fila_inicio": 1, 
            "fila_fin": 50,
            "col_inicio": 0, 
            "col_fin": min(50, len(df_preview.columns))
        }
        if debug:
            print(f"  Usando configuración por defecto para {config_key}")
    else:
        rangos = RANGOS_CONFIG[config_key]
    
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, 
                      nrows=rangos["fila_fin"] + 5)
    
    df = df.iloc[rangos["fila_inicio"]:rangos["fila_fin"], rangos["col_inicio"]:rangos["col_fin"]]
    
    if df.empty:
        raise ValueError("DataFrame vacío después de aplicar rangos")
    
    if df.columns[0].startswith('Unnamed'):
        df.rename(columns={df.columns[0]: dim_fila}, inplace=True)
    
    if 'Total' in df[dim_fila].astype(str).values:
        df = df[df[dim_fila].astype(str).str.strip().str.lower() != 'total']
    df = df.dropna(how='all')
    df = df[df[dim_fila].notna()]
    #df[dim_fila] = df[dim_fila].astype(str).str.strip()
    df[dim_fila] = df[dim_fila].apply(lambda x: str(x).strip() if pd.notna(x) else np.nan)
    
    value_cols = [col for col in df.columns if col != dim_fila]
    df_largo = df.melt(id_vars=[dim_fila], value_vars=value_cols, 
                       var_name=dim_columna, value_name=metrica)
    
    df_largo[metrica] = pd.to_numeric(df_largo[metrica], errors='coerce')
    #df_largo[dim_columna] = df_largo[dim_columna].astype(str).str.strip()
    df_largo[dim_columna] = df_largo[dim_columna].apply(lambda x: str(x).strip() if pd.notna(x) else np.nan)
    
    todas_dims = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                  "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                  "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
    for dim in todas_dims:
        if dim not in df_largo.columns:
            df_largo[dim] = np.nan
    
    cols_orden = todas_dims + [metrica, "hoja_origen"]
    df_largo = df_largo[[col for col in cols_orden if col in df_largo.columns]]
    df_largo["hoja_origen"] = sheet_name
    
    return df_largo

def procesar_todas_las_hojas(file, header_row=9, debug=False, solo_primeras=None):
    """Procesa todas las hojas del Excel"""
    
    xls = pd.ExcelFile(file)
    hojas = [h for h in xls.sheet_names if h.startswith('cuadro')]
    
    if solo_primeras:
        hojas = hojas[:solo_primeras]
        print(f"MODO TESTING: {solo_primeras} hojas\n")
    
    print(f"{len(hojas)} hojas encontradas\n")
    
    lista_df = []
    errores = []
    
    for hoja in hojas:
        try:
            print(f"{hoja}...", end=" ")
            df = transformar_hoja_auto(file, hoja, header_row, debug)
            lista_df.append(df)
            print(f"OK {len(df)} filas")
        except Exception as e:
            errores.append((hoja, str(e)))
            print(f"ERROR: {str(e)[:50]}")
    
    if errores:
        print(f"\n{len(errores)} errores:")
        for hoja, error in errores[:10]:
            print(f"   {hoja}: {error}")
    
    if lista_df:
        df_maestro = pd.concat(lista_df, ignore_index=True)
        
        print(f"\nDataset antes de limpieza: {len(df_maestro):,} filas")
        
        print(f"\nLimpiando dataset...")
        df_maestro = limpiar_dataset_maestro(df_maestro)
        
        print(f"\n{'='*60}")
        print(f"RESUMEN FINAL:")
        print(f"   Filas procesadas: {len(df_maestro):,}")
        print(f"   Hojas procesadas: {len(lista_df)}/{len(hojas)}")
        print(f"{'='*60}")
        
        metricas_disponibles = [col for col in ['accidentes_totales', 'vehiculos_involucrados_accidentes', 
                                              'victimas_totales', 'lesionados_totales', 'fallecidos_totales',
                                              'tasa_victimas_involucradas', 'tasa_lesionados', 'tasa_fallecidos'] 
                              if col in df_maestro.columns]
        
        print(f"\nTOTALES POR METRICA:")
        for metrica in metricas_disponibles:
            total = df_maestro[metrica].sum()
            no_null = df_maestro[metrica].notna().sum()
            print(f"  {metrica:35} : {total:12,.0f} ({no_null:,} registros)")
        
        print(f"\nDISTRIBUCION POR DIMENSION:")
        dimension_columns = ["Departamento", "Mes", "Anio", "Weekday", "Hora", "Zona", 
                           "Dia_Mes", "Tipo_Accidente", "Tipo_Vehiculo", "Color_Vehiculo", 
                           "Modelo_Vehiculo", "Edad", "Sexo", "Condicion_Conductor"]
        
        for dim in dimension_columns:
            if dim in df_maestro.columns and df_maestro[dim].notna().any():
                n = df_maestro[dim].notna().sum()
                unicos = df_maestro[dim].dropna().nunique()
                pct = (n / len(df_maestro)) * 100
                print(f"  {dim:25} : {n:6,} registros ({unicos:3} unicos) - {pct:5.1f}%")
        
        return df_maestro
    else:
        raise ValueError("No se procesó ninguna hoja")

# ============================================================================
# USO
# ============================================================================

"""
df_maestro = procesar_todas_las_hojas(
    file="archivo.xlsx",
    header_row=9,
    debug=False
)

df_maestro.to_csv('accidentes_guatemala.csv', index=False)
"""

df_maestro = procesar_todas_las_hojas(
    file=file,
    header_row=9,
    debug=False
)

# Guardar resultado
df_maestro.to_csv("dataset_consolidado_accidentes.csv", index=False, encoding='utf-8')
print(f"\n💾 Dataset guardado como 'dataset_consolidado_accidentes.csv'")

# Mostrar preview
print(f"\n📋 Preview del dataset final:")
print(df_maestro.head(10))
print(f"\n📊 Dimensiones finales: {df_maestro.shape}")