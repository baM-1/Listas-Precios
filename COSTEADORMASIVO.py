import pandas as pd

# Variables globales
tipo_cambio = 22      # Tipo de cambio USD a moneda local
porcentaje_costos_indirectos = 40  # Porcentaje de costos indirectos
margen_utilidad = 30               # Porcentaje de margen de utilidad deseado
margen_utilidad_alto = 40          # Porcentaje de margen de utilidad alto
margen_utilidad_bajo = 20          # Porcentaje de margen de utilidad bajo

precios_resina_usd = {
    'PET': 1.69,      # Precio en USD/kg
    'PP': 1.70        # Precio en USD/kg
}

# Leer el archivo de Excel
df_productos = pd.read_excel('/Users/bam/Documents/WorkWork/GrupoLaser/Dir_Gen/A_Comercial/Listas_Precios/LS_BASE.xlsx', sheet_name='Productos')

# Asegurarnos de que 'Peso Neto (g)' es numérico
df_productos['Peso Neto (g)'] = pd.to_numeric(df_productos['Peso Neto (g)'], errors='coerce')

# Verificar si hay valores NaN
print(df_productos[df_productos['Peso Neto (g)'].isna()])

# Función para calcular costos y precios
def calcular_costo_y_precio(row):
    # Obtener valores de la fila
    peso_neto = row['Peso Neto (g)']
    resina = row['Resina']
    precio_pigmento_usd = row['Precio Pigmento USD/kg']
    porcentaje_pigmento = row['Porcentaje Pigmento (%)']
    precio_resina_usd = row.get('Precio Resina USD/kg', precios_resina_usd[resina])

    # Aplicar la merma del 4% solo si hay pigmento
    if porcentaje_pigmento > 0:
        porcentaje_merma = 4
    else:
        porcentaje_merma = 0
    
    # Paso 1: Peso ajustado por merma
    peso_ajustado = peso_neto * (1 + porcentaje_merma / 100)
    
    # Paso 2: Peso de pigmento y resina
    peso_pigmento = peso_ajustado * (porcentaje_pigmento / 100)
    peso_resina = peso_ajustado - peso_pigmento
    
    # Convertir pesos a kilogramos
    peso_resina_kg = peso_resina / 1000
    peso_pigmento_kg = peso_pigmento / 1000
    
    # Precios en moneda local
    precio_resina_mxn = precio_resina_usd * tipo_cambio
    precio_pigmento_mxn = precio_pigmento_usd * tipo_cambio
    
    # Costo de materiales
    costo_resina = peso_resina_kg * precio_resina_mxn
    costo_pigmento = peso_pigmento_kg * precio_pigmento_mxn
    costo_materiales = costo_resina + costo_pigmento
    
    # Costo total incluyendo costos indirectos
    costo_total = costo_materiales / (1 - porcentaje_costos_indirectos / 100)
    
    # Precios de venta con diferentes márgenes de utilidad
    precio_venta_20 = costo_total * (1 + margen_utilidad_bajo / 100)
    precio_venta_30 = costo_total * (1 + margen_utilidad / 100)
    precio_venta_40 = costo_total * (1 + margen_utilidad_alto / 100)
    
    # Resultados detallados
    resultado = {
        'CODIGO SAE': row['CODIGO SAE'],
        'DESCRIPCION': row['DESCRIPCION'],
        'Peso Neto (g)': peso_neto,
        'Costo Resina': costo_resina,
        'Costo Pigmento': costo_pigmento,
        'Costo Materiales': costo_materiales,
        'Costo Total': costo_total,
        'Precio Venta (30%)': precio_venta_30,
        'Precio Venta (20%)': precio_venta_20,
        'Precio Venta (40%)': precio_venta_40
    }
    
    return resultado

# Crear una lista para almacenar los resultados
resultados = []

# Iterar sobre cada fila del DataFrame
for index, row in df_productos.iterrows():
    resultado = calcular_costo_y_precio(row)
    resultados.append(resultado)

# Convertir los resultados a un DataFrame para facilitar la visualización
df_resultados = pd.DataFrame(resultados)

# Exportar los resultados a un nuevo archivo de Excel
df_resultados.to_excel('resultados_productos.xlsx', index=False)

# Mostrar los resultados
print(df_resultados)