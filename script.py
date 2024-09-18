import pandas as pd

# Ruta al archivo Excel
file_path = '/Users/bam/Library/Mobile Documents/com~apple~CloudDocs/Documents/WorkWork/GrupoLaser/Dir_Gen/A_Comercial/Listas_Precios/Envases-Natural.xlsx'  # Cambia esto por la ruta a tu archivo

# Cargar los datos desde Excel
excel_data = pd.read_excel(file_path)

# Función para calcular el costo y precios con utilidad
def calcular_costos_con_utilidad(peso_gramos, precio_resina_usd, tipo_cambio):
    # Calcular el precio de la resina en pesos mexicanos
    precio_resina_mxn = precio_resina_usd * tipo_cambio  # Precio de la resina en MXN por kilo
    precioResina = precio_resina_mxn / 1000  # Precio de la resina por gramo en pesos
    
    # Cálculo de costos
    costoDirecto = precioResina * peso_gramos  # Costos directos
    costoTotal = costoDirecto / 0.6  # Costo total, donde los costos directos son el 60%
    
    # Cálculo de precios de venta con diferentes márgenes de utilidad
    precioVentaBajo = costoTotal / 0.8  # 20% de utilidad
    precioVentaNormal = costoTotal / 0.7  # 30% de utilidad
    precioVentaAlto = costoTotal / 0.6  # 40% de utilidad
    
    return costoTotal, precioVentaBajo, precioVentaNormal, precioVentaAlto

# Aplicar el cálculo a cada producto
precio_resina_usd = 1.69  # Precio de la resina en USD por kilo
tipo_cambio = 21  # Tipo de cambio en MXN por USD

# Añadir los cálculos al archivo
excel_data[['Costo_PET_MXN', 'Precio_20_Utilidad', 'Precio_30_Utilidad', 'Precio_40_Utilidad']] = excel_data.apply(
    lambda row: calcular_costos_con_utilidad(
        row['Peso_Gramos'], 
        precio_resina_usd, 
        tipo_cambio
    ), 
    axis=1, result_type='expand'
)

# Guardar los resultados en un nuevo archivo Excel
output_file_path = '/Users/bam/Library/Mobile Documents/com~apple~CloudDocs/Documents/WorkWork/GrupoLaser/Dir_Gen/A_Comercial/Listas_Precios/Envases-NaturalSALIDA.xlsx'  # Cambia esto por la ruta donde quieras guardar el archivo
excel_data.to_excel(output_file_path, index=False)

print("Cálculos completados y archivo guardado.")