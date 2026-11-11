import pandas as pd

folder = r'C:\Users\julian.cobos\Downloads\reseguimientos'
df_int = pd.read_excel(folder + r'\Interacciones.xlsx')

# Ver todas las claves únicas en Recomendaciones
all_keys = set()
for val in df_int['Recomendaciones'].dropna():
    for line in str(val).split('\n'):
        line = line.strip()
        if ':' in line:
            key = line.split(':', 1)[0].strip()
            all_keys.add(key)

print('=== TODAS LAS CLAVES EN RECOMENDACIONES ===')
for k in sorted(all_keys):
    print(f'  {repr(k)}')

col_caso = df_int.columns[0]
print()
print(f'Total filas interacciones: {len(df_int)}')
print(f'Casos unicos: {df_int[col_caso].nunique()}')

# Ver Inicio con hora real (leer como object para ver el raw)
df_int2 = pd.read_excel(folder + r'\Interacciones.xlsx', dtype={'Inicio': str})
print()
print('Muestra Inicio (raw string):')
for v in df_int2['Inicio'].head(10):
    print(' ', repr(v))

# Cuántas filas tienen mismo Número del caso y mismo día en Inicio
df_int['Inicio_dt'] = pd.to_datetime(df_int['Inicio'])
df_int['Inicio_date'] = df_int['Inicio_dt'].dt.date
dups = df_int.groupby([col_caso, 'Inicio_date']).size()
print()
print('Casos con >1 interacción el mismo día:')
print(dups[dups > 1].head(10))
