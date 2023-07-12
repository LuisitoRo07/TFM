from pruebas_h0vsha import cargar_datos
import pandas as pd
from scipy.stats import shapiro, levene, f_oneway, kruskal
import statsmodels.api as sm
from statsmodels.formula.api import ols
import pingouin as pg
from to_excel import to_excel as tx

def filtro (data):
    Nq = 'Q8'
    subject = "Física"
        
    # Obtener columnas que contienen 'Q8' y 'Física'
    columna_freq = [col for col in data.columns if Nq in col and subject in col]
    # print(len(columna_freq))
    
    # Obtener columnas que contienen 'Q11'
    columnas_mtype = [col for col in data.columns if 'Q11' in col]
    
    # Obtener columnas que contienen 'Q12' y no contienen 'Otro'
    columnas_mstyle = [col for col in data.columns if 'Q12' in col and 'Otro' not in col]
    
    # Columna notas física
    columna_notas = [col for col in data.columns if 'Q17' in col and 'Física' in col]
    
    # Filtrar datos usando las columnas obtenidas
    datos_filtrados = data[columna_freq + columnas_mtype + columnas_mstyle + columna_notas]
    
    # Mantener solo las filas donde columna_freq sea mayor que 3
    filtro_freq = datos_filtrados[columna_freq] > 3
    datos_filtrados = datos_filtrados[filtro_freq.any(axis=1)]
    
    # Filtrar los que en columna_freq no tienen ningún valor numérico
    filtro_numerico = data[columna_notas].apply(lambda x: pd.to_numeric(x, errors='coerce')).notna().all(axis=1)
    datos_filtrados = datos_filtrados[filtro_numerico]

    return datos_filtrados

def grupos_mtype(datos_filtrados):
    grupos = pd.DataFrame(columns=['Grupo', 'Notas'])

    columnas_q11 = [col for col in datos_filtrados.columns if 'Q11' in col]

    for col in columnas_q11:
        titulo_grupo = col.split('(')[-1].strip(')')  # Extraer el texto entre paréntesis del título de la columna
        notas = datos_filtrados.loc[datos_filtrados[col] == 1, datos_filtrados.columns[-1]]
        grupo = pd.DataFrame({'Grupo': titulo_grupo, 'Notas': notas})
        grupos = pd.concat([grupos, grupo], ignore_index=True)

    return grupos

def estadísticas_iniciales(datos):
    # Contar el total en cada grupo
    conteo_grupos = datos['Grupo'].value_counts().reset_index()
    conteo_grupos.columns = ['Grupo', 'Total de alumnos']
    # Calcular la media, mediana y desviación estándar de la columna 'Notas'
    media = datos.groupby('Grupo')['Notas'].mean().round(2)
    mediana = datos.groupby('Grupo')['Notas'].median()
    desviacion_tipica = datos.groupby('Grupo')['Notas'].std()
    # Crear una tabla con las medidas estadísticas
    medidas_estadisticas = pd.DataFrame({'Media': media, 'Mediana': mediana, 'Desviación Típica': desviacion_tipica})
    # Unir el conteo en cada grupo y las medidas estadísticas en una misma tabla
    tabla_estadisticas = pd.merge(conteo_grupos, medidas_estadisticas, left_on='Grupo', right_index=True)
    # Imprimir la tabla de medidas estadísticas
    # print(tabla_estadisticas)
    # print("\n")
    return tabla_estadisticas

def test_normalidad(datos):
    # Eliminar filas con datos faltantes y convertir 'Notas' a tipo numérico
    datos['Notas'] = pd.to_numeric(datos['Notas'], errors='coerce')
    datos = datos.dropna(subset=['Notas'])

    # Calcular la normalidad usando pg.normality()
    tabla_resultados = pg.normality(data=datos, dv='Notas', group='Grupo')

    return tabla_resultados
    
def test_homocedasticidad(grupos):  
   # Realizar el test de homocedasticidad (Levene)
    levene_test = pg.homoscedasticity(data=grupos, dv='Notas', group='Grupo')
    levene_results = levene_test[['W', 'pval', 'equal_var']]
    levene_results.columns = ['Estadístico de Levene', 'Valor p', 'Igualdad de varianzas']
    return levene_results

def test_previos (grupos):
    return test_normalidad(grupos), test_homocedasticidad(grupos)

def anova(grupos):
    # Anova de una vía
    anova_result = pg.anova(data=grupos, dv='Notas', between='Grupo', detailed=True)
    
    # Mostramos ANOVA
    print("Resultado detallado del análisis de varianza (ANOVA):")
    print(anova_result)
    print("\n")
    
    return anova_result

def no_parametric_test (grupos):
    # Realizar el test de Kruskal-Wallis
    kruskal_result = pg.kruskal(data=grupos, dv='Notas', between='Grupo')
    
    # Imprimir el resultado detallado
    print("Resultado detallado del test de Kruskal-Wallis:")
    print(kruskal_result)
    print("\n")
    
    return kruskal_result

def comparacion (grupos):
    tabla_anova = anova (grupos)
    tabla_kruskal = no_parametric_test (grupos)
    return tabla_anova, tabla_kruskal

def main():
    datos = cargar_datos()
    datos_filtrados = filtro(datos)
    grupos = grupos_mtype(datos_filtrados)
    tabla_alumnos = estadísticas_iniciales(grupos)
    print(tabla_alumnos)
    tabla_normalidad, tabla_homocedasticidad = test_previos (grupos)
    print('\n')
    print(tabla_normalidad)
    print ('\n')
    print(tabla_homocedasticidad)
    print ('\n')
    tabla_anova, tabla_kruskal = comparacion (grupos)
    rute = 'F:/Cursos y master/Master educación/TFM/Resultados/Resultados_mtype.xlsx'
    sheet_name = 'mtype'
    guardado = tx(rute, sheet_name, tabla_alumnos, tabla_normalidad, tabla_homocedasticidad, tabla_anova, tabla_kruskal)
    if guardado == 1:
        print ('Tablas guardadas correctamente \n')
    

if __name__ == "__main__":
    main()
