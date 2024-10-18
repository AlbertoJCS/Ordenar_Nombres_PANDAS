import os
import pandas as pd
import datetime as dt

# Lee el archivo de Excel
ruta_archivo = 'J:\\python\\Nombres\\Nombres.xlsx'
df = pd.read_excel(ruta_archivo, sheet_name='Hoja1')

# Crear una lista para almacenar los resultados
perf = {}
imperf = {}
soc = {}
lim = {}

# Recorrer las filas de 'CODIGO NOMBRE' y 'id'
for index, row in df.iterrows():
    codigo_nombre = row['CODIGO NOMBRE']
    orden = row['ORDEN']

    # Dividir 'CODIGO NOMBRE' en palabras
    partes = codigo_nombre.split()

    if 'limitada' in codigo_nombre.lower():
        lim[orden] = partes
        continue

    if 'sociedad' in codigo_nombre.lower():
        soc[orden] = partes
        continue

    # Si tiene 3 o 4 partes, lo almacenamos
    if len(partes) == 3 or len(partes) == 4:
        perf[orden] = partes
        
    else:
        imperf[orden] = partes 


######################  PERFECTOS   ############################
for orden, partes in perf.items():
    if len(partes) == 3:
        partes.insert(1, " ")  
    perf[orden] = partes
###################################################



######################  SOCIEDAD ANONIMA   ############################ 
for orden, partes in soc.items():
    if len(partes) >= 2 and partes[-1].strip().upper() == "ANONIMA" and partes[-2].strip().upper() == "SOCIEDAD":
        nuevas_partes1 = ["S.A.", " ", " ".join(partes[:-2]), " "]  # Unir el resto sin "SOCIEDAD ANONIMA"
        soc[orden] = nuevas_partes1
    else:
        soc[orden] = partes
###################################################

 

######################  SOCIEDAD LIMITADA   ############################ 
for orden, partes in lim.items():
    if len(partes) >= 2 and partes[-1].strip().upper() == "LIMITADA" and partes[-2].strip().upper() == "RESPONSABILIDAD":
        nuevas_partes2 = ["L.T.D.A.", " ", " ".join(partes[:-2]), " "]  # Unir el resto sin "RESPONSABILIDAD LIMITADA"
        lim[orden] = nuevas_partes2
    else:
        lim[orden] = partes
###################################################



######################  IMPERFECTOS   ############################
# Definir los patrones a buscar
patrones_3_valores = [
    ["DE", "LOS", "ANGELES"],
    ["DE", "LA", "TRINIDAD"],
    ["DE", "LAS", "NIEVES"],
    ["DEL", "CARMEN"],
    ["DE", "JESUS"],
    ["DEL", "SOCORRO"],
] 

for orden, partes in imperf.items():
    for patron in patrones_3_valores:
        # Recorremos la lista y buscamos el patrón
        for i in range(len(partes) - len(patron) + 1):
            if partes[i:i + len(patron)] == patron:
                
                # Caso 1: Lista tiene 5 elementos
                if len(partes) == 5:
                    # Combina el primer elemento con el patrón y el resto se mantiene
                    nuevas_partes3 = [partes[0] + " " + patron[0], patron[1]] + partes[i + len(patron):]
                    imperf[orden] = nuevas_partes3
                    break  # Salimos de la búsqueda de patrones

                # Caso 2: Lista tiene 6 elementos
                elif len(partes) == 6:
                    partes_unidas = " ".join(partes[i:i + len(patron) - 1])  # Une el patrón "DE LOS"
                    nuevas_partes4 = [partes[0] + " " + partes_unidas] + partes[i + len(patron) - 1:]
                    imperf[orden] = nuevas_partes4
                    break  # Salimos de la búsqueda de patrones

                # Caso 3: Lista tiene 7 elementos
                elif len(partes) == 7:
                    nuevas_partes5 = [partes[0], partes[1], partes[5], partes[6]]
                    imperf[orden] = nuevas_partes5
                    break  # Salimos de la búsqueda de patrones

        # Si se encontró el patrón, salimos del bucle de patrones
        else:
            # Si no se encontró el patrón en esta iteración, continúa con la siguiente
            continue
        # Salimos del segundo bucle si se encontró un patrón
        break


patrones_3_valores = [
    ["DEL", "VALLE"],
    ["DA", "SILVA"]
] 



# Crear un escritor de Excel
with pd.ExcelWriter('resultados_nombres.xlsx') as writer:
    # Convertir cada diccionario en un DataFrame y guardarlo en una hoja diferente
    pd.DataFrame.from_dict(perf, orient='index').to_excel(writer, sheet_name='Perfectos', header=False)
    pd.DataFrame.from_dict(imperf, orient='index').to_excel(writer, sheet_name='Imperfectos', header=False)
    pd.DataFrame.from_dict(soc, orient='index').to_excel(writer, sheet_name='Sociedad', header=False)
    pd.DataFrame.from_dict(lim, orient='index').to_excel(writer, sheet_name='Limite', header=False)

    # Crear un DataFrame combinado de todos los diccionarios
    combined_df = pd.DataFrame({
        'Tipo': ['Perfecto'] * len(perf) + ['Imperfecto'] * len(imperf) + ['Sociedad'] * len(soc) + ['Limite'] * len(lim),
        'Orden': list(perf.keys()) + list(imperf.keys()) + list(soc.keys()) + list(lim.keys()),
        'Nombres': list(perf.values()) + list(imperf.values()) + list(soc.values()) + list(lim.values())
    })

    # Guardar el DataFrame combinado en una nueva hoja
    combined_df.to_excel(writer, sheet_name='Combinados', index=False)

print("Archivo Excel creado con éxito.")