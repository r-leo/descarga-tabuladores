import requests
import pandas as pd
from openpyxl import load_workbook
from os import remove

mapeo = {
  'Población total': 'M11',
  'PEA': 'M12',
  'PEA ocupada': 'M13',
  'PEA desocupada': 'M14',
  'Años de escolaridad promedio de la PEA': 'M255',
  'Horas trabajadas a la semana por la población ocupada (media)': 'M258',
  'Horas trabajadas a la semana por la población ocupada (mediana)': 'M259',
  'Ingreso (pesos) por hora trabajada de la población ocupada (empleadores) (media)': 'M261',
  'Ingreso (pesos) por hora trabajada de la población ocupada (empleadores) (mediana)': 'M262',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores por cuenta propia) (media)': 'M267',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores por cuenta propia) (mediana)': 'M268',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores por cuenta propia en actividades no calificadas) (media)': 'M270',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores por cuenta propia en actividades no calificadas) (mediana)': 'M271',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores subordinados y remunerados asalariados) (media)': 'M273',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores subordinados y remunerados asalariados) (mediana)': 'M274',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores subordinados y remunerados con percepciones no salariales) (media)': 'M276',
  'Ingreso (pesos) por hora trabajada de la población ocupada (trabajadores subordinados y remunerados con percepciones no salariales) (mediana)': 'M277',
}

def tabulados(anio, mes):

  # Construir periodo (MMAA)
  periodo = str(mes).zfill(2) + str(anio).zfill(4)[-2:]

  # Enviar request
  url = f'https://www.inegi.org.mx/contenidos/programas/enoe/15ymas/tabulados/enoe_indicadores_estrategicos_{periodo}.xlsx'
  r = requests.get(url, stream=True)

  # Guardar archivo tempral descargado
  with open('temp_data.xlsx', 'wb') as f:
    for chunk in r.iter_content(chunk_size = 16*1024):
      f.write(chunk)
  
  # Leer archivo temporal y extraer de acuerdo al mapeo
  wb = load_workbook('temp_data.xlsx')
  resultado = []
  for variable in mapeo.keys():
    valor = wb['1.1'][mapeo[variable]].value
    resultado.append([variable, valor])
  df = pd.DataFrame(resultado, columns = ['Variable', 'Valor'])
  
  # Eliminar archivo temporal
  remove('temp_data.xlsx')

  return df
