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


def tabulado(anio, mes):
  """Descarga un único tabulado y lo devuelve como un DataFrame de pandas.

    Argumentos:
    anio -- el año del periodo a descargar (entero)
    mes  -- el mes del periodo a descargar (entero: 1, 2, ..., 12)
    """
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
  resultado = [f'20{str(anio).zfill(4)[-2:]}-{str(mes).zfill(2)}']
  for variable in mapeo.keys():
    valor = wb['1.1'][mapeo[variable]].value
    resultado.append(valor)
  df = pd.DataFrame([resultado], columns = ['Periodo'] + list(mapeo.keys()))
  
  # Eliminar archivo temporal
  remove('temp_data.xlsx')

  return df


def periodos(anio_inicial, mes_inicial, anio_final, mes_final):
  if anio_inicial == anio_final:
    periodos = [(anio_inicial, m) for m in range(mes_inicial, mes_final + 1)]
  elif anio_final == anio_inicial + 1:
    periodos = [(anio_inicial, m) for m in range(mes_inicial, 13)] + [(anio_final, m) for m in range (1, mes_final + 1)]
  else:
    periodos = [(anio_inicial, m) for m in range(mes_inicial, 13)]
    for y in range(anio_inicial + 1, anio_final):
      periodos = periodos + [(y, m) for m in range(1, 13)]
    periodos = periodos + [(anio_final, m) for m in range (1, mes_final + 1)]
  return periodos


def tabulados(anio_inicial, mes_inicial, anio_final, mes_final):
  """Descarga un conjunto de tabulados y los devuelve como un DataFrame de pandas.

    Argumentos:
    anio_inicial -- el año del periodo inicial a descargar (entero)
    mes_inicial  -- el mes del periodo inicial a descargar (entero: 1, 2, ..., 12)
    anio_final -- el año del periodo final a descargar (entero)
    mes_final  -- el mes del periodo final a descargar (entero: 1, 2, ..., 12)
    """
  dataframes = []
  for periodo in periodos(anio_inicial, mes_inicial, anio_final, mes_final):
    df = tabulado(*periodo)
    dataframes.append(df)
  resultado = pd.concat(dataframes, ignore_index = True).sort_values('Periodo', ascending = True)
  return resultado
