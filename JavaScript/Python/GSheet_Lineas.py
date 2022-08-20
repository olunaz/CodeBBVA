import pickle
import os
import sys
import configparser

from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

import datetime
import win32com.client as win32


CLIENT_SECRET_FILE = "cred_personal_gsheet.json"
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPE = ["https://www.googleapis.com/auth/drive"]


excel_path = ""
wsheet_name = ""
gsheet_id = ""
gsheet_name = ""

# Google Sheet Id
#gsheet_id = '1rZlxgW18qbdA0cmrZ3dCCSD-uWnQ3EfXdZGUHbSTTT4'

gsheet_id = '10h3t9voW7jgSU7vp9cBAR_x9T8yHyFowg8Qx-WAHTsQ'

gsheet_name = 'INDICADORES_COLAS'

# Microsoft Excel
excel_path = r"C:\Users\User\Downloads\NS_Data\Network_Support\Indicadores_NS.xlsm"
wsheet_name = 'INDICADORES_COLAS'


def create_gservice(CLIENT_SECRET_FILE, API_SERVICE_NAME,
	API_VERSION, SCOPE):

    cred = None

    pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPE)
            cred = flow.run_local_server()

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service.spreadsheets()

    except Exception as e:
        print('service error!')
        print(e)


#  Carga de datos desde Excel a Google Sheets 
def upload_gservice(gs,gsheet_id,gsheet_name,excel_path,wsheet_name):

    # Capturar rango desde el archivo local -----------------------------------
    # Abrir el archivo Excel desde la ruta y asignar variables
    xlApp = win32.Dispatch('Excel.Application') 
    wb = xlApp.Workbooks.Open(excel_path)
    ws = wb.Worksheets(wsheet_name)
    # Para marcar todo el rango actual de la region
    rngDataExcel = ws.Range('A1').CurrentRegion()
    # Convertir de tupla a lista
    rngData = list(rngDataExcel)
    # Quitar fila de encabezado
    rngData.pop(0) 

    # Capturar rango desde el archivo local
    response = gs.values().append(
        spreadsheetId=gsheet_id,
        valueInputOption='RAW',
        range = gsheet_name,
        body=dict(
            majorDimension='ROWS',
            values=rngData     
        )   
    ).execute()
    # Cerrar el libro de Excel
    wb.Close()

#  Descarga de datos desde Google Sheets a Excel
def download_gservice(gs,gsheet_id,gsheet_name,excel_path,wsheet_name):

    # extract information from 'gs' object to array
    try: 
        rows = gs.values().get(
                    spreadsheetId=gsheet_id,
                    range=gsheet_name,
                    ).execute().get('values')

        # Creando un nuevo libro
        # xlApp = win32.Dispatch("Excel.Application")
        # xlApp.Visible = 1
        # wb = xlApp.Workbooks.Add()
        # wsData = wb.Worksheets("Hoja1")
        # wsData.Name = 'Test01'

        # Seleccionando un libro 
        xlApp = win32.Dispatch('Excel.Application') 
        wb = xlApp.Workbooks.Open(excel_path) # Desde una ruta determinada
        # Para seleccionar una hoja creada
        wsData = wb.Worksheets(wsheet_name)
        # Para crear una hoja nueva
        # wsData = wb.Worksheets.Add()
        # wsData.Name = 'BASE'

        wsData.Cells.ClearContents()

        rowNumber = 1
        colCount = len(rows[0]) # 0 es el primer elemento de la lista "rows", para contar cuantos valores hay en la primera fila (cabecera) que sean el valor máximo
        # colCount = 14 #Mínimo número de columnas que deben estar completadas.


        for row in rows:
            # Para completar con valores vacios las celdas vacías del sheet y sean contadas en la fila
            lrow = len(row)  
            while lrow< colCount:
                row.append('')
                lrow+=1

            # Copiar las filas en la hoja de Excel 
            wsData.Range(wsData.cells(rowNumber, 1), wsData.cells(rowNumber, colCount)).value = row
            rowNumber += 1

        # Cuando se creo un nuevo libro y se enecesita "guardar como" en una ruta específica
        # FullPathName = r"D:\Python\DemoDescarga.xlsx"
        # wb.SaveAs(FullPathName)

        # Guardar en la ruta indicada 
        wb.Save()

    except Exception as e:
            print(e)

    # Cerrar el libro de Excel
    #wb.Close()


def MainPrueba():
        
    # Crear el servicio de Google
    gs = create_gservice(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPE)

    # Funcion de upload y download
    upload_gservice(gs,gsheet_id,gsheet_name,excel_path,wsheet_name)
    # Adicionar funcion o un timeset 
    #download_gservice(gs,gsheet_id,gsheet_name,excel_path,wsheet_name)

    k=input("press close to exit") 

    
  
MainPrueba()