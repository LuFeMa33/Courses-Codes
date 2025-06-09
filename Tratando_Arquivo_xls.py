import win32com.client as win32
import pandas as pd
import time
import os
import schedule

caminho_arquivo = r'C:\Users\Users\Downloads\Planilha.xls'

def tratar_arquivo():

    print('Tratando arquivo em formato xls')
    excel_app = win32.Dispatch('Excel.Application')
    wb = excel_app.Workbooks.open(caminho_arquivo)
    excel_app.DisplayAlerts = False
    wb.Save() 
    excel_app.quit()
    time.sleep(2)

def salvar_arquivo():

    data_agora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") 
    caminho_salvamento = r'C:\Users\Users\Downloads\Planilha_' + data_agora + '.xlsx'
    print('Salvando o novo arquivo em xlsx')
    df_monitor = pd.read_excel(caminho_arquivo)
    df_monitor.to_excel(caminho_salvamento, index=False)
    print(f'Arquivo salvo em: {caminho_salvamento}')
    time.sleep(30)
    os.remove(caminho_arquivo)
    print('Arquivo primário excluído')

def executando_processo():
    tratar_arquivo()
    salvar_arquivo()

schedule.every(20).minutes.do(executando_processo)

while True:  
    schedule.run_pending()
    print("Aguardando próximo agendamento...")  
    time.sleep(10)











