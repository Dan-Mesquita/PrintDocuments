from datetime import date
import os
import subprocess
import win32com.client
import pygetwindow as gw
import time

class utils:

    @staticmethod
    def get_time():
        date_today = date.today()
        date_now_format = date_today.strftime('%d/%m/%Y')
        return date_now_format
    
    @staticmethod
    def transforme_path(path_user):
        user_profile_dir = os.path.expanduser("~")
        return user_profile_dir + path_user
    
    @staticmethod
    def kill_sap(): 
                subprocess.run (['taskkill', '/f', '/im', "saplgpad.exe"]) #"sapglpad.exe" é o nome do SAP que está no Task Manager do Windows
                subprocess.run (['taskkill', '/f', '/im', "sapgui.exe"])
                subprocess.run (['taskkill', '/f', '/im', "saplogon.exe"]) #Consultar o Task Manager do Windows para checar qual o nome do SAP, por enquanto encontramos apenas estas 3 opções
    
    #Deleta todos os arquivos ".xlsx" da pasta escolhida; Recebe por parâmetro o caminho da pasta selecionada 
    @staticmethod
    def clean_paste(path):
        Output_Files = [arquivo for arquivo in os.listdir(path) if arquivo.endswith(".xlsx")]
        for Output_File in Output_Files:
            os.remove(os.path.join(path, Output_File))
    
    @staticmethod
    def close_excel(file_path):
        excel = win32com.client.Dispatch("Excel.Application")
        for workbook in excel.Workbooks:
            if workbook.Name == file_path:
                workbook.Save()
                workbook.Close()
                break

    @staticmethod
    def wait_open_window(Page):
        while True:
            if any(Page in title for title in gw.getAllTitles()):
                break
            time.sleep(1)
    
    @staticmethod
    def wait_open_window_two_window(Page_one,Page_two):
        while True:
            if any(Page_one in title for title in gw.getAllTitles()):
                return Page_one
            elif any(Page_two in title for title in gw.getAllTitles()):
                return Page_two
            time.sleep(1)
    
    @staticmethod
    def wait_close_window(Page):
        open = True
        while open == True:
            if any(Page in title for title in gw.getAllTitles()):
                time.sleep(1)
            else:
                open = False


    

    
   
    