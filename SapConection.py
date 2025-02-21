# Importing the Libraries
from Utils import utils
import win32com.client
import sys
import subprocess

# This function will Login to SAP from the SAP Logon window
class SapGUI():
    def __init__(self, ) -> None:
                try:
                        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  

                        utils.kill_sap()

                        subprocess.Popen (self.path) #Abrir o SAP
                        
                        utils.wait_open_window("SAP Logon 800")

                        self.SapGuiAuto = win32com.client.GetObject('SAPGUI')#Instance SAP      
                        
                        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
                                return
                        
                        application = self.SapGuiAuto.GetScriptingEngine
                        if not type(application) == win32com.client.CDispatch:
                                self.SapGuiAuto = None
                                return  
                        
                        # Open Connection 
                        self.connection = application.OpenConnection ("101 : RP1SUB - SUBSEA S/4 HANA 2020 Production System", True)
                        #self.connection = application.OpenConnection ("121 : RQ1SUB - SUBSEA S/4 HANA 2023 QA System", True)

                        if not type(self.connection) == win32com.client.CDispatch:
                                application = None
                                self.SapGuiAuto = None
                                return
                        
                        self.session = self.connection.Children(0)   
                        if not type(self.session) == win32com.client.CDispatch:
                                self.connection = None
                                application = None
                                self.SapGuiAuto = None   

                except Exception as err:
                        print (err, type(err))
                        print("Error",sys.exc_info()[0])
                
                