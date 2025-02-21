from SapConection import SapGUI
import pyautogui
from Utils import utils
import uiautomation as auto
from datetime import date
import time
import fitz
from Path import path
import os
from Main_treatment import main_treatment

class tcodes ():
    def __init__(self):
        self._SAP = SapGUI()
        
        try:
            self._SAP.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "101"
            self._SAP.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "362308"
            self._SAP.session.findById("wnd[0]").sendVKey (0)
        except:
            ...

    def start_tcode (self,tcode):
        self._SAP.session.findById("wnd[0]/tbar[0]/okcd").text = (tcode)
        self._SAP.session.findById("wnd[0]").sendVKey (0)

    def Tcode_CV04N(self, Document, OP):
        try:
            mt = main_treatment()

            image = mt.stamp(OP)  # Imagem gerada (300x200)
            
            if image == "User not found":
                return "Fail - User not found"
            
            #Inicia variavel
            self.start_tcode("/nCV04N")

            #Seleciona o documento e a ultima versão
            self._SAP.session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/ctxtSTDOKNR-LOW").text = Document
            self._SAP.session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/radLATVER").select ()
            self._SAP.session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/ctxtSTDOKAR-LOW").text = ""
            self._SAP.session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/ctxtSTDOKTL-LOW").text = ""
            self._SAP.session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/ctxtSTDOKVR-LOW").text = ""
            self._SAP.session.findById("wnd[0]/usr/tabsMAINSTRIP/tabpTAB1/ssubSUBSCRN:SAPLCV100:0401/subSCR_MAIN:SAPLCV100:0402/txtRESTRICT").text = ""
            
            #Roda a transação
            self._SAP.session.findById("wnd[0]").sendVKey (8)
            
            #Por algum motivo, quando se seleciona isso, da um problema no campo "Maximum Number of Hits", esse campo seleciona "ok"
            self._SAP.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            #Salva o tipo do documento
            doc_type = self._SAP.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").GetCellValue(0, "DOKAR")
            
            #Da um duplo clique na celula retornada
            self._SAP.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
        
            #Daqui pra frente, não foi usado SAP Script
            #Inicia a variavel grid que vai conter a tabela da transação CV04N
            grid = self._SAP.session.findById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell[1]")
            
            #Inicia as colunas
            columns = {grid.GetColumnTitleFromName(name): name for name in grid.GetColumnNames()}
            
            #Inicia as chaves dos modulos
            node_keys = grid.GetAllNodeKeys()

            #Inicia uma lista em branco
            originals_list: list[dict[str, str]] = []

            #Percore um laço por todas as chaves
            for key in node_keys:
                node_dict = {}

                for col_title, col_name in columns.items():
                    text = grid.GetItemText(key, col_name)
                    tooltip = None
                    if not text:
                        tooltip = grid.GetItemToolTip(key, col_name)

                    node_dict[col_title] = tooltip if tooltip and tooltip != "Blank" else text

                originals_list.append(node_dict)

            Finded_Converted_Document = False
            
            #Testa os campos para encontrar o documento com o nome "CONVERTED DOCUMENT"
            for row_index, original in enumerate(originals_list):
                if original["Application/Description"] == "CONVERTED DOCUMENT":
                    Finded_Converted_Document = True
                    grid.NodeContextMenu(node_keys[row_index])
                    grid.SelectContextMenuItem("CF_EXP_COPY")
                    break
            
            if Finded_Converted_Document == False:
                return "Fail - Converted Document not found"

            self._SAP.session.findById("wnd[1]/usr/ctxtDRAW-FILEP").text = f"{path.path_export}{Document}.pdf"
            self._SAP.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            pdf_document = fitz.open(f"{path.path_export}{Document}.pdf")  # Obtém o objeto do PDF aberto

            if doc_type == "PDC":
                doc_len = [0]

            else:
                doc_len = range(len(pdf_document))

            # Loop em todas as páginas
            for page_num in doc_len:
                            
                page = pdf_document[page_num]

                position = mt.find_space(page, 120, 80)

                if position:
                    x, y = position

                    # Inserir a imagem na posição encontrada
                    page.insert_image(fitz.Rect(x, y, x + 120, y + 80), pixmap=image)

                else:
                    return f"Fail - Page {page_num+1} without space"

            pdf_document.save(f"{path.path_export}{Document}_edited.pdf") 
                
            pdf_document.close()
                
            os.startfile(f"{path.path_export}{Document}_edited.pdf")
                
            utils.wait_open_window('Adobe Acrobat')
            
            pyautogui.hotkey('ctrl', 'p')
                    
            utils.wait_open_window('Print')
                
            Print_Dialog = auto.WindowControl(Name="Print")
            Print_Dialog.SetActive()
                
            #---------------------------Inicio da explicação--------------------------------
            # Na primeira versão, as proximas duas linhas selecionam a impressora Global Print, porém essa é a impressora padrão, devido a isso, essa parte do codigo foi comentada.
            # De todo modo, o codigo foi mantido, caso a impressora global print deixe de ser padrão no futuro é possivel descomentar e ajustar 
            # O if testa se o checkbox esta flegado(True) e se deveria ser desflegado (False) e vice e versa.
            # Caso o flag esteja errado, ele clica. Como o padrão é vir sempre flegado (True), ele clica caso a opção correta seja desflegado
            # Essa parte do codigo foi mantida com a mesma justificativa de caso pare de ser o padrão, da pra descomentar e seguir.
            #------------------------------Fim da explicação--------------------------------
                
            # Print_Dialog.ComboBoxControl(Name="Printer:").Click()
            # Print_Dialog.ListItemControl(Name="Global_Print").Click()
            #if (Print_Dialog.CheckBoxControl(Name="Print on both sides of paper").GetTogglePattern().ToggleState == auto.ToggleState.On and Duplex=="No") or (Print_Dialog.CheckBoxControl(Name="Print on both sides of paper").GetTogglePattern().ToggleState == auto.ToggleState.Off and Duplex=="Yes"):
            #    Print_Dialog.CheckBoxControl(Name="Print on both sides of paper").Click()
                
            if doc_type !="PDC":
                Print_Dialog.CheckBoxControl(Name="Print on both sides of paper").Click()
                    
                Print_Dialog.ButtonControl(Name="Page Setup...").Click()    
                    
                Page = utils.wait_open_window_two_window('Page Setup','Configurar Página')   

                if Page == 'Page Setup':
                    Page_Config = auto.WindowControl(Name="Page Setup")
                    Page_Config.SetActive()   
                    Page_Config.ComboBoxControl(Name="Size:").Click()
            
                elif Page == 'Configurar Página':
                    Page_Config = auto.WindowControl(Name='Configurar Página')
                    Page_Config.SetActive()   
                    Page_Config.ComboBoxControl(Name="Tamanho:").Click()
                
                else:
                    return "Fail - Print Setup Page Not Found"
                
    
                Page_Config.ListItemControl(Name="A3").Click()
                Page_Config.ButtonControl(Name="OK").Click()

            Print_Dialog.ButtonControl(Name="Print").Click()

            time.sleep(1)

            utils.wait_close_window('Progress')

            time.sleep(1)
                
            return "Success"
        
        except:
            return "Generic Fail"

    def Tcode_CO02(self, Document):

        try:
            self.start_tcode("/nCO02")

            self._SAP.session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = Document
            self._SAP.session.findById("wnd[0]/usr/radR62CLORD-FLG_OVIEW").select()
            self._SAP.session.findById("wnd[0]").sendVKey (0)

            material = int(self._SAP.session.findById("wnd[0]/usr/ctxtCAUFVD-MATNR").text)
            
            self._SAP.session.findById("wnd[0]/mbar/menu[0]/menu[8]/menu[5]").select()

            self._SAP.session.findById("wnd[1]/usr/subSPOOLPARAM:SAPL0C17:0455/tblSAPL0C17TCTRL_T496D_DSP_0455/chkT496D_SUB_DISPLAY-X[0,0]").selected = "true"
                
            for i in range(1,20):
                try:
                    self._SAP.session.findById("wnd[1]/usr/subSPOOLPARAM:SAPL0C17:0455/tblSAPL0C17TCTRL_T496D_DSP_0455/chkT496D_SUB_DISPLAY-X[0,"+str(i)+"]").selected = "false"
                except:
                    continue
                    
            self._SAP.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self._SAP.session.findById("wnd[0]/tbar[0]/btn[86]").press()
            
            try:
                self._SAP.session.findById("wnd[1]").close()
                self._SAP.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                self._SAP.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                return material, "Fail - OP alredy printed"

            except:
                self._SAP.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                return material, "Success"
        
        except:
            return material, "Generic Fail"


    def Tcode_ZZPLM_BOM(self, Document):
        
        try:
            self.start_tcode("/nZZPLM_BOM")

            self._SAP.session.findById("wnd[0]/usr/ctxtP_MTNRV").text = Document
            self._SAP.session.findById("wnd[0]/usr/ctxtP_WERKS").text = "5260"
            self._SAP.session.findById("wnd[0]/usr/ctxtP_STLAN").text = "3"
            self._SAP.session.findById("wnd[0]/usr/ctxtP_DATUV").text = str(date.today().strftime('%d.%m.%Y'))
            self._SAP.session.findById("wnd[0]/usr/radP_C").select()
            self._SAP.session.findById("wnd[0]/usr/chkP_C_OPT").selected = "true"
            self._SAP.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self._SAP.session.findById("wnd[0]/usr/ctxtP_LANGU").text = "PT"
            self._SAP.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self._SAP.session.findById("wnd[0]/tbar[0]/btn[86]").press()
            self._SAP.session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").text = "Barcode Printing"
            self._SAP.session.findById("wnd[1]/tbar[0]/btn[13]").press()
            
            return "Success"
        
        except:
            return "Generic Fail"