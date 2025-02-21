from Path import path
from Transactions import tcodes
from Utils import utils
import pandas as pd
import subprocess
import os
import time
import shutil
import sys

def main_application ():
    
    confirm = '1'
    
    while confirm == '1':
        try:
            os.system('cls')
            op = int(input("Insira o número da OP: "))
            
            while op > 1000000 or op < 99999:
                os.system('cls')
                op = int(input("Insira o número da OP: "))

            confirm = input(f"Você inseriu o número {op}. Pressione Enter para continuar ou digite 1 para inserir outra OP: ")

        except:
            confirm = '1'
    
    os.system('cls')
    print (f"OP sendo impressa: {op}")

    df_op = pd.DataFrame([{'Document': op,'Status': None}])
    
    sap = tcodes()
    
    os.system('cls')
    print (f"OP sendo impressa: {op}")
    
    material, df_op.at[0,"Status"] = sap.Tcode_CO02(op)
    
    try:
        utils.close_excel("Documents.xlsx")
    except:
        ...
    
    #Le todas as abas do arquivo Documents.xlsx, exclui a aba Planners e concatena todas as outras na df_documents
    sheets = pd.read_excel(path.path_documents + "Documents.xlsx", sheet_name=None)
    sheets = {name: df for name, df in sheets.items() if name != "Planners"}
    df_documents = pd.concat(sheets.values(), ignore_index=True)

    df_documents = df_documents[df_documents["PN"] == material]
    
    if df_documents.empty:
        df_op = pd.concat([df_op, pd.DataFrame([{'Document': "Documentos para o PN dessa OP não foram encontrados.",'Status': "Fail all documents"}])], ignore_index=True)
        
        try:
            df_op.to_excel(path.path_output_onedrive + f"OP_{op}_Output.xlsx", index=False)
        except:
            df_op.to_excel(path.path_output_local + f"OP_{op}_Output.xlsx", index=False)
        
        sys.exit()
    
    df_documents = df_documents.reset_index(drop=True)

    df_documents = df_documents.drop(columns=[col for col in df_documents.columns if col != "Document"])
        
    df_documents["Status"] = None
        
    df_op = pd.concat([df_op, df_documents], ignore_index=True)
        
    #Deleta a pasta, caso exista
    if os.path.exists(path.path_export):
        shutil.rmtree(path.path_export)
        time.sleep(1)

    #Cria a pasta
    os.makedirs(path.path_export)
    time.sleep(1)
    
    for index, row in df_op.iterrows():
            
        os.system('cls')
        print (f"OP sendo impressa: {op}")
            
        if row['Document'] >=1e10 and row['Document']< 1e11:
            df_op.at[index, "Status"] = sap.Tcode_CV04N(row['Document'], op)
            subprocess.run (['taskkill', '/f', '/im', "Acrobat.exe"])
        elif row['Document'] >=1e7 and row['Document']< 1e8:
            df_op.at[index, "Status"] = sap.Tcode_ZZPLM_BOM(row['Document'])
    
    try:
        df_op.to_excel(path.path_output_onedrive + f"OP_{op}_Output.xlsx", index=False)
    except:
        df_op.to_excel(path.path_output_local + f"OP_{op}_Output.xlsx", index=False)

    try:
        #Deleta a pasta temp e tudo que esta dentro dela
        shutil.rmtree(path.path_export)
    except:
        ...

    os.system('cls')

main_application() 