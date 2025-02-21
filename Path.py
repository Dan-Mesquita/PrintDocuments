import os 

class path(): 
    if os.getcwd().split(os.sep)[-1] == "Phyton":
        path_documents = os.sep.join(os.getcwd().split(os.sep)[:-1]) + os.sep + "Database" + os.sep
    else:
        path_documents = os.getcwd() + os.sep + "Database" + os.sep
    
    path_export = os.getcwd().split(os.sep)[0] + os.sep + os.getcwd().split(os.sep)[1] + os.sep + os.getcwd().split(os.sep)[2] + "\\Downloads\\Temp\\"

    path_output_onedrive = os.getcwd().split(os.sep)[0] + os.sep + os.getcwd().split(os.sep)[1] + os.sep + os.getcwd().split(os.sep)[2]+ "\\OneDrive - SLB\\Desktop\\"
    
    path_output_local = os.getcwd().split(os.sep)[0] + os.sep + os.getcwd().split(os.sep)[1] + os.sep + os.getcwd().split(os.sep)[2]+ "\\Desktop\\"
    