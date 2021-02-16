import os
import pandas as pd

def get_path(): 
    """
        Pegar caminho atual de execução
    """ 

    return os.getcwd()

def source():
    """
        Pegar base de dados crua
    """
    
    file_path = get_path()
    downloads_path = os.path.join(os.path.expanduser("~"), "/Downloads") 
    print(downloads_path)
    IMAB = pd.read_excel(os.path.join(downloads_path, "IMA_12022021.xlsx"))
    REUNE = pd.read_excel(os.path.join(downloads_path, "REUNE_Acumulada_08022021.xlsx"))

    return IMAB, REUNE