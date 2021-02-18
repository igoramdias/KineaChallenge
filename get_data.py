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
    downloads_path = "~\Downloads"
    downloads_path = os.path.expanduser(downloads_path) 
    
    IMAB = pd.read_csv(os.path.join(downloads_path, "IMA_12022021.csv"), sep=";", header=1, encoding='ANSI')
    REUNE = pd.read_csv(os.path.join(downloads_path, "REUNE_Acumulada_17022021.csv"), sep=";", header=3, encoding='ANSI')

    return IMAB, REUNE