# |--------------------------------------------|
# | Kinea Challenge 2021                       |
# | Banco de Dados de Debêntures               |
# |                                            |
# |                  Igor Amâncio Machado Dias |
# |--------------------------------------------| 

import get_data as gd
import os
import pandas as pd

if __name__ == "__main__":
    print("Iniciando programa para atualização...")
    
    dataIMAB, dataREUNE, dataIPCA, dataDI = gd.source()
    dataDEBENT = gd.data_base()
    
    