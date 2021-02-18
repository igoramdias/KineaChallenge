import get_data as gd
import os
import pandas as pd

if __name__ == "__main__":
    print("\
        |--------------------------------------------|\n\
        | Kinea Challenge 2021                       |\n\
        | Banco de Dados de Debêntures               |\n\
        |                                            |\n\
        |                  Igor Amâncio Machado Dias |\n\
        |--------------------------------------------|\n\
            Iniciando programa para atualização... \n")
    
    dataIMAB, dataREUNE = gd.source()
    
    print(dataREUNE) 
    