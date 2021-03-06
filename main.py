# |--------------------------------------------|
# | Kinea Challenge 2021                       |
# | Banco de Dados de Debêntures               |
# |                                            |
# |                  Igor Amâncio Machado Dias |
# |--------------------------------------------| 

import get_data as gd
from ast import literal_eval

if __name__ == "__main__":
    print("Iniciando programa para atualização...")

    wants_incr = literal_eval(input("Deseja incrementar dados? (True/False):"))
    
    while wants_incr is not True and wants_incr is not False:
        wants_incr = literal_eval(input("Input não válido! Responda com True ou False:"))

    if wants_incr:
        data = input("Insira a data de referência (dd/mm/yy):")
        data_RATING = gd.source(data)
        gd.fill_sheets(data_RATING)
    
    print("Fim do programa!")