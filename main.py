# |--------------------------------------------|
# | Kinea Challenge 2021                       |
# | Banco de Dados de Debêntures               |
# |                                            |
# |                  Igor Amâncio Machado Dias |
# |--------------------------------------------| 

import get_data as gd
import crawler as cr

if __name__ == "__main__":
    print("Iniciando programa...")

    wants_cra_week = input("Deseja realizar a busca de dados da semana? (True/False):")

    while wants_cra_week != "True" and wants_cra_week != "False":
        wants_cra_week = input("Input não válido! Responda com True ou False:")

    if wants_cra_week == "True":
        data_int = input("Insira a data de ínicio para busca (dd/mm/yy):")
        data_fin = input("Insira a data de final para busca (dd/mm/yy):")

        print("Iniciando o programa para as datas {} e {}...".format(data_int, data_fin))
        cr.crawl_data_week(data_int, data_fin)
        print("Busca de dados para as datas inseridas concluída!")

    wants_cra_td = input("Deseja realizar a busca de dados para hoje? (True/False):")

    while wants_cra_td != "True" and wants_cra_td != "False":
        wants_cra_td = input("Input não válido! Responda com True ou False:")

    if wants_cra_td == "True":
        print("Iniciando coleta de debênture ativas para o dia de hoje...")
        cr.crawl_data_today()
        print("Realização de coleta efetivada!")

    wants_incr = input("Deseja incrementar dados? (True/False):")
    
    while wants_incr != "True" and wants_incr != "False":
        wants_incr = input("Input não válido! Responda com True ou False:")

    while wants_incr == "True":
        data = input("Insira a data de referência (dd/mm/yy):")
        get_rat = input("Atualiza a sheet Rating também ? (True/False):")
        
        print("Iniciando o programa para a data: {}".format(data))
        gd.source(data, get_rat)
        print("Fim da atualização!")

        wants_incr = input("Deseja incrementar mais dados ainda? (True/False):")
    
        while wants_incr != "True" and wants_incr != "False":
            wants_incr = input("Input não válido! Responda com True ou False:")
    
    print("----")
    print("Fim do programa!")