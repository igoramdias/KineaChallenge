# KineaChallenge
OBJETIVO: Repositório para desenvolvimento de um sistema de controle para o mercado secundário de debêntures

## Fluxograma e Organização de Arquivos

![Fluxograma](https://user-images.githubusercontent.com/62359383/112543931-2e380980-8d95-11eb-8b0e-47f64241d82d.png)

## Entrando no Python Virtual Env
Para usar o Python venv disponível, basta seguir os seguintes passos:


1. No CMD, coloque: path_to_KC_Project/KCvenv/Scripts/activate.bat
2. Ainda, no CMD: pip install -r requirements.txt
3. Pode rodar o programa!

Com isso, pode-se rodar o programa tranquilo, uma vez que todas as bibliotecas necessárias já estarão disponíveis.

## Operando o programa
O programa deve ser rodado de forma semanal, com o intuito de atualizar os arquivos com as transações que ocorreram na semana. Pode-se separá-lo em três partes:

### Realizando pull de dados semanais com o crawler
>Essa parte servirá para colocar o crawler para rodar e ir atrás de dados da semana para preencher as planilhas de IPCA - ANBIMA, IPCA - Mercado, CDI - ANBIMA, CDI - Mercado, %CDI - ANBIMA e %CDI - Mercado. 

Tais dados estarão nos sites:

- IMAB: https://www.anbima.com.br/pt_br/informar/ima-resultados-diarios.htm
- REUNE: https://www.anbima.com.br/pt_br/informar/sistema-reune.htm
- ETTJ: https://www.anbima.com.br/pt_br/informar/curvas-de-juros-fechamento.htm 
- TAXAS: https://www.anbima.com.br/informacoes/merc-sec-debentures/default.asp

Para realizá-la, basta inserir 'True' em:

![Pull_Dados_Week](https://user-images.githubusercontent.com/62359383/112395886-f91cb000-8cdd-11eb-904d-2bc31107d7af.png)
  
 Em sendo 'True' a resposta inserida, inputa-se as datas de início e fim:
 
 ![Pull_Dados_Week_2](https://user-images.githubusercontent.com/62359383/112395892-ff129100-8cdd-11eb-91e9-abada3f7a9ca.png)
  
 Término da operação é mercado com:
 
 ![Pull_Dados_Week_3](https://user-images.githubusercontent.com/62359383/112395903-046fdb80-8cde-11eb-9f09-73224236d282.png)

### Realizando pull de dados diário com o crawler
>Essa parte servirá para colocar o crawler para rodar e ir atrás de dados do dia para preencher as planilhas de Cadastro e Cadastro - INFRA. 

Tais dados estarão no site:

- Debêntures: http://www.debentures.com.br/exploreosnd/consultaadados/emissoesdedebentures/caracteristicas_f.asp?tip_deb=publicas

Para realizá-la, basta inserir 'True' em:

![Pull_Dados_Today](https://user-images.githubusercontent.com/62359383/112395912-0c2f8000-8cde-11eb-9635-9e03899af402.png)
  
 Em sendo 'True' a resposta inserida, realiza-se a pesquisa para o dia. Término da operação é mercado com:
 
 ![Pull_Dados_Today_2](https://user-images.githubusercontent.com/62359383/112395922-105b9d80-8cde-11eb-9e7d-8a95b8092f87.png)

### Incrementação dos dados
>Essa parte servirá para atualizar a planilha com os dados coletados. 

Sendo do interesse em realizar uma incrementação nos dados, insira 'True' em:

![Inc_Dados](https://user-images.githubusercontent.com/62359383/112395939-15b8e800-8cde-11eb-8569-c9e891dfb7d0.png)

Para tal, será necessário qual a data que deseja realizar a atualização, em:

![Inc_Dados_2](https://user-images.githubusercontent.com/62359383/112395949-1baec900-8cde-11eb-8212-8fb5e7e1e97f.png)

Caso seja necessário atualizar os dados de Rating, responda com 'True' a pergunta:

![Inc_Dados_3](https://user-images.githubusercontent.com/62359383/112395961-1fdae680-8cde-11eb-9001-9a8f85a70c00.png)
  
Programa termina com:

![Inc_Dados_4](https://user-images.githubusercontent.com/62359383/112395967-236e6d80-8cde-11eb-9b63-7fa885cda5c6.png)
  
Caso queira incrementar mais datas, basta um 'True' em:

![Inc_Dados_5](https://user-images.githubusercontent.com/62359383/112395977-2701f480-8cde-11eb-9fd6-654ed6d08125.png)
  
## Explanação das tabelas

### Cadastro
> Contém todos os Tickers disponíveis
- Ticker: Ticker da debênture
- Emissor: Companhia emissora da debênture
- Infraestrutura: Pertecence (1) ou não (0) ao interesse do time de INFRA
- Data de Saida/Nova Data de Vencimento: Data de término da debênture
- Garantia/Especie:	Tipo de garantia usada no contrato
- Valor Nominal na Emissão: Valor declarado na emissão
- Quantidade no Mercado: Quantidade disponível para negociação no mercado
- Índice: Índice ao qual a debênture é atrelada
- Percentual Multiplicador/Rentabilidade: Em relação ao índice, o quanto de rendimento
- CNPJ: CNPJ da companhia emissora
- Deb. Incentivada (Lei12.431): Possui auxílio fiscal
- Resgate Antecipado: Permite resgate antecipado
- Kinea: Ticker que está presente na carteira da Kinea atualmente

### Cadastro Infra
> Contém todos os Tickers disponíveis e que estão avaliados como de interesse do Time de Infraestrutura
- Ticker: Ticker da debênture
- Indexador: Indexador ao qual a debênture é atrelada
- Incentivada: 
- Emissor: Companhia emissora da debênture
- Setor: Setor ao qual a companhia está inserida
- Subsetor: Subsetor ao qual a companhia está inserida
- Flag Consolidado:	Define inclusão ou não do Ticker em Consolidado
- Flag Resumo: Define inclusão ou não do Ticker em Resumo
- Agência rating: Agência responsável por análise de rating da companhia	
- Rating de emissão: Rating recebido pela companhia

### Rating
> Rating de Tickers
- Dia: Dia ao qual o ticker obteve determinado rating
- Ticker: Ticker da debênture
- Agência de Rating: Agência responsável por análise de rating da companhia
- Rating: Rating recebido pela companhia
- Rating Equivalente: Normalização dos ratings

### IPCA - ANBIMA
> Dos Tickers disponíveis em Cadastro, possuem aqueles que estão fixados ao IPCA e todos que estão no cadastro da ANBIMA para aquele dia
- Ticker: Ticker da debênture
- Dia: Dia ao qual o ticker foi avaliado
- PU ANBIMA: Preço unitário da debenture
- Duration ANBIMA: Prazo em dias úteis ponderado pelos pagamentos
- Taxa ANBIMA: Correspondem às taxas avaliadas pela instituição como preço justo de negócio para cada emissão 
- NTN-B Ref ANBIMA: Data de vencimento correspondente caso o ticker fosse um NTN-B
- Taxa NTN-B Ref ANBIMA: Taxa correspondente caso o ticker fosse um NTN-B
- Spread B Ref ANBIMA: O quão esticado está o ticker em relação ao seu análogo mais seguro (NTN-B)
- Taxa ETTJ Ref ANBIMA: Taxa correspondente em relação ao mercado de juros de IPCA
- Spread ETTJ Ref ANBIMA: O quão esticado está o ticker em relação ao seu análogo no mercado de IPCA como um todo

### IPCA - Mercado
> Dos Tickers disponíveis em Cadastro, possuem aqueles que estão fixados ao IPCA e todos que foram negociados naquele dia
- Ticker: Ticker da debênture	
- Dia: Dia ao qual o ticker foi avaliado
- Volume Negociado: O quanto de debênture, em volume, foi negociado
- PU Mercado: Preço unitário da debenture
- Duration Mercado: Prazo em dias úteis ponderado pelos pagamentos
- Taxa Mercado: Correspondem às taxas de negócio para cada emissão
- NTN-B Ref Mercado: Data de vencimento correspondente caso o ticker fosse um NTN-B
- Taxa NTN-B Ref Mercado: Taxa correspondente caso o ticker fosse um NTN-B
- Spread B Ref Mercado: O quão esticado está o ticker em relação ao seu análogo mais seguro (NTN-B)
- Taxa ETTJ Ref Mercado: Taxa correspondente em relação ao mercado de juros de IPCA
- Spread ETTJ Ref Mercado: O quão esticado está o ticker em relação ao seu análogo no mercado de IPCA como um todo

### CDI - ANBIMA
> Dos Tickers disponíveis em Cadastro, possuem aqueles que estão fixados ao CDI e todos que estão no cadastro da ANBIMA para aquele dia
- Ticker: Ticker da debênture
- Dia: Dia ao qual o ticker foi avaliado
- PU ANBIMA: Preço unitário da debenture
- Duration ANBIMA: Prazo em dias úteis ponderado pelos pagamentos
- Taxa ANBIMA: Correspondem às taxas de negócio para cada emissão
- Spread ANBIMA: Correspondem às taxas de negócio para cada emissão (Mesmo valor que Taxa ANBIMA)

### CDI - Mercado
> Dos Tickers disponíveis em Cadastro, possuem aqueles que estão fixados ao CDI e todos que foram negociados naquele dia
- Ticker: Ticker da debênture
- Dia: Dia ao qual o ticker foi avaliado
- Volume Negociado: O quanto de debênture, em volume, foi negociado
- PU Mercado: Preço unitário da debenture
- Duration Mercado: Prazo em dias úteis ponderado pelos pagamentos
- Taxa Mercado: Correspondem às taxas de negócio para cada emissão
- Spread Mercado: Correspondem às taxas de negócio para cada emissão (Mesmo valor que Taxa Mercado)

### %CDI - ANBIMA
> Dos Tickers disponíveis em Cadastro, possuem aqueles que estão fixados ao %CDI e todos que estão no cadastro da ANBIMA para aquele dia
- Ticker: Ticker da debênture
- Dia: Dia ao qual o ticker foi avaliado
- Taxa Emissão: Correspondem às taxas de negócio para cada emissão no momento de emissão
- PU ANBIMA: Preço unitário da debenture
- Duration ANBIMA: Prazo em dias úteis ponderado pelos pagamentos
- Taxa ANBIMA: Correspondem às taxas de negócio para cada emissão no momento de análise
- Pré Ref ANBIMA: Taxa do Ticker em Pré Ref levando em consideração do o seu duration
- Taxa Pré Ref ANBIMA: Taxa ANBIMA descontado em 1 
- Spread Pré Ref ANBIMA: O quão esticado está o ticker em relação ao seu análogo em PRÉ  
- Taxa ETTJ Ref ANBIMA: Taxa correspondente em relação ao mercado de juros de PRÉ
- Spread ETTJ Ref ANBIMA: O quão esticado está o ticker em relação ao seu análogo no mercado de PRÉ como um todo

### %CDI - Mercado
> Dos Tickers disponíveis em Cadastro, possuem aqueles que estão fixados ao %CDI e todos que foram negociados naquele dia
- Ticker: Ticker da debênture
- Dia: Dia ao qual o ticker foi avaliado
- Taxa Emissão: Correspondem às taxas de negócio para cada emissão no momento de emissão
- Volume Negociado: O quanto de debênture, em volume, foi negociado
- PU Mercado: Preço unitário da debenture
- Duration Mercado: Prazo em dias úteis ponderado pelos pagamentos
- Taxa Mercado: Correspondem às taxas de negócio para cada emissão no momento de análise
- Pré Ref Mercado: Taxa do Ticker em Pré Ref levando em consideração do o seu duration
- Taxa Pré Ref Mercado: Taxa Mercado descontado em 1 
- Spread Pré Ref Mercado: O quão esticado está o ticker em relação ao seu análogo em PRÉ
- Taxa ETTJ Ref Mercado: Taxa correspondente em relação ao mercado de juros de PRÉ
- Spread ETTJ Ref Mercado: O quão esticado está o ticker em relação ao seu análogo no mercado de PRÉ como um todo
