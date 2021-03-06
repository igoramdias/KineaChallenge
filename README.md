# KineaChallenge
OBJETIVO: Repositório para desenvolvimento de um sistema de controle para o mercado secundário de debêntures

## Explanação das tabelas

### Cadastro
> Contém todos os Tickers disponíveis
- Ticker: Ticker da debênture
- Emissor: Companhia emissora da debênture
- Infraestrutura: Pertecence (1) ou não (0) ao interesse do time de INFRA
- Data de Saida/Nova Data de Vencimento: Data de término da debênture
- Garantia/Especie:	Tipo de garantia usada no contrato
- Valor Nominal na Emissão: Valor declarado na emissão
- Índice: Índice ao qual a debênture é atrelada
- Percentual Multiplicador/Rentabilidade: Em relação ao índice, o quanto de rendimento
- CNPJ: CNPJ da companhia emissora
- Deb. Incentivada (Lei12.431): Possui auxílio fiscal
- Resgate Antecipado: Permite resgate antecipado

### Cadastro Infra
> Contém todos os Tickers disponíveis e que estão avaliados como de interesse do Time de Infraestrutura
- Ticker: Ticker da debênture
- Indexador: Indexador ao qual a debênture é atrelada
- Incentivada: 
- Emissor: Companhia emissora da debênture
- Setor: Setor ao qual a companhia está inserida
- Subsetor: Subsetor ao qual a companhia está inserida
- Flag Consolidado:	
- Flag Resumo: 
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
- Pré Ref ANBIMA: Data de vencimento correspondente caso o ticker fosse um PRÉ
- Taxa Pré Ref ANBIMA: Taxa correspondente caso o ticker fosse um PRÉ 
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
- Pré Ref Mercado: Data de vencimento correspondente caso o ticker fosse um PRÉ
- Taxa Pré Ref Mercado: Taxa correspondente caso o ticker fosse um PRÉ
- Spread Pré Ref Mercado: O quão esticado está o ticker em relação ao seu análogo em PRÉ
- Taxa ETTJ Ref Mercado: Taxa correspondente em relação ao mercado de juros de PRÉ
- Spread ETTJ Ref Mercado: O quão esticado está o ticker em relação ao seu análogo no mercado de PRÉ como um todo
