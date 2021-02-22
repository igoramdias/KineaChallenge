import os
import pandas as pd
import openpyxl as opxl
from openpyxl.utils.dataframe import dataframe_to_rows

global dict

dict = {'IPCA - ANBIMA': 'REUNE', 
        'Columns': {
            'CETIP': 'Ticker', 
            'Dia': 'Dia', #Criar uma coluna na REUNE com a data
            '': 'PU ANBIMA', 
            '': 'Duration ANBIMA', 
            '': 'Taxa ANBIMA', 
            '': 'NTN-B Ref ANBIMA', 
            '': 'Taxa NTN-B Ref ANBIMA', 
            '': 'Spread B Ref ANBIMA', 
            '': 'Taxa ETTJ Ref ANBIMA',
            '': 'Spread ETTJ Ref ANBIMA'
        },
        'IPCA - MERCADO': 'TAXAS - IPCA', 
        'Columns': {
            'Código': 'Ticker', 
            'Repac./  Venc.': 'Dia',
            '': 'Volume Negociado', 
            '': 'PU Mercado',
            '': 'Duration Mercado', 
            '': 'Taxa Mercado', 
            'Referência NTN-B': 'NTN-B Ref Mercado', 
            '': 'Taxa NTN-B Ref Mercado', 
            '': 'Spread B Ref Mercado',
            '': 'Taxa ETTJ Ref Mercado', 
            '': 'Spread ETTJ Ref Mercado'
        }, 
        'CDI - ANBIMA': 'IMAB', 
        'Colunas': {
            '': 'Ticker', 
            '': 'Dia', 
            '': 'PU ANBIMA', 
            '': 'Duration ANBIMA', 
            '': 'Taxa ANBIMA', 
            '': 'Spread ANBIMA'
        }, 
        'CDI - MERCADO': 'TAXAS - DI', 
        'Colunas': {
            '': 'Ticker', 
            '': 'Dia', 
            '': 'Volume Negociado', 
            '': 'PU Mercado', 
            '': 'Duration Mercado', 
            '': 'Taxa Mercado', 
            '': 'Spread Mercado'
        }, 
        '%CDI - ANBIMA': {
            '': 'Ticker', 
            '': 'Dia', 
            '': 'Taxa Emissão', 
            '': 'PU ANBIMA', 
            '': 'Duration ANBIMA', 
            '': 'Taxa ANBIMA', 
            '': 'Pré Ref ANBIMA', 
            '': 'Taxa Pré Ref ANBIMA', 
            '': 'Spread Pré Ref ANBIMA', 
            '': 'Taxa ETTJ Ref ANBIMA', 
            '': 'Spread ETTJ Ref ANBIMA'
        },
        '%CDI - MERCADO': {
            '': 'Ticker', 
            '': 'Dia', 
            '': 'Taxa Emissão', 
            '': 'Volume Negociado', 
            '': 'PU Mercado', 
            '': 'Duration Mercado', 
            '': 'Taxa Mercado', 
            '': 'Pré Ref Mercado', 
            '': 'Taxa Pré Ref Mercado', 
            '': 'Spread Pré Ref Mercado', 
            '': 'Taxa ETTJ Ref Mercado', 
            '': 'Spread ETTJ Ref Mercado'
        }
    }

										
											
			
	
						 
		

def write_in_sheet(sheet, df, header, dict):
    """
        Função para preenchimento das células
    """

    row_str = sheet.max_row+1
    for row in dataframe_to_rows(df, index = False, header = header):
        coluna = 0
        for celula in row:
            sheet.cell(row = row_str, column = coluna).value = celula
            coluna += 1
        row_str += 1

def clean(df, type):
    """
        Realizar pré-processamentos para bases a serem exploradas
    """
    if (type == 'IPCA') | (type == 'DI'):
        col_1 = list(df.columns)[0]
        idx_srt = list(df.index[df[col_1] == 'Código'])[0]
        idx_end = list(df.index[df[col_1].astype(str).str.contains("[*]")])[0]
        columns = pd.Series(list(df.iloc[idx_srt]))
        col_intind = list(columns[columns.astype(str).str.contains('Intervalo Indicativo')])
        if len(col_intind) == 1:
            columns = columns.fillna('Intervalo Indicativo Máxima')
            columns = columns.replace(col_intind[0], "Intervalo Indicativo Mínimo")
        df.columns = columns
        df = df.iloc[idx_srt+2:idx_end]
        df = df.reset_index(drop=True)
        df['Nome'] = [name[:name.index('*')-2] if '*' in name else name for name in df['Nome']]
        return df

    if type == 'REUNE':
        df = df.reset_index()
        col_1 = list(df.columns)[0]
        idx_srt = list(df.index[df[col_1].astype(str).str.contains('CETIP')])[0]
        columns = pd.Series(list(df.iloc[idx_srt]))
        df.columns = columns
        df = df.iloc[idx_srt+1:]
        df = df.reset_index(drop=True)
        df = df[df['Agrupamento'] != 'SUBTOTAL']
        return df

    if type == 'IMAB':
        df = df.reset_index()
        col_1 = list(df.columns)[0]
        idx_srt = list(df.index[df[col_1].astype(str).str.contains('Data')])[0]
        columns = pd.Series(list(df.iloc[idx_srt]))
        columns = [col[:col.index('*')] if '*' in col else col for col in columns]
        df.columns = columns
        df = df.iloc[idx_srt+1:]
        df = df.reset_index(drop=True)  
        return df

def data_base():
    """
        Carregamento da base de dados já existente
    """

    global downloads_path

    df = opxl.load_workbook(os.path.join(downloads_path, "BD Infra - Secundário v03.xlsx"))['Cadastro']
    df = pd.DataFrame(df.values, columns = [header.value for header in df[1]]).iloc[1:]
    
    return df

def source():
    """
        Pegar base de dados crua
    """
    
    global downloads_path 

    downloads_path = "~\Downloads"
    downloads_path = os.path.expanduser(downloads_path) 
    
    IMAB = pd.read_csv(os.path.join(downloads_path, "IMA_12022021.csv"), sep=";", encoding='ANSI')
    REUNE = pd.read_csv(os.path.join(downloads_path, "REUNE_Acumulada_17022021.csv"), sep=";", header=2, encoding='ANSI')
    TAXAS_IPCA = pd.read_excel(os.path.join(downloads_path, "d21fev17.xls"), sheet_name='IPCA_SPREAD')
    TAXAS_DI = pd.read_excel(os.path.join(downloads_path, "d21fev17.xls"), sheet_name='DI_SPREAD')

    IMAB = clean(IMAB, 'IMAB')
    REUNE = clean(REUNE, 'REUNE')
    TAXAS_IPCA = clean(TAXAS_IPCA, 'IPCA')
    TAXAS_DI = clean(TAXAS_DI, 'DI')

    return IMAB, REUNE, TAXAS_IPCA, TAXAS_DI