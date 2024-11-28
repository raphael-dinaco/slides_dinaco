import pandas as pd
import sys
sys.path.append("..")

def tratamento_faturamento():
    df = pd.read_csv('SQL/Outputs/faturamento_meta_carteira.csv',sep=',', low_memory=False)
    df = df.rename(columns={'UN_NEG':'UNIDADENEG',
                                            'GRUPO_ECONOMICO':'GRUPOECONOMICO',
                                                'VENDEDOR_CAB':'EXECUTIVO',
                                                'PRODUTO': 'DESCRPROD',
                                                'GRUPO_PRODUTO': 'DESCRGRUPOPROD'})
    df.drop(columns=['Unnamed: 0'], inplace=True)
    df.loc[df['TIPMOV'] == 'D-Devolução de venda', 'TIPMOV'] = 'V-Venda'
    df = df[df['TIPMOV']=='V-Venda']
    df['Mês/Ano'] = df['MES'].astype(int).astype(str) + '/' + df['ANO'].astype(int).astype(str).str[-2:]
    df['Mês/Ano']= pd.to_datetime(df['Mês/Ano'], format='%m/%y')
    df['ANO'] = df['ANO'].astype(int)
    return df

def tratamento_meta():
    df = pd.read_csv('SQL/Outputs/faturamento_meta_carteira.csv',sep=',', low_memory=False)
    df = df.rename(columns={'UN_NEG':'UNIDADENEG',
                                            'GRUPO_ECONOMICO':'GRUPOECONOMICO',
                                                'VENDEDOR_CAB':'EXECUTIVO',
                                                'PRODUTO': 'DESCRPROD',
                                                'GRUPO_PRODUTO': 'DESCRGRUPOPROD'})
    df.drop(columns=['Unnamed: 0'], inplace=True)
    df = df[df['TIPMOV']=='Meta']
    df['Mês/Ano'] = df['MES'].astype(int).astype(str) + '/' + df['ANO'].astype(int).astype(str).str[-2:]
    df['Mês/Ano']= pd.to_datetime(df['Mês/Ano'], format='%m/%y')
    df['ANO'] = df['ANO'].astype(int)
    return df 

def tratamento_oportunidade():
    df_oportunidade = pd.read_csv('SQL/Outputs/Oportunidades.csv', sep=',', low_memory=False)
    df_oportunidade.drop(columns=['Unnamed: 0'], inplace=True)
    df_oportunidade['DTNEG'] = pd.to_datetime(df_oportunidade['DTNEG'])
    df_oportunidade['DTESTFECHAMENTO'] = pd.to_datetime(df_oportunidade['DTESTFECHAMENTO'])
    df_oportunidade['DTFECHAMENTO'] = pd.to_datetime(df_oportunidade['DTFECHAMENTO'])
    df_oportunidade['DT_FATUR_AMOSTRA'] = pd.to_datetime(df_oportunidade['DT_FATUR_AMOSTRA'])
    df_oportunidade['NOME_PROJETO'] = df_oportunidade['NOME_PROJETO'].str.upper()
    df_oportunidade['NOME_PROJETO'] = df_oportunidade['NOME_PROJETO'].map(lambda x: str(x)[:-1] if str(x).endswith(' ') else x)
    df_oportunidade['NOME_PROJETO'] = df_oportunidade['NOME_PROJETO'].fillna('SEM NOME')

    # Remoção de oportunidades de grupos mortos
    grupo_morto = ['01 - LUBRIZOL','02 - HONEYWELL','10 - HALL STAR','13B - HALLSTAR INDUSTRIAL','14 - LONGCHEM','17 - JOS.H.LOWENSTEIN','19B - SCHULKE INDUSTRIAL','23 - BLUESTAR','28 - POLY ONE','42 - OXEN',
               '42B - OXEN COSMETICOS','45 - NANOX','54 - OCO','56 - NOVACHEM','60 - BOTANECO','62 - RHODIA','66 - ASSESSA','67 - ALFA CHEMICALS','69 - HPF','74 - LIPOTEC']

    df_oportunidade = df_oportunidade[~df_oportunidade['GRUPO_PRODUTO'].isin(grupo_morto)]
    df_oportunidade = df_oportunidade[~df_oportunidade['GRUPO_PRODUTO'].isnull()].sort_values(by='NUNEGOCIACAO')

    # Remoção de Substituto
    df_oportunidade = df_oportunidade[df_oportunidade['PRODUTO_SUBSTITUTO']=='Não'].sort_values(by='DTNEG', ascending=True)

    # Remoção de Laboratório de Inovação e de Márcia Coelho
    df_oportunidade = df_oportunidade[(df_oportunidade['GRUPO_ECONOMICO']!='Laboratório de Inovação')&(df_oportunidade['VENDEDOR_PAR']!='MARCIA.COELHO')]

    # Remoção de oportunidades de Consultor
    df_oportunidade = df_oportunidade[df_oportunidade['AD_CONSULTORUNI'] != 'Sim']

    # Remoção Regra 1 - Mesma oportunidade, uma cancelada
    duplicados_cancelado = df_oportunidade[df_oportunidade.duplicated(subset=['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'], keep=False)].groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO']).filter(lambda x: (x['STATUS'].nunique() > 1))
    duplicados_cancelado = duplicados_cancelado[duplicados_cancelado['STATUS']=='Cancelado']
    df_oportunidade = df_oportunidade[~df_oportunidade.index.isin(duplicados_cancelado.index)]

    # Remoção Regra 2 - Mesma oportunidade, uma delas com valor zerado
    duplicados_valor = df_oportunidade[df_oportunidade.duplicated(subset=['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'], keep=False)].groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO']).filter(lambda x: (x['VLRTOT'].nunique() > 1))
    duplicados_valor = duplicados_valor[duplicados_valor['VLRTOT']==0]
    df_oportunidade = df_oportunidade[~df_oportunidade.index.isin(duplicados_valor.index)]

    # Remoção Regra 3 - Criação de ordem de prioridade por status (Mesma oportunidade, fica a com maior prioridade)
    prioridade = {'STATUS': ['Cancelado', 'Reprovado', 'Faturado', 'Início de Projeto','Aprovado', 'Negociação', 'Teste em laboratório/formulação'
                        ,'Teste de estabilidade', 'Teste não iniciado','Teste no consumidor', 'Teste piloto', '']
                        , 'Prioridade': [0,1,10,2,9,8,4,5,3,7,6,0],
                        'STATUS RESUMIDO': ['Cancelado','Reprovado','Faturado','Em Andamento','Em Andamento','Em Andamento','Em Andamento','Em Andamento','Em Andamento','Em Andamento','Em Andamento','Em Andamento']}
    prioridade = pd.DataFrame(prioridade)
    prioridade = prioridade.sort_values(by='Prioridade', ascending=True)
    '''
    ('Cancelado'(0), 'Reprovado'(8), 'Faturado'(10), 'Início de Projeto'(1),
        'Aprovado'(9), 'Negociação'(7), 'Teste em laboratório/formulação'(3),
        'Teste de estabilidade'(4), 'Teste não iniciado'(2),
        'Teste no consumidor'(5), 'Teste piloto'(6))'''
    df_oportunidade = pd.merge(df_oportunidade, prioridade, on='STATUS', how='left')

    duplicados_status = df_oportunidade[df_oportunidade.duplicated(subset=['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'], keep=False)].groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO']).filter(lambda x: (x['Prioridade'].nunique() > 1)).sort_values(by=['Prioridade'], ascending=False)
    duplicados_status_out = duplicados_status[duplicados_status.groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'])['Prioridade'].transform(max) == duplicados_status['Prioridade']]
    duplicados_status = duplicados_status[~duplicados_status.index.isin(duplicados_status_out.index)]
    df_oportunidade = df_oportunidade[~df_oportunidade.index.isin(duplicados_status.index)]

    # Remoção Regra 4 - Oportunidades iguais, valores diferentes
    duplicados_menor_valor = df_oportunidade[df_oportunidade.duplicated(subset=['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'], keep=False)].groupby(['GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO']).filter(lambda x: (x['VLRTOT'].nunique() > 1)).sort_values(by=['VLRTOT'], ascending=False)
    duplicados_menor_valor_out = duplicados_menor_valor[duplicados_menor_valor.groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'])['VLRTOT'].transform(max) == duplicados_menor_valor['VLRTOT']]
    duplicados_menor_valor = duplicados_menor_valor[~duplicados_menor_valor.index.isin(duplicados_menor_valor_out.index)]
    df_oportunidade = df_oportunidade[~df_oportunidade.index.isin(duplicados_menor_valor.index)]

    # Remoção Regra 5 - oportunidades iguais, datas de criação diferentes
    duplicados_data = df_oportunidade[df_oportunidade.duplicated(subset=['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'], keep=False)].groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO']).filter(lambda x: (x['DTNEG'].nunique() > 1)).sort_values(by=['DTNEG'], ascending=False)
    duplicados_data_out = duplicados_data[duplicados_data.groupby(['NOME_PROJETO', 'GRUPO_ECONIMICO_FINAL', 'GRUPO_PRODUTO'])['DTNEG'].transform(max) == duplicados_data['DTNEG']]
    duplicados_data = duplicados_data[~duplicados_data.index.isin(duplicados_data_out.index)]
    df_oportunidade = df_oportunidade[~df_oportunidade.index.isin(duplicados_data.index)].sort_values(by='NUNEGOCIACAO', ascending=True)

    #Remoção Regra 6 - o resto (?)
    df_oportunidade = df_oportunidade.drop_duplicates(subset=['NOME_PROJETO','GRUPO_ECONIMICO_FINAL','GRUPO_PRODUTO'],keep = 'first')

    df_oportunidade['STATUS'] = df_oportunidade['STATUS'].str.replace('Aprovado','Approved'
                                                                      ).str.replace('Início de Projeto','First Steps'
                                                                                    ).str.replace('Teste não iniciado','First Steps'
                                                                                                  ).str.replace('Teste de estabilidade','Stability'
                                                                                                                ).str.replace('Teste em laboratório/formulação','Formulations Test'
                                                                                                                              ).str.replace('Teste no consumidor','Final Tests'
                                                                                                                                            ).str.replace('Teste piloto','Final Tests'
                                                                                                                                                          ).str.replace('Negociação','Negociation'
                                                                                                                                                                        ).str.replace('Faturado','Converted'
                                                                                                                                                                                      ).str.replace('Cancelado','Canceled'
                                                                                                                                                                                                    ).str.replace('Reprovado','Reproved')

    return df_oportunidade

def tratamento_compras():
    df_compras = pd.read_csv('SQL/Outputs/compras.csv', low_memory=False)
    df_compras['DT_REL'] = pd.to_datetime(df_compras['DT_REL'])
    df_compras = df_compras[['TIPMOV','GRUPO','DT_REL','CODPROD','GRUPO_PRODUTO','QTD_NEG','VLRUNITDOLAR','VLRTOT_DOLAR','REPRESENTADA']]
    return df_compras

def tabela_compras():
    df_compras = pd.read_csv('SQL/Outputs/compras.csv')
    df_compras['DT_REL'] = pd.to_datetime(df_compras['DT_REL'])
    df_compras = df_compras[['TIPMOV','GRUPO','DT_REL','CODPROD','GRUPO_PRODUTO','QTD_NEG','VLRUNITDOLAR','VLRTOT_DOLAR','REPRESENTADA']]
    tabela_compras = df_compras.sort_values(by='DT_REL', ascending = False)[['GRUPO_PRODUTO','VLRUNITDOLAR']].drop_duplicates(subset='GRUPO_PRODUTO',keep = 'first')
    return tabela_compras

def tabela_representadas():
    df_representadas = pd.read_csv('SQL/Outputs/representadas.csv')
    return df_representadas

def tratamento_visitas():
    df_visitas = pd.read_csv('SQL/Outputs/visitas.csv')
    df_visitas.drop(columns=['Unnamed: 0'], inplace=True)
    df_visitas['AD_DHVISITA'] = pd.to_datetime(df_visitas['AD_DHVISITA'])
    df_visitas['DESCRHIST'] = df_visitas['DESCRHIST'].str.replace('ANÁLISE DE CLIENTE','CALL').str.replace('FEIRA','VISITA').str.replace('WORKSHOP','VISITA')
    return df_visitas