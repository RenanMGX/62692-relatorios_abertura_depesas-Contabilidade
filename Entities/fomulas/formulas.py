import pandas as pd
from copy import deepcopy
from datetime import datetime
import numpy as np
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def criar_colunas_por_data(df:pd.DataFrame, datas:list, *, coluna:str='Descrição') -> pd.DataFrame:
    df_final = deepcopy(df[coluna]) #type:ignore
    df_final.drop_duplicates(inplace=True)
    
    
    for data in datas:
        df_temp:pd.DataFrame = deepcopy(df[[coluna,'Montante em moeda interna', 'Ano/Mês']])
                              
        df_temp = df_temp[
            df_temp['Ano/Mês'] == data
        ]
                                
        df_temp["Montante em moeda interna"] = pd.to_numeric(df_temp["Montante em moeda interna"], errors="coerce")
                
        df_temp = df_temp.groupby([coluna], as_index=False).sum().round(2)
        df_temp.drop(columns=['Ano/Mês'], inplace=True, errors='ignore')

        df_temp.rename(columns={'Montante em moeda interna': data}, inplace=True)
                
        df_final:pd.DataFrame = pd.merge(df_temp, df_final, how='outer', on=coluna).fillna(0).round(2)
    
    return df_final

def __criar_sub_total(df: pd.DataFrame, *, coluna_primaria:str, colunas:list) -> pd.DataFrame:
    
    # Calcula o subtotal das colunas numéricas
    subtotal = df[colunas].sum()
    subtotal[coluna_primaria] = 'Subtotal'

    # Reordena as colunas para ficar na mesma ordem do DataFrame original
    subtotal = subtotal.reindex(df.columns)

    # Concatena a linha de subtotal no final do DataFrame
    return pd.concat([df, subtotal.to_frame().T], ignore_index=True)

def criar_acumulado(df: pd.DataFrame, datas:list, *, coluna:str) -> pd.DataFrame:
    df_temp = deepcopy(df)
    #import pdb; pdb.set_trace()
    subtotal = df_temp.sum()
    subtotal[coluna] = 'Total Despesas'
    subtotal = subtotal.reindex(df.columns)
    
    subtotal = pd.concat([df_temp, subtotal.to_frame().T], ignore_index=True)

    subtotal = criar_colunas_calculadas(subtotal, datas)
    
    return subtotal[subtotal[coluna] == "Total Despesas"]
    
def criar_colunas_calculadas(
        df_final:pd.DataFrame,
        _datas:list,
        *,
        sub_total:bool=False,
        somar_ultima_coluna:bool=False,
        ) -> pd.DataFrame:
    
    
    datas = deepcopy(_datas) #type:ignore       
    datas.reverse() #type:ignore
    
    df_final['Total'] = df_final[datas].sum(axis=1).round(2) #type:ignore
            
    total = df_final['Total'].sum()#.round(2)
            
    #df_final['Repres.%'] = df_final.apply(lambda x: round(x['Total']/ total, 4), axis=1)
    df_final['Repres.%'] = divisao_entre_coluna_valor_total(df_final['Total'], total, _round=4)
    ultimo_mes_extenso = datetime.strptime(datas[-1], '%Y/%m').strftime('%b/%y')
            
    nome_coluna_r = f"V.H R$ {ultimo_mes_extenso.title()} - (2m - µ)"
    if len(datas) == 1:
        df_final[nome_coluna_r] = df_final[datas[0]].round(2)
    elif len(datas) == 2:
        df_final[nome_coluna_r] = round(df_final[datas[1]] - df_final[datas[0]], 2)
    elif len(datas) >= 3:
        df_final[nome_coluna_r] = round(df_final[datas[-1]] - df_final[datas[0:-1]].mean(axis=1), 2) #type:ignore
    else:
            df_final[nome_coluna_r] = 0     
                
                
    nome_coluna_p = f"V.H % {ultimo_mes_extenso.title()} - (2m - µ)"    
    
    
    if sub_total:
        if somar_ultima_coluna:
            df_final[nome_coluna_p] = divisao_entre_duas_colunas(df_final[nome_coluna_r], df_final[datas[-1]])
        colunas = []
        colunas += datas
        if somar_ultima_coluna:
            colunas += ['Total', 'Repres.%', nome_coluna_r, nome_coluna_p]
        else:
            colunas += ['Total', 'Repres.%', nome_coluna_r]
        
        df_final = __criar_sub_total(df_final,
                                     coluna_primaria='Descrição',
                                     colunas=colunas
                                     )
        
    if not somar_ultima_coluna:
        df_final[nome_coluna_p] = divisao_entre_duas_colunas(df_final[nome_coluna_r], df_final[datas[-1]])
    
    return df_final

def ordenar(
    df: pd.DataFrame,
    *,
    coluna:str,
    ordem:list
    ) -> pd.DataFrame:# ordem desejada
    
    df[coluna] = pd.Categorical(df[coluna], categories=ordem, ordered=True)
    df = df.sort_values(coluna) #type:ignore
    
    return df


def divisao_entre_duas_colunas(df1:pd.Series, df2:pd.Series, *, _round:int=2) -> list:
    numerador = df1.tolist()
    denominador = df2.tolist()
    
    if len(numerador) != len(denominador):
        raise ValueError("As listas devem ter o mesmo tamanho.")
    
    resultado = []
    for num in range(len(numerador)):
        try:
            resultado.append(round(numerador[num] / denominador[num], _round))
        except:
            resultado.append(0)
            
    return resultado

def divisao_entre_coluna_valor_total(df1:pd.Series, valor:int|float, *, _round:int=2) -> list:
    numerador = df1.tolist()
    valor = valor
        
    resultado = []
    for num in range(len(numerador)):
        try:
            if (valor != 0) and (numerador[num] != 0):
                resultado.append(round(numerador[num] / valor, _round))
            else:
                resultado.append(0)
        except:
            resultado.append(0)
            
    return resultado

