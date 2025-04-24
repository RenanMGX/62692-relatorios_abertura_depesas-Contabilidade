import pandas as pd
import xlwings as xw
import numpy as np
from xlwings.main import Sheet, Book
from .dependencies.functions import Functions, P
from time import sleep
import os
from typing import Union
from copy import deepcopy
from datetime import datetime
from .fomulas import formulas
import re

class AlimentarBase:
    @property
    def file_path(self) -> str:
        return self.__file_path
    
    def __init__(self, file_path: str, *, sheet_name: str = 'Base'):
        print(P("Iniciando AlimentarBase", color='cyan'))
        self.valid_path(file_path)
        self.__file_path = file_path
        self.__sheet_name = sheet_name
    
    @staticmethod
    def valid_path(path:str) -> bool:
        if not isinstance(path, str):
            raise TypeError(f"file_path must be a string, got {type(path)}")
        if not path.endswith(('.xlsx', '.xlsm', '.xls')):
            raise ValueError(f"file_path must be an xlsx file, got {path}")
        return True
        
    def __read_df(self, file_path:str) -> pd.DataFrame:
        print(P("Lendo arquivo df", color='yellow'))
        self.valid_path(file_path)
        df = pd.read_excel(file_path)
        df = df[
            ~df["Ano/Mês"].isna()
        ]
        df.rename(columns={"Divisão": "Divisão 1"}, inplace=True)
        df_columns = df.columns.tolist()
        df['Divisão'] = ""
        df['Conta 1'] = ""
        df_columns.insert(2, "Divisão")
        df_columns.insert(4, "Conta 1")
        df = df[df_columns]
        df = df.astype({
            "Conta": int,
            "Conta lnçto.contrap.": int,
            "Nº documento": int,
            "Chave de lançamento": int
            })
        
        return df     
        
        
    
    def add(self, file_path:str):
        df_append = self.__read_df(file_path)
        
        print(P("Lendo arquivo Excel", color='yellow'))
        app = xw.App(visible=False)
        with app.books.open(self.file_path) as wb:
            sheet:Sheet = wb.sheets[self.__sheet_name]
            sheet.api.AutoFilterMode = False
            
            print(P("Lendo dados da planilha", color='yellow'))
            tabela = sheet.range('A1').expand('table').value
            if isinstance(tabela[0], str):
                tabela = [tabela]
            df_temp = pd.DataFrame(tabela)
            df_temp.columns = df_temp.iloc[0]
            df_temp = df_temp[1:]
            if not df_temp.empty:
                df_temp['Divisão'] = ""
                df_temp['Conta 1'] = ""
            
            print(P("Tratando dados da planilha", color='yellow'))
            df_merge = pd.concat([df_temp, df_append], ignore_index=True)
            df_merge = df_merge.astype(
                {   
                    "Divisão 1": str,
                    "Conta": int,
                    "Conta lnçto.contrap.": int,
                    "Nº documento": int,
                    "Chave de lançamento": int,
                    "Atribuição": str                
                }
                )
            df_merge.drop_duplicates(inplace=True)
            df_merge = df_merge.sort_values(by='Ano/Mês', ascending=True)            
            
            divisao_formula:str = sheet.range('C2').formula
            if not divisao_formula == '':
                df_merge['Divisão'] = [divisao_formula.replace("B2", f"B{n}") for n in range(2, (len(df_merge)+2))]
            else:
                df_merge['Divisão'] = [f'=VLOOKUP(B{n},Dados!D:E,2,0)' for n in range(2, (len(df_merge)+2))]
            
            conta_formula:str = sheet.range('E2').formula
            if not conta_formula == '':
                df_merge['Conta 1'] = [conta_formula.replace("D2", f"D{n}") for n in range(2, (len(df_merge)+2))]            
            else:
                df_merge['Conta 1'] = [f'=VLOOKUP(D{n},Dados!A:B,2,0)' for n in range(2, (len(df_merge)+2))]           
            
            print(P("Limpando planilha", color='yellow'))
            sheet.range('A2:R2').expand('down').delete()
            
            print(P("Escrevendo dados na planilha", color='yellow'))
            sheet.range("B:B").api.NumberFormat = "@"
            sheet.range("O:O").api.NumberFormat = "@"
            sheet.range("F:H").api.NumberFormat = "@"
            sheet.range("M:N").api.NumberFormat = "@"
            sheet.range('A2').expand('table').value = df_merge.values.tolist() 
            
            print(P("salvando planilha", color='yellow'))
            wb.save()
        
        app.kill()
        try:
            del app
        except:
            pass
        sleep(5)
        Functions.fechar_excel(self.file_path)
        
        print(P("Planilha alimentada com sucesso!", color='green'))
        return self.file_path
        

class Tabela:
    @property
    def file_path(self) -> str:
        return self.__file_path
    
    def __init__(self, file_path: str, *, sheet_name: str = 'Base'):
        print(P(f"Iniciando Tabela do arquivo {os.path.basename(file_path)}", color='cyan'))
        AlimentarBase.valid_path(file_path)
        self.__file_path = file_path
        self.__sheet_name = sheet_name

    def criar_adm(self) -> bool:
        Functions.fechar_excel(self.file_path)
        app = xw.App(visible=False)
        with app.books.open(self.file_path) as wb:
            sheet:Sheet = wb.sheets[self.__sheet_name]
            sheet.api.AutoFilterMode = False
            
            dados = sheet.range("A1").expand('table').value
            
            df:pd.DataFrame = pd.DataFrame(dados[1:], columns=dados[0])
            
            mask = (df["Conta 1"] != "5302010010 REEMBOLSO DESPESAS") | (
                (df["Conta 1"] == "5302010010 REEMBOLSO DESPESAS") & (df["Divisão 1"] == "0001")
            )
            
            df = df[mask]
            
            try:
                ws_desp = wb.sheets.add('Despesas')
            except ValueError:
                wb.sheets['Despesas'].delete()
                ws_desp = wb.sheets.add('Despesas')
            
            datas:list = df['Ano/Mês'].unique().tolist() #type:ignore
            if isinstance(datas, list):
                datas.reverse()
            
            #df_final = deepcopy(df['Conta 1']) #type:ignore
            #df_final.drop_duplicates(inplace=True)
            
            df.rename(columns={'Conta 1': 'Descrição'}, inplace=True)
            
            df_final = df[df["Descrição"] != "5302010010 REEMBOLSO DESPESAS"]
            
            df_final = formulas.criar_colunas_por_data(df_final, datas)
            df_final = formulas.criar_colunas_calculadas(df_final, datas, sub_total=True)
            df_final = formulas.ordenar(df_final, coluna='Descrição',
                                        ordem= [
                                            "5302010000 SALARIOS, ENCARGOS E BENEFICIOS",
                                            "5302010003 CONSULTORIAS E ASSESSORIAS",
                                            "5302010006 DESPESAS GERAIS",
                                            "5302010007 DEPRECIACAO E AMORTIZACOES",
                                            "5302010001 SERVICOS DE TERCEIROS",
                                            "5302010002 SERVIÇOS JURÍDICOS",
                                            "5302010004 UTILIDADES",
                                            "Subtotal",
                                        ]
                                        )
            
            df_reem = df[df["Descrição"] == "5302010010 REEMBOLSO DESPESAS"]
            df_reem = formulas.criar_colunas_por_data(df_reem, datas)
            df_reem = formulas.criar_colunas_calculadas(df_reem, datas)
            
            df_subtotal = df_final[df_final['Descrição'] == 'Subtotal']
            df_subtotal = pd.concat([df_subtotal, df_reem])
            df_subtotal = formulas.criar_acumulado(df_subtotal, datas, coluna='Descrição')

            ws_desp.range("A1").value = df_final.columns.tolist()
            
            ws_desp.range("A1").expand("right").api.Interior.Color = 12611584.0
            ws_desp.range("A1").expand("right").api.Font.Color = 16777215.0
            ws_desp.range("A1").expand("right").api.Font.Bold = True
            
            ws_desp.range("A1").expand("table").api.Columns.AutoFit()
            
            ws_desp.range("A2").value = df_final.values.tolist() 
            
            cell_addres = ws_desp.range("A1").expand('down').address
            if (cell:=re.search(r'(?<=:[$]A[$])[0-9]+', cell_addres)):
                num_row = int(cell.group()) + 2
                ws_desp.range(f"A{num_row}").value = df_reem.values   
                ws_desp.range(f"A{num_row + 2}").value = df_subtotal.values  
                
                row1 = ws_desp.range("A1").expand('right')
                for cell in row1:
                    coluna:str = cell.value
                    endere:str = cell.address
                    if (letra:=re.search(r'[$][A-z]+', endere)):
                        letra = letra.group()
                        if "Repres.%" == coluna:       
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "0%"
                        elif 'V.H %' in coluna:
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "0%"
                        elif 'Descrição' == coluna:
                            continue
                        else:
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "#.##0;(#.##0)"            
                
                cell = ws_desp.range("A1").expand('right').address
                if (address:=re.search(r'(?<=:)[$][A-z]+', cell)):
                    address = address.group()
                    ws_desp.range(f'$A:{address}').api.Columns.AutoFit()
                    ws_desp.range(f'$A$1:{address}${num_row + 2}').api.Borders.LineStyle = -4118
               
            wb.save()
                
        app.kill()
        try:
            del app
        except:
            pass
        sleep(5)
        Functions.fechar_excel(self.file_path)
        return True
        #df_final.to_excel('teste.xlsx', index=False)

    def criar_comercial(self) -> bool:
        Functions.fechar_excel(self.file_path)
        app = xw.App(visible=False)
        with app.books.open(self.file_path) as wb:
            sheet:Sheet = wb.sheets[self.__sheet_name]
            sheet.api.AutoFilterMode = False
            
            dados = sheet.range("A1").expand('table').value
            
            df:pd.DataFrame = pd.DataFrame(dados[1:], columns=dados[0])
                        
            try:
                ws_desp = wb.sheets.add('Despesas')
            except ValueError:
                wb.sheets['Despesas'].delete()
                ws_desp = wb.sheets.add('Despesas')
            
            datas:list = df['Ano/Mês'].unique().tolist() #type:ignore
            if isinstance(datas, list):
                datas.reverse()
            
            #df_final = deepcopy(df['Conta 1']) #type:ignore
            #df_final.drop_duplicates(inplace=True)
            
            df.rename(columns={'Conta 1': 'Descrição'}, inplace=True)
            
            df_final = df
            
            df_final = formulas.criar_colunas_por_data(df_final, datas)
            df_final = formulas.criar_colunas_calculadas(df_final, datas, sub_total=True, somar_ultima_coluna=True)
            df_final = formulas.ordenar(df_final, coluna='Descrição',
                                        ordem= [
                                            "5301010002 COMISSOES E CORRETAGENS",
                                            "5301010007 DEPRECIACAO E AMORTIZACOES",
                                            "5301010001 PROPAGANDA E PUBLICIDADE",
                                            "5301010008 DESPESAS GERAIS",
                                            "5301010000 SALARIOS, ENCARGOS E BENEFICIOS",
                                            "5301010005 DESPESAS GERAIS-CENTRAL DE VENDAS",
                                            "5301010010 BRINDES",
                                            "5301010009 PROMOCAO COMERCIAL",
                                            "5301010003 CONDOMINIO IMOVEIS EM ESTOQUE",
                                            "5301010004 IPTU DE IMOVEIS EM ESTOQUE",
                                            "Subtotal",
                                        ]
                                        )

            ws_desp.range("A1").value = df_final.columns.tolist()
            
            ws_desp.range("A1").expand("right").api.Interior.Color = 12611584.0
            ws_desp.range("A1").expand("right").api.Font.Color = 16777215.0
            ws_desp.range("A1").expand("right").api.Font.Bold = True
            
            ws_desp.range("A1").expand("table").api.Columns.AutoFit()
            
            ws_desp.range("A2").value = df_final.values.tolist() 
            
            cell_addres = ws_desp.range("A1").expand('down').address
            if (cell:=re.search(r'(?<=:[$]A[$])[0-9]+', cell_addres)):
                num_row = int(cell.group())
                
                row1 = ws_desp.range("A1").expand('right')
                for cell in row1:
                    coluna:str = cell.value
                    endere:str = cell.address
                    if (letra:=re.search(r'[$][A-z]+', endere)):
                        letra = letra.group()
                        if "Repres.%" == coluna:       
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "0%"
                        elif 'V.H %' in coluna:
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "0%"
                        elif 'Descrição' == coluna:
                            continue
                        else:
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "#.##0;(#.##0)"            
                
                cell = ws_desp.range("A1").expand('right').address
                if (address:=re.search(r'(?<=:)[$][A-z]+', cell)):
                    address = address.group()
                    ws_desp.range(f'$A:{address}').api.Columns.AutoFit()
                    ws_desp.range(f'$A$1:{address}${num_row}').api.Borders.LineStyle = -4118
               
            wb.save()
                
        app.kill()
        try:
            del app
        except:
            pass
        sleep(5)
        Functions.fechar_excel(self.file_path)
        return True
        #df_final.to_excel('teste.xlsx', index=False)

    def criar_outras_desp(self) -> bool:
        Functions.fechar_excel(self.file_path)
        app = xw.App(visible=False)
        with app.books.open(self.file_path) as wb:
            sheet:Sheet = wb.sheets[self.__sheet_name]
            sheet.api.AutoFilterMode = False
            
            dados = sheet.range("A1").expand('table').value
            
            df:pd.DataFrame = pd.DataFrame(dados[1:], columns=dados[0])
            
            #mask = ((df["Conta 1"] != "5304010022 OUTRAS RECEITAS OPERACIONAIS - CRÉDITO IMPOSTO") &
                #    (df["Conta 1"] != "5101010007 GANHO/PERDA DISTRATO") & 
                #    (df["Conta 1"] != "5304010015 OUTRAS RECEITAS OPERACIONAIS"))
            
            #df = df[mask]
            
            try:
                ws_desp = wb.sheets.add('Despesas')
            except ValueError:
                wb.sheets['Despesas'].delete()
                ws_desp = wb.sheets.add('Despesas')
            
            datas:list = df['Ano/Mês'].unique().tolist() #type:ignore
            if isinstance(datas, list):
                datas.reverse()
            
            #df_final = deepcopy(df['Conta 1']) #type:ignore
            #df_final.drop_duplicates(inplace=True)
            
            df.rename(columns={'Conta 1': 'Descrição'}, inplace=True)
            
            df_final = df[
                (df["Descrição"] != "5304010022 OUTRAS RECEITAS OPERACIONAIS - CRÉDITO IMPOSTO") &
                (df["Descrição"] != "5101010007 GANHO/PERDA DISTRATO") &
                (df["Descrição"] != "5304010015 OUTRAS RECEITAS OPERACIONAIS") 
            ]
            
            df_final = formulas.criar_colunas_por_data(df_final, datas)
            df_final = formulas.criar_colunas_calculadas(df_final, datas, sub_total=True)
            df_final = formulas.ordenar(df_final, coluna='Descrição',
                                        ordem= [
                                            "5304010014 OUTRAS DESPESAS OPERACIONAIS",
                                            "5304010020 DESPESAS JUDICIAIS",
                                            "5304010023 GANHO/PERDA CARTEIRA CLIENTES",
                                            "5304010018 GANHO COM IMOBILIZADO",
                                            "5304010019 IMPOSTOS E TAXAS",
                                            "5302010008 DESPESAS TRIBUTÁRIAS",
                                            "5304010017 PERDAS EVENTUAIS",
                                            "5304010021 GANHO E PERDA LIQUIDA OUTROS INVESTIMENTOS",
                                            "5304010016 DOACOES E MULTAS INDEDUTIVEIS",
                                            "5304010007 (-) DESPESA CREDITO IMOBILIARIO",
                                            "5304010011 COFINS S/OUTRAS REC. OPERACIONAIS",
                                            "5304010004 RECEITAS TAXA DE REGISTRO",
                                            "5304010012 PIS S/OUTRAS REC. OPERACIONAIS",
                                            "5304010000 DESPESAS COM CONTINGENCIAS",
                                            "Subtotal",
                                        ]
                                        )
            
            df_reem = df[
                (df["Descrição"] == "5304010022 OUTRAS RECEITAS OPERACIONAIS - CRÉDITO IMPOSTO") |
                (df["Descrição"] == "5101010007 GANHO/PERDA DISTRATO") |
                (df["Descrição"] == "5304010015 OUTRAS RECEITAS OPERACIONAIS") 
            ]
            #import pdb; pdb.set_trace()             
            df_reem = formulas.criar_colunas_por_data(df_reem, datas)
            df_reem = formulas.criar_colunas_calculadas(df_reem, datas, sub_total=True)
            
            df_subtotal = df_final[df_final['Descrição'] == 'Subtotal']
            df_subtotal = pd.concat([df_subtotal, df_reem])
            df_subtotal = formulas.criar_acumulado(df_subtotal, datas, coluna='Descrição')

            ws_desp.range("A1").value = df_final.columns.tolist()
            
            ws_desp.range("A1").expand("right").api.Interior.Color = 12611584.0
            ws_desp.range("A1").expand("right").api.Font.Color = 16777215.0
            ws_desp.range("A1").expand("right").api.Font.Bold = True
            
            ws_desp.range("A1").expand("table").api.Columns.AutoFit()
            
            ws_desp.range("A2").value = df_final.values.tolist() 
            
            cell_addres = ws_desp.range("A1").expand('down').address
            if (cell:=re.search(r'(?<=:[$]A[$])[0-9]+', cell_addres)):
                num_row = int(cell.group()) + 2
                ws_desp.range(f"A{num_row}").value = df_reem.values   
                ws_desp.range(f"A{num_row + len(df_reem) + 1}").value = df_subtotal.values  
                
                row1 = ws_desp.range("A1").expand('right')
                for cell in row1:
                    coluna:str = cell.value
                    endere:str = cell.address
                    if (letra:=re.search(r'[$][A-z]+', endere)):
                        letra = letra.group()
                        if "Repres.%" == coluna:       
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "0%"
                        elif 'V.H %' in coluna:
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "0%"
                        elif 'Descrição' == coluna:
                            continue
                        else:
                            ws_desp.range(f'{letra}:{letra}').api.NumberFormat = "#.##0;(#.##0)"            
                
                cell = ws_desp.range("A1").expand('right').address
                if (address:=re.search(r'(?<=:)[$][A-z]+', cell)):
                    address = address.group()
                    ws_desp.range(f'$A$1:{address}${num_row + len(df_reem) + 1}').api.Columns.AutoFit()
                    ws_desp.range(f'$A$1:{address}${num_row + len(df_reem) + 1}').api.Borders.LineStyle = -4118
               
            wb.save()
                
        app.kill()
        try:
            del app
        except:
            pass
        sleep(5)
        Functions.fechar_excel(self.file_path)
        return True
        #df_final.to_excel('teste.xlsx', index=False)
