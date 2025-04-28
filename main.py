from Entities.tratar_dados import AlimentarBase, Tabela
from Entities.extrair_relatorio import ExtrairRelatorio, datetime
from Entities.dependencies.arguments import Arguments
from Entities.dependencies.functions import P
from Entities.dependencies.logs import Logs, traceback
import shutil
import os
import sys
import json

class JsonArgs:
    json_path = os.path.join(os.getcwd(), 'json', 'args.json')
    
    @staticmethod
    def get(delete_after:bool=True) -> dict:
        if not os.path.exists(JsonArgs.json_path):
            return {}
        else:
            with open(JsonArgs.json_path, 'r', encoding='utf-8') as _file:
                data = json.load(_file)
            if delete_after:
                os.remove(JsonArgs.json_path)
                
            data['date'] = datetime.fromisoformat(data['date'])
            return data

class Execute:
    @staticmethod
    def start():
        
        if not (args:=JsonArgs.get(delete_after=True)):
            print(P("Nenhum argumento encontrado."))
            raise Exception("Nenhum argumento encontrado.")
        else:
            files_path:dict = args['files_path']
            date:datetime = args['date']
            
            sap = ExtrairRelatorio()
            sap.limpar_download_path()
            
            # Relatório de Despesas Administrativas
            if files_path.get('desp_adm'):
                try:
                    path_adm_file = sap.fbl3n(relatorio='despAdm', date=date)
                    AlimentarBase(files_path['desp_adm']).add(path_adm_file)
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    print(P("Erro ao processar o relatório de Despesas Administrativas."))
                try:
                    Tabela(files_path['desp_adm']).criar_adm()
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    print(P("Erro ao criar tabela de Despesas Administrativas."))

            # Relatório de Despesas Comerciais
            if files_path.get('desp_comercial'):
                try:
                    path_comer_file = sap.fbl3n(relatorio='despCom', date=date)
                    AlimentarBase(files_path['desp_comercial']).add(path_comer_file)
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    print(P("Erro ao processar o relatório de Despesas Comerciais."))
                try:
                    Tabela(files_path['desp_comercial']).criar_comercial()
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    print(P("Erro ao criar tabela de Despesas Comerciais."))
            
            # Relatório de Outras Despesas
            if files_path.get('outras_despesas'):
                try:
                    path_outras_desp_file = sap.fbl3n(relatorio='outrasDesp', date=date)
                    AlimentarBase(files_path['outras_despesas']).add(path_outras_desp_file)
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    print(P("Erro ao processar o relatório de Outras Despesas."))
                try:
                    Tabela(files_path['outras_despesas']).criar_outras_desp()
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    print(P("Erro ao criar tabela de Outras Despesas."))
                
            print(P("Atualização concluída com sucesso!"))
            
            #import pdb; pdb.set_trace()
    
    @staticmethod
    def test():
        print(JsonArgs.get(delete_after=False))
    
if __name__ == "__main__":
    Arguments({
        'start': Execute.start,
        'test': Execute.test
    })