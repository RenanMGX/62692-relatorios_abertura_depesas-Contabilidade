from Entities.tratar_dados import AlimentarBase, Tabela
from Entities.extrair_relatorio import ExtrairRelatorio, datetime
from Entities.dependencies.arguments import Arguments
from Entities.dependencies.functions import P
from Entities.dependencies.logs import Logs, traceback
from Entities.dependencies.informativo import Informativo
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


def verificar_pastas():
    pastas = ['json', 'files']
    for pasta in pastas:    
        path = os.path.join(os.getcwd(), pasta)
        if not os.path.exists(path):
            os.makedirs(path)
    
    
class Execute:
    @staticmethod
    def start():
        verificar_pastas()
        Informativo.limpar()
        
        Informativo.register("Iniciando o processo de atualização.", color='<django:green>')
        if not (args:=JsonArgs.get(delete_after=True)):
            Informativo.register("Nenhum argumento encontrado.", color='<django:red>')
            raise Exception("Nenhum argumento encontrado.")
        else:
            files_path:dict = args['files_path']
            date:datetime = args['date']
            
            sap = ExtrairRelatorio()
            sap.limpar_download_path()
            
            # Relatório de Despesas Administrativas
            if files_path.get('desp_adm'):
                Informativo.register("Iniciando o processo de atualização de Despesas Administrativas.", color='<django:blue>')
                try:
                    path_adm_file = sap.fbl3n(relatorio='despAdm', date=date)
                    Informativo.register(f"Relatório de Despesas Administrativas baixado com sucesso: {path_adm_file}", color='<django:yellow>')
                    
                    AlimentarBase(files_path['desp_adm']).add(path_adm_file)
                    Informativo.register("Relatório de Despesas Administrativas processado com sucesso.", color='<django:yellow>')
                    
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    Informativo.register(f"Erro ao processar o relatório de Despesas Administrativas.<br>{str(err)}", color='<django:red>')
                    
                try:
                    Tabela(files_path['desp_adm']).criar_adm()
                    Informativo.register("Tabela de Despesas Administrativas criada com sucesso.", color='<django:yellow>')
                    
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    Informativo.register(f"Erro ao criar tabela de Despesas Administrativas..<br>{str(err)}", color='<django:red>')

            # Relatório de Despesas Comerciais
            if files_path.get('desp_comercial'):
                Informativo.register("Iniciando o processo de atualização de Despesas Comerciais.", color='<django:blue>')
                try:
                    path_comer_file = sap.fbl3n(relatorio='despCom', date=date)
                    Informativo.register(f"Relatório de Despesas Comerciais baixado com sucesso: {path_comer_file}", color='<django:yellow>')
                    
                    AlimentarBase(files_path['desp_comercial']).add(path_comer_file)
                    Informativo.register("Relatório de Despesas Comerciais processado com sucesso.", color='<django:yellow>')
                    
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    Informativo.register(f"Erro ao processar o relatório de Despesas Comerciais.<br>{str(err)}", color='<django:red>')
                    
                try:
                    Tabela(files_path['desp_comercial']).criar_comercial()
                    Informativo.register("Tabela de Despesas Comerciais criada com sucesso.", color='<django:yellow>')
                    
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    Informativo.register(f"Erro ao criar tabela de Despesas Comerciais.<br>{str(err)}", color='<django:red>')
            
            # Relatório de Outras Despesas
            if files_path.get('outras_despesas'):
                Informativo.register("Iniciando o processo de atualização de Outras Despesas.", color='<django:blue>')
                try:
                    path_outras_desp_file = sap.fbl3n(relatorio='outrasDesp', date=date)
                    Informativo.register(f"Relatório de Outras Despesas baixado com sucesso: {path_outras_desp_file}", color='<django:yellow>')
                    
                    AlimentarBase(files_path['outras_despesas']).add(path_outras_desp_file)
                    Informativo.register("Relatório de Outras Despesas processado com sucesso.", color='<django:yellow>')
                    
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    Informativo.register(f"Erro ao processar o relatório de Outras Despesas.<br>{str(err)}", color='<django:red>')
                    
                try:
                    Tabela(files_path['outras_despesas']).criar_outras_desp()
                    Informativo.register("Tabela de Outras Despesas criada com sucesso.", color='<django:yellow>')
                    
                except Exception as err:
                    Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    Informativo.register(f"Erro ao criar tabela de Outras Despesas.<br>{str(err)}", color='<django:red>')
                
            Informativo.register("Atualização concluída com sucesso!", color='<django:green>')
            
            #import pdb; pdb.set_trace()
    
    @staticmethod
    def test():
        print(JsonArgs.get(delete_after=False))
    
if __name__ == "__main__":
    Arguments({
        'start': Execute.start,
        'test': Execute.test
    })