from Entities.dependencies.sap import SAPManipulation
from Entities.dependencies.config import Config
from Entities.dependencies.credenciais import Credential
from Entities.dependencies.functions import Functions, P
from time import sleep
import os
import shutil
from datetime import datetime
from typing import Dict, Literal
from Entities import utils


class ExtrairRelatorio(SAPManipulation):
    download_path = os.path.join(os.getcwd(), 'download')
    if not os.path.exists(download_path):
        os.makedirs(download_path)
        
    @property
    def dict_relatorios(self) -> Dict[str, dict]:
        return {
            'despAdm': {
                'variante_nome' : 'DESP ADM',
                'variante_usuario' : 'LTORRES',
                'relatorio_path' : 'relat_despAdm.xlsx'
            },
            'despCom': {
                'variante_nome' : 'DESP COM',
                'variante_usuario' : 'LTORRES',
                'relatorio_path' : 'relat_despCom.xlsx'
            },
            'outrasDesp': {
                'variante_nome' : 'OUTRAS DEPES',
                'variante_usuario' : 'LTORRES',
                'relatorio_path' : 'relat_outrasDesp.xlsx'
            }
        }
    
    def __init__(self) -> None:
        crd:dict = Credential(Config()['crd']['sap']).load()
        super().__init__(user=crd['user'], password=crd['password'], ambiente=crd['ambiente'], new_conection=True)
        
            
    @staticmethod
    def limpar_download_path():
        download_path = ExtrairRelatorio.download_path
        if os.path.exists(download_path):
            for file in os.listdir(download_path):
                file_path = os.path.join(download_path, file)
                if os.path.isfile(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
    

    @SAPManipulation.start_SAP  
    def fbl3n(self, *, relatorio:Literal['despAdm', 'despCom', 'outrasDesp'], date:datetime) -> str:
        try:
            print(P(f"Extraindo relatório {relatorio}...", color='cyan'))
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n fbl3n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").maximize()
            
            self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
            self.session.findById("wnd[1]/usr/txtV-LOW").text = self.dict_relatorios[relatorio]['variante_nome']
            self.session.findById("wnd[1]/usr/txtENAME-LOW").text = self.dict_relatorios[relatorio]['variante_usuario']
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            
            self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = utils.primeiro_dia_mes(date).strftime("%d.%m.%Y")
            self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = utils.ultimo_dia_mes(date).strftime("%d.%m.%Y")
            
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            if "Nenhuma partida selecionada" in (text:=self.session.findById("wnd[0]/sbar/pane[0]").text):
                self.fechar_sap()
                print(P(text, color='red'))
                raise Exception(text)
            
            file_path = os.path.join(ExtrairRelatorio.download_path, datetime.now().strftime(f"%Y%m%d_%H%M%S_{self.dict_relatorios[relatorio]['relatorio_path']}"))

            self.session.findById("wnd[0]").sendVKey(16)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(file_path)
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(file_path)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            print(P(f"Relatório {relatorio} extraído com sucesso!", color='green'))
            sleep(7)
            Functions.fechar_excel(file_path)
            
            self.fechar_sap()
            
            return file_path
        except Exception as err:
            self.fechar_sap()
            raise err



    @SAPManipulation.start_SAP      
    def test(self):
        import pdb;pdb.set_trace()