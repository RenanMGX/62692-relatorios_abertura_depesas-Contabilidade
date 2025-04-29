import os
import json
from datetime import datetime
from typing import Literal
from .logs import Logs, traceback

class Informativo:
    path = os.path.join(os.getcwd(), "json", "informativo.json")
    if not os.path.exists(os.path.dirname(path)):
        os.makedirs(os.path.dirname(path))
    
    @staticmethod
    def load() -> list:
        if not os.path.exists(Informativo.path):
            with open(Informativo.path, 'w') as f:
                json.dump([], f)
        with open(Informativo.path, 'r') as f:
            return json.load(f)
        
    @staticmethod
    def register(text:str, *, color:Literal["", "<django:red>", "<django:yellow>", "<django:green>", "<django:blue>"]= "", register_log:bool=False) -> None:
        data = Informativo.load()
        text = datetime.now().strftime(f"[%d/%m/%Y - %H:%M:%S] - {text} {color}")
        data.append(text)
        with open(Informativo.path, 'w') as f:
            json.dump(data, f)
        print(text)
        if register_log:
            Logs().register(status='Report', description=text, exception=traceback.format_exc())
        
    @staticmethod
    def limpar() -> None:
        with open(Informativo.path, 'w') as f:
            json.dump([], f)
        
if __name__ == "__main__":
    Informativo.limpar()
    Informativo.register("Hello", color='<django:green>')
    Informativo.register("World!", color='<django:green>')
    