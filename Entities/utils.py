from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import Dict, List
from Entities.dependencies.functions import Functions

import pandas as pd
import os
import shutil
from time import sleep

class Utils:
    @staticmethod
    def first_day_of_last_year(date:datetime=datetime.now()) -> datetime:
        date = date - relativedelta(years=1)
        date = date.replace(day=1, month=1, hour=0, minute=0, second=0, microsecond=0)
        return date

    @staticmethod
    def last_day_of_current_year(date: datetime=datetime.now()) -> datetime:
        date = date.replace(day=31, month=12, hour=0, minute=0, second=0, microsecond=0)
        return date
    
    @staticmethod
    def get_dates_per_moth(date:datetime=datetime.now()) -> List[Dict[str, datetime]]:
        dates:List[Dict[str, datetime]] = []
        for month in range(1, date.month+1):
            d = {
                "first_day": date.replace(day=1, month=month),
                "last_day": (date.replace(day=1, month=month) + relativedelta(months=1)) - relativedelta(days=1)
            }
            dates.append(d)
            
        return dates
    
    @staticmethod
    def concatenar_planilhas(lista_planilhas:List[str]) -> str:
        df = pd.DataFrame()
        for planilha in lista_planilhas:
            if planilha.endswith(('.xlsx', '.xls', 'xlsm')):
                planilha_path = planilha
                planilha = pd.read_excel(planilha_path)
                df = pd.concat([df, planilha], ignore_index=True)
                try:
                    os.remove(planilha_path)
                except PermissionError:
                    Functions.fechar_excel(planilha_path)
                    sleep(1)
                    os.remove(planilha_path)
                    
            else:
                raise Exception(f"{planilha=} não é um arquivo excel!")
        
        df = df.drop_duplicates()
        
        destiny=os.path.join(os.getcwd(), datetime.now().strftime("%Y%m%d%H%M%S_concatenado.xlsx"))
        df.to_excel(destiny, index=False)
        
        return destiny
    
    @staticmethod
    def save_to_folder(*, origin:str, destinity:str, name:str="") -> None:
        if os.path.exists(origin):
            if os.path.basename(origin).endswith(('.xlsx', '.xls', 'xlsm')):
                if os.path.exists(destinity):
                    if not name:
                        destinity = os.path.join(destinity, os.path.basename(origin)).replace('.xlsx', '.json').replace('.xls', '.json').replace('.xlsm', '.json')
                    else:
                        if name.endswith('.json'):
                            destinity = os.path.join(destinity, name)
                        else:
                            raise Exception("nome do arquivo deve terminar com '.json'")
                    pd.read_excel(origin).to_json(destinity, orient='records', date_format='iso')
                else:
                    raise Exception(f"{destinity=} não existe!")
            else:
                raise Exception(f"{origin=} não é um arquivo excel!")
        else:
            raise Exception(f"{origin=} não existe!")
        


if __name__ == "__main__":
    bot = Utils.get_dates_per_moth(datetime(2024,3,1))
    print("\n__main__\n")
    print(bot)