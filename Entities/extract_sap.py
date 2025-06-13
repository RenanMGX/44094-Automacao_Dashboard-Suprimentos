from patrimar_dependencies.sap import SAPManipulation
from patrimar_dependencies.functions import Functions, P
from utils import Utils, datetime, List, Dict
from getpass import getuser
from time import sleep
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
from botcity.maestro import * #type: ignore
import traceback

import os

class ExtractSAP(SAPManipulation):
    def __init__(self, *, maestro:BotMaestroSDK|None, user:str, password:str, ambiente:str):
        super().__init__(user=user, password=password, ambiente=ambiente)
        
        self.download_path = os.path.join(os.getcwd(), "Downloads")
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
    
    @SAPManipulation.start_SAP
    def base_dados_dire(self, date:datetime=datetime.now()) -> str:
        file = os.path.join(self.download_path, datetime.now().strftime("%Y%m%d%H%M%S_base_dados_dire.xlsx"))
        
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n zmm010"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtP_DTINI").text = Utils.first_day_of_last_year(date).strftime("%d.%m.%Y")
        self.session.findById("wnd[0]/usr/ctxtP_DTFIM").text = Utils.last_day_of_current_year(date).strftime("%d.%m.%Y")
        self.session.findById("wnd[0]/usr/ctxtP_VARI").text = "/WALLISON2"
        self.session.findById("wnd[0]").sendVKey(8)
        
        self.session.findById("wnd[0]/usr/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(file)
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(file)
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        sleep(5)
        Functions.fechar_excel(file)
        
        self.fechar_sap()
        
        return file
    
    @SAPManipulation.start_SAP
    def base_dados(self, date:datetime=datetime.now()) -> list:
        result:list = []
        
        list_of_dates:List[Dict[str, datetime]] = Utils.get_dates_per_moth(date)
        for dates in list_of_dates:
            month = dates['first_day'].strftime("%Y_%B")
            file = os.path.join(self.download_path, datetime.now().strftime(f"%Y%m%d%H%M%S_{month}_base_dados.xlsx"))
            
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n zmm029"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtS_BADAT-LOW").text = dates['first_day'].strftime("%d.%m.%Y")
            self.session.findById("wnd[0]/usr/ctxtS_BADAT-HIGH").text = dates['last_day'].strftime("%d.%m.%Y")
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            self.session.findById("wnd[0]/usr/cntlGC_CCALV/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
            self.session.findById("wnd[0]/usr/cntlGC_CCALV/shellcont/shell").selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(file)
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(file)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            sleep(5)
            
            Functions.fechar_excel(file)
            
            result.append(file)
        
        self.fechar_sap()
        
        return result


if __name__ == "__main__":
    pass
