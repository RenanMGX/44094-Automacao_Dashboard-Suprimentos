from Entities.extract_sap import ExtractSAP, Utils, datetime, os, Config
from Entities.dependencies.arguments import Arguments

class Execute:
    sap = ExtractSAP()
    date = datetime.now()
    
    
    @staticmethod
    def base_dados_dire():
        base_dados_dire_path = Execute.sap.base_dados_dire(Execute.date)
        Utils.save_to_folder(
            origin=base_dados_dire_path,
            destinity=Config()['path']['sharepoint_target'],
            name="base_dados_dire.json"
        )
        os.unlink(base_dados_dire_path)
    
    @staticmethod
    def base_dados():
        base_dados_path = Execute.sap.base_dados(Execute.date)
        base_dados_path = Utils.concatenar_planilhas(base_dados_path)
        Utils.save_to_folder(
            origin=base_dados_path,
            destinity=Config()['path']['sharepoint_target'],
            name="base_dados.json"
        )
        os.unlink(base_dados_path)
        

    
    
if __name__ == "__main__":
    Arguments({
        'base_dados_dire': Execute.base_dados_dire,
        'base_dados': Execute.base_dados
    })