from Entities.extract_sap import ExtractSAP, Utils, datetime, os
from patrimar_dependencies.arguments import Arguments
from patrimar_dependencies.config import Config
from patrimar_dependencies.credenciais import Credential

class Execute:
    date = datetime.now()
    
    crd = Credential(
        path_raiz=r'C:\Users\renan.oliveira\PATRIMAR ENGENHARIA S A\RPA - Documentos\RPA - Dados\CRD\.patrimar_rpa\credenciais',
        name_file="SAP_PRD"
    ).load()
    
    sap = ExtractSAP(
        maestro=None,
        user=crd['user'],
        password=crd['password'],
        ambiente=crd['ambiente']
    )
    
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
    