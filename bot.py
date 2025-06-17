"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/custom-automations/python-custom/
"""

# Import for integration with BotCity Maestro SDK
from botcity.maestro import * #type: ignore
import traceback
from patrimar_dependencies.gemini_ia import ErrorIA
from patrimar_dependencies.screenshot import screenshot
from patrimar_dependencies.sharepointfolder import SharePointFolders
from Entities.extract_sap import ExtractSAP, Utils, datetime, os
from time import sleep


# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False #type: ignore

class Processados:
    def __init__(self):
        self.total_items = 1
        self.processados = 0
        
class Execute:
    @staticmethod
    def start():
        date_param = execution.parameters.get("date")
        if date_param:
            date = datetime.strptime(str(date_param), "%d/%m/%Y")
        else:
            date = datetime.now()  # Data de exemplo, pode ser alterada conforme necessário
        
        crd_param = execution.parameters.get("crd")
        if not isinstance(crd_param, str):
            raise ValueError("Parâmetro 'crd_param' deve ser uma string representando o label da credencial.")
        
        sharepoint_target = SharePointFolders(r'RPA - Dados\Relatorios\Relatorios Suprimentos').value
        if not os.path.exists(sharepoint_target):
            raise FileNotFoundError(f"O caminho {sharepoint_target} não existe. Verifique o caminho e tente novamente.")
    
        try:
            sap = ExtractSAP(
                    maestro=maestro,
                    user=maestro.get_credential(label=crd_param, key="user"),
                    password=maestro.get_credential(label=crd_param, key="password"),
                    ambiente=maestro.get_credential(label=crd_param, key="ambiente")
            )
            
            base_dados_dire_path = sap.base_dados_dire(date)
            Utils.save_to_folder(
                origin=base_dados_dire_path,
                destinity=sharepoint_target,
                name="base_dados_dire.json"
            )
            processados.processados += 1
            os.unlink(base_dados_dire_path)
            
            base_dados_path = sap.base_dados(date)
            base_dados_path = Utils.concatenar_planilhas(base_dados_path)
            Utils.save_to_folder(
                origin=base_dados_path,
                destinity=sharepoint_target,
                name="base_dados.json"
            )
            processados.processados += 1
            os.unlink(base_dados_path)
        finally:
            sap.fechar_sap() #type: ignore

if __name__ == '__main__':
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")
    
    processados = Processados()
    processados.total_items = 2
    
    try:
        Execute.start()
            
        maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.SUCCESS,
                    message="Tarefa Alimentar Dashboard Suprimentos finalizada com sucesso",
                    total_items=processados.total_items, # Número total de itens processados
                    processed_items=processados.processados, # Número de itens processados com sucesso
                    failed_items=(processados.total_items - processados.processados) # Número de itens processados com falha
        )
            
            
    except Exception as error:
        ia_response = "Sem Resposta da IA"
        try:
            token = maestro.get_credential(label="GeminiIA-Token-Default", key="token")
            if isinstance(token, str):
                ia_result = ErrorIA.error_message(
                    token=token,
                    message=traceback.format_exc()
                )
                ia_response = ia_result.replace("\n", " ")
        except Exception as e:
            maestro.error(task_id=int(execution.task_id), exception=e)

        maestro.error(task_id=int(execution.task_id), exception=error, screenshot=screenshot(), tags={"IA Response": ia_response})
        maestro.finish_task(
                    task_id=execution.task_id,
                    status=AutomationTaskFinishStatus.FAILED,
                    message="Tarefa Alimentar Dashboard Suprimentos finalizada com sucesso",
                    total_items=processados.total_items, # Número total de itens processados
                    processed_items=processados.processados, # Número de itens processados com sucesso
                    failed_items=(processados.total_items - processados.processados) # Número de itens processados com falha
        )
        