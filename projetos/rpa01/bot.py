# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Faz interações com o drop-down
from botcity.web.util import element_as_select

# Converte tabela da web para dicionário
from botcity.web.parsers import table_to_dict

# Biblioteca do excel
# pip install botcity-excel-plugin
from botcity.plugins.excel import BotExcelPlugin

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

# Criar um arquivo excel
excel = BotExcelPlugin()

# Adiciona o header
excel.add_row(["Id", "Cidade", "População"])

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    # O Maestro é responsável por fazer a governança dos robôs.
    # Conectar ao Maestro 
    # O login é feito automaticamente quando rodamos dentro do maestro
    # maestro.login(server="https://developers.botcity.dev", login="83cf25be-7317-4d28-8c8f-d2454f2e13e6", key="83C_UAQFWLTP3TPKAJW8QXHT")

    # Subir o robô de forma manual pelo Easy Deploys
    # Compactar o projeto para zip
    # Escolher a runner
    
    # Abrir o wizard novamente
    # Launch Botcity Runner e clicar em Start
    # Nesse momento nos conectamos ao Maestro e ficamos ouvindo por novas tarefas.
    # Fica perguntando para o Maestro se tem novas coisas para rodar.
    # Ir no Maestro e criar uma nova task.

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\Users\epereiradua2\Downloads\treinamento_t2c_botcity\chromedriver-win64\chromedriver-win64\chromedriver.exe"

    # modelo xpath: //tag[@atributo='valor']
    # modelo xpath: //tag[contains(tag, 'valor')]
    
    # Entrar no site do Busca CEP
    bot.browse("https://buscacepinter.correios.com.br/app/faixa_cep_uf_localidade/index.php")
    bot.wait(2000)

    # Pesquisar por cidades de um estado específico
    drop_uf = element_as_select(bot.find_element("//select[@id='uf']", By.XPATH)) 
    drop_uf.select_by_value("SP")
    bot.wait(2000)
    
    # Capturar tabela com nomes das cidades
    btn_pesquisar = bot.find_element("//button[@id='btn_pesquisar']", By.XPATH)
    btn_pesquisar.click()
    bot.wait(2000)  

    # Converte tabela em dicionário
    table_dados = bot.find_element("//table[@id='resultado-DNEC']", By.XPATH) 
    table_dados = table_to_dict(table=table_dados)

    # Entrar no site do IBGE
    bot.navigate_to("https://cidades.ibge.gov.br/brasil/sp/panorama")
    
    int_contador = 1
    str_cidade_anterior = ""

    # Para cada cidade capturada
    for cidade in table_dados:
        
        # cidade da vez
        str_cidade = cidade["localidade"]

        # Controle de cidade ja processada
        if str_cidade_anterior == str_cidade:
            continue

        # Controle de quantiade de ciclo    
        if int_contador <= 5:

            # Inserindo cidade no campo de pesquisa
            campo_pesquisa = bot.find_element("//input[@placeholder='O que você procura?']", By.XPATH)
            campo_pesquisa.send_keys(str_cidade)

            # opcao_cidade = bot.find_element(f"//a[contains(span, '{str_cidade}')]", By.XPATH)
            # Definindo uma âncora com mais tags
            opcao_cidade = bot.find_element(f"//a[span[contains(text(), '{str_cidade}')] and span[contains(text(), 'SP')]]", By.XPATH)
            bot.wait(2000)

            opcao_cidade.click()
            bot.wait(2000)

            # Capturar população
            populacao = bot.find_element("//div[@class='indicador__valor']", By.XPATH)
            str_populacao = populacao.text

            # Adiciona linha no excel
            print(int_contador, str_cidade, str_populacao) 
            excel.add_row([int_contador, str_cidade, str_populacao])
            
            # Criar nova entrada de log no maestro
            maestro.new_log_entry(activity_label="CIDADES", values={"CIDADE": f"{str_cidade}", 
                                                                    "POPULACAO": f"{str_populacao}",
                                                                    "ID": f"{int_contador}"
                                                                     })         

            int_contador = int_contador + 1
            str_cidade_anterior = str_cidade
        else:
            print("Limite de cidades atingidos.")
            break   
    
    # No final do processo, exportar um relatório em Exlcel com as colunas: Cidade, População
    excel.write(r"C:\Users\epereiradua2\Downloads\treinamento_t2c_botcity\projetos\rpa01\saida\Infos_Cidades.xlsx")

    # Wait 3 seconds before closing
    bot.wait(5000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
         task_id=execution.task_id,
         status=AutomationTaskFinishStatus.SUCCESS,
         message="Task Finished OK."
    )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
