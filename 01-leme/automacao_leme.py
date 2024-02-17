# %% [markdown]
# **O QUE ESSE CÓDIGO FAZ?**
# 
# Esse código faz o input dos dados de mensuração dos indicadores das iniciativas do Sebrae Minas no Leme.  
# 
# **METODOLOGIA**
# 
# A partir das bibliotecas selenium, webdriver e pandas, criamos um robô que executa as ações de clique e inputs no site do Leme com base em dados coletados pelas Unidades de Gestão Estratégica do Sebrae e também extraídos de dashboards disponíveis no Data Sebrae. Esses dados são alimentados em uma planilha modelo disponível em ".\01-dados\planilha-modelo.xlsx" e o robô utiliza essa planilha para fazer os inputs. 
# 
# Essa metodologia também é conhecida como Web Automation. O código está bem robusto com diversos recursos para agir diante de erros que podem acontecer durante a execução, sejam eles por conta de queda do servidor ou por questões relacionadas à má conexão de rede e ao próprio HTML do site. Porém, pode ser que novas alterações sejam necessárias com o tempo para a evolução e a robustez do robô. 
# 
# 
# ***Developers: Marcilio Duarte e Victor Hugo - Unidade de Inteligência Estratégica do Sebrae Minas.***
# 

# %% [markdown]
# # **1. IMPORTANDO LIBS**

# %%
# Importando bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException
import pandas as pd
import datetime
import time

# %% [markdown]
# # **2. CRIANDO FUNÇÕES**

# %%
# Função para acessar o leme
def leme_access(email, password):
    attempts = 0
    while attempts < MAX_RETRY_ATTEMPTS:
        try:
            # Criando o objeto Service do ChromeDriver
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service)

            # 1º Passo: Acessar o site leme.sebrae.com.br
            driver.get("https://leme.sebrae.com.br")

            # 2º Passo: Clicar no botão de login
            driver.find_element('xpath', '//*[@id="test-58"]/div/form/div/div[6]/a').click()
            time.sleep(2)

            # 3º Passo: Clicar em "Entrar com conta Microsoft"
            driver.find_element('xpath', '//*[@id="login-conta-microsoft"]/section/a').click()
            time.sleep(2)

            # 4º Passo: Inserir o email
            email_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]')))
            email_input.send_keys(email)
            time.sleep(2)

            # 5º Passo: Clicar em "Avançar"
            next_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="idSIButton9"]')))
            next_button.click()
            time.sleep(2)

            # 6º Passo: Inserir a senha
            password_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]')))
            password_button.send_keys(password)
            time.sleep(2)

            # 7º Passo: Clicar em "Sim" para continuar conectado
            try:
                sim_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="idSIButton9"]')))
                sim_button.click()
                time.sleep(2)
            except:
                pass
            # Verificar se o login foi bem-sucedido
            driver.get("https://leme.sebrae.com.br/web/sebrae/2024/home")
            time.sleep(2)

            entrar_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="test-58"]/div/form/div/div[6]/a')))
            entrar_button.click()
            time.sleep(2)

            driver.find_element('xpath', '//*[@id="login-conta-microsoft"]/section/a').click()
            time.sleep(10)

            current_url = driver.current_url
            if current_url == "https://leme.sebrae.com.br/web/sebrae/2024/home":
                print("O login foi realizado com sucesso.")
            return driver
        except WebDriverException as e:
            if 'Service Unavailable' in str(e) or 'maintenance downtime' in str(e) or '503' in str(e):
                attempts += 1
                print(f"Erro 503 (queda de servidor) detectado. Tentativa {attempts} de {MAX_RETRY_ATTEMPTS}. Aguardando {WAIT_TIME} segundos antes de tentar novamente.")
                time.sleep(WAIT_TIME)
            else:
                # Para outros erros, você pode querer relançar a exceção ou lidar de maneira diferente
                raise
    if attempts < MAX_RETRY_ATTEMPTS:
        return True
    else:
        print("Máximo de tentativas alcançadas. Falha ao acessar o site durante o login.")
        return None

# %%
def filtrar_unidade(driver, unidade):
    filter_input_selecionar_unidade = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Selecionar Unidade']")
    filter_input_selecionar_unidade.click()
    filter_input_selecionar_unidade.send_keys(unidade)
    search_response_button_selecionar_unidade = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, f"//h5[text()='{unidade}']")))
    search_response_button_selecionar_unidade.click()
    time.sleep(3)

def filtrar_componente_vinculado(driver, filter_key):
    filter_input_componente_vinculado = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Selecionar Componentes Vinculadas']")
    filter_input_componente_vinculado.click()
    filter_input_componente_vinculado.send_keys(filter_key)
    response_button_componente_vinculado = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, f"//h5[text() = '{filter_key}']")))
    response_button_componente_vinculado.click()
    time.sleep(3)

# %%
# Função para adicionar dados aos indicadores
def adicionando_dados_indicadores(driver, search_key, filter_key, value_to_add, unidade, clientes_prioritarios):
    attempts = 0
    while attempts<MAX_RETRY_ATTEMPTS:
        try:
            # 1º Passo: Acessar o site leme.sebrae.com.br
            driver.get("https://leme.sebrae.com.br/web/sebrae/2023/home")
            time.sleep(10)

            # 2º Passo: ir para a aba de monitoramento
            driver.get("https://leme.sebrae.com.br/web/sebrae/indicadoresmonitoramento")
            time.sleep(5)

            # 3º Passo: limpando filtros
            error_count_filter = 0
            while error_count_filter <2:
                try:
                   clear_filter = WebDriverWait(driver, 10).until(
                         EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Limpar filtros')]"))
                     )
                   clear_filter.click()
                   break
                except:
                    error_count_filter += 1
            print('Limpeza de filtro aplicada! Seguindo...')
            time.sleep(5)

            # 4º Passo: inserir o termo no buscador e buscar
            search_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="portlet-wrapper-indicator"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/input'))
            )
            search_input.send_keys(search_key)

            search_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="portlet-wrapper-indicator"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/span/button[1]/span[1]/span'))
            )
            search_button.click()

            # 5º Passo: inserir filtros
            filter_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="portlet-wrapper-indicator"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/span/button[2]/span[1]'))
            )
            filter_button.click()
            time.sleep(2)

            if search_key == 'Recomendação (NPS)':
                filtrar_unidade(driver, unidade)
                filtrar_componente_vinculado(driver, filter_key)

            else:
                filtrar_componente_vinculado(driver, filter_key)
            time.sleep(2)

            filter_apply_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="indicator-advanced-search"]/div[2]/button'))
            )
            filter_apply_button.click()
            print("Filtro aplicado com sucesso!")
            time.sleep(5)

            # 6º Passo: clicar no botão para acessar detalhes do indicador
            details_button = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, f"//a[text()='{search_key}']"))
            )
            details_button.click()
            time.sleep(10)

             # 7º Passo:add dados do indicador
            add_data_button = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "a.btn-tab span.icon-pencil"))
                    )
            add_data_button.click()
            time.sleep(2)

            if clientes_prioritarios != '-':
                # Inserir o valor específico no mês correspondente
                value_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="modal-tab4"]/div[2]/tg-indicator-table/div/div/table/tbody/tr[13]/td[8]/div[1]/input'))
                )
                value_input.clear()
                value_input.send_keys(value_to_add)

                # Inserindo qtd de clientes
                client_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="modal-tab4"]/div[2]/tg-indicator-table/div/div/table/tbody/tr[13]/td[9]/div[1]/input'))
                )
                client_input.clear()
                client_input.send_keys(clientes_prioritarios)

            else:
                # Inserir o valor específico no mês correspondente
                value_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="modal-tab4"]/div[2]/tg-indicator-table/div/div/table/tbody/tr[13]/td[8]/div[1]/input'))
                )
                value_input.clear()
                value_input.send_keys(value_to_add)
            time.sleep(5)

            # 8º Passo: clicar no botão de salvar

            save_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="modal-tab4"]/div[1]/button'))
            )
            save_button.click()
            time.sleep(5)

            # 9º Passo: clicar no botão de fechar
            try:
               close_button = WebDriverWait(driver, 10).until(
                       EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label='Close']"))
                   )
               close_button.click()

               time.sleep(5)
            except:
               close_button = driver.find_element(By.CSS_SELECTOR, "body > div.indicator-snippet.view-indicator-modal.modal.fade.in > div > div > div.modal-header > button")
               close_button.click()


            # 10º Passo: Removendo filtros aplicados
            filter_remove_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="portlet-wrapper-indicator"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/span/button[2]'))
            )
            filter_remove_button.click()

            time.sleep(5)

            try:
                if search_key == 'Recomendação (NPS)':
                   # Retirando filtros para não afetar próxima busca
                    filter_reset_unidade = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="indicator-advanced-search"]/div[1]/div[2]/tg-agency-advanced-search/div/div/div/button/div'))
                    )
                    filter_reset_unidade.click()

                    time.sleep(2)

                    filter_reset_componente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="indicator-advanced-search"]/div[1]/div[5]/tg-project-advanced-search/div/div/div/button/span'))
                        )
                    filter_reset_componente.click()

                else:
                     filter_reset_componente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="indicator-advanced-search"]/div[1]/div[5]/tg-project-advanced-search/div/div/div/button/span'))
                        )
                     filter_reset_componente.click()  
            except:
                pass
            time.sleep(5)

            # 11º Passo: clicar no botão de LIMPAR filtros
            try:
               clear_filter = WebDriverWait(driver, 10).until(
                     EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Limpar filtros')]"))
                 )
               clear_filter.click()

            except:
                pass
            
            # 12º Passo:aplicando filtro vazio
            apply_empty_filter_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="indicator-advanced-search"]/div[2]/button'))
            )
            apply_empty_filter_button.click()
            break
        except WebDriverException as e:
            if 'Service Unavailable' in str(e) or 'maintenance downtime' in str(e) or '503' in str(e):
                attempts +=1
                print(f"Erro 503 (queda de servidor) detectado na adição de dados ao indicador bucado com o termo: '{search_key}' e com o filtro '{filter_key}', referente ao valor: '{value_to_add}'.\nTentativa {attempts} de {MAX_RETRY_ATTEMPTS}. Vamos tentar novamente em {WAIT_TIME} segundos.")
                time.sleep(WAIT_TIME)
            else:
                raise
    if attempts < MAX_RETRY_ATTEMPTS:
        print(f"Sucesso ao adicionar dados para o indicador buscado com o termo: '{search_key}' e com o filtro: '{filter_key}', referente ao mês de 12/2023, o valor: {value_to_add} e quantidade de clientes {clientes_prioritarios}.")
        return True
    else:
        print("Máximo de tentativas alcançadas. Falha ao acessar o site.")
        return None

# %%
def acionar_web_scrapping(df, email, password, results_path):
    # Inicializa o driver do Selenium com o ChromeDriver
    start = datetime.datetime.now()
    print(f'Começando a rodar em: {start}')
    driver = leme_access(email=email, password=password)

    for index, row in df.iterrows():
        cod_ind = row['cod_ind']
        termo_buscador = row['termo_buscador']
        mes_de_ref = row['mes_de_ref']
        valor = row['valor']
        clientes_prioritarios = row['clientes_prioritarios']
        componente_vinculado = row['componente_vinculado']
        unidade = row['unidade']
        erro = row['erro']

        # Inicializando a contagem de erros para a linha atual    
        error_count = 0
        while error_count < 4:
            try:
                print('***************** Iniciando Processo *****************\n')     
                print(f'Iniciando processo para o indicador {termo_buscador} e componentes {componente_vinculado}...')

                if error_count > 1:  # Na terceira e quarta tentativas
                    if " -MR " in componente_vinculado:
                        if " - MR " in componente_vinculado:
                            # Remove o espaço após o traço
                            componente_vinculado = componente_vinculado.replace(" - MR ", " -MR ")
                        else:
                            # Adiciona um espaço após o traço
                            componente_vinculado = componente_vinculado.replace(" -MR ", " - MR ")

                adicionando_dados_indicadores(driver=driver, search_key=termo_buscador, filter_key=componente_vinculado, value_to_add=str(valor), unidade=unidade, clientes_prioritarios=str(clientes_prioritarios))
                print(f"Indicador {termo_buscador} e componente {componente_vinculado} atualizado com sucesso no Leme! Partindo para o próximo...")

                df.loc[index, 'erro'] = 0
                break
            except Exception as e:
                error_count += 1
                print(f"Erro ao atualizar o indicador {termo_buscador}. Tentativa {error_count} de 4. Erro: {str(e)}")

        if error_count == 4:
            df.loc[index, 'erro'] = 1

    df.to_excel(results_path+fr'\{termo_buscador}_planilha_automacao_leme_error.xlsx',index=False)

    # Fechando o navegador depois de concluir todas as iterações
    driver.quit()
    end = datetime.datetime.now()
    print(f'*****É TETRAAAAAA, ACABOUUUUUUUU!!! DEMOROU {end-start} DE TEMPO PRA RODAR, MAS FOI!\nÉ O FIM E ATÉ A PRÓXIMA MEUS AMIGOS.*****')

# %%
def ler_df_e_aba_certa(file_path, sheet):
    df = pd.read_excel(file_path, sheet_name=sheet)
    df['valor'] = df.apply(lambda row: f"{row['valor']:.2f}" if row['termo_buscador'] in ['Recomendação (NPS)', 'Faturamento'] else row['valor'], axis=1)
    return df

# %% [markdown]
# # **3. EXECUTANDO**

# %%
# Constantes para múltiplas tentativas em caso de QUEDA DO SERVIDOR
MAX_RETRY_ATTEMPTS = 3
WAIT_TIME = 1800

# Credenciais de acesso ao leme
email = str(input("Digite o email de rede para acessar o leme (EX: abc@sebraemg.com.br)."))
senha = str(input("Digite a senha de acesso ao leme"))
file = str(input("Insira o caminho da planilha para alimentar o leme."))
aba = str(input("Insira o nome EXATO da aba que contém os dados do indicador que você quer alimentar no LEME."))
results = str(input("Insira o caminho da armazenamento da planilha de erros."))


df = ler_df_e_aba_certa(file, aba)
acionar_web_scrapping(df, email, senha, results)


