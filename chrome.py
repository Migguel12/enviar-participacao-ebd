from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Ler arquivo Excel e remover linhas com valores ausentes
nome_do_arquivo = 'membros.xlsx'
df = pd.read_excel(nome_do_arquivo)
df = df.dropna(subset=['NOME', 'CPF', 'Resposta'])

# Inicialize o Service fora do loop
service = Service('chromedriver.exe')

# Loop pelos registros do DataFrame
for index, row in df.iterrows():
    print(f"Index: {index} - Nome: {row['NOME']}")
    
    # Inicialize o navegador
    chrome = webdriver.Chrome(service=service)
    chrome.maximize_window()  # Abrir o navegador em tela cheia
    chrome.get("https://www.igrejacristamaranata.org.br/ebd/participacoes/")
    
    try:
        # Localizar e preencher o campo de CPF
        elemento_texto_cpf = WebDriverWait(chrome, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="icmEbdNacionalForm"]/div[2]/div/div[1]/input'))
        )
        print("Elemento CPF encontrado")
        elemento_texto_cpf.send_keys(row['CPF'])
        
        # Localizar e preencher o campo de Resposta
        elemento_texto_resposta = WebDriverWait(chrome, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="icmEbdNacionalForm"]/div[5]/div/div[2]/div/div[2]/div[1]/p'))
        )
        elemento_texto_resposta.send_keys(row['Resposta'])
        
        # Localizar e clicar na caixa de confirmação, garantindo que ela esteja visível
        elemento_caixa_de_confirmacao = WebDriverWait(chrome, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="aceitoOsTermos"]'))
        )
        chrome.execute_script("arguments[0].scrollIntoView(true);", elemento_caixa_de_confirmacao)
        time.sleep(1)  # Pausa para garantir visibilidade
        elemento_caixa_de_confirmacao.click()
        print("Caixa de confirmação marcada com sucesso")
        
        # Localizar e clicar no botão de participação
        elemento_botao_participacao = WebDriverWait(chrome, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="icmEbdNacionalForm"]/div[5]/div/div[1]/input[2]'))
        )
        chrome.execute_script("arguments[0].scrollIntoView(true);", elemento_botao_participacao)
        time.sleep(1)
        elemento_botao_participacao.click()
        print("Botão de participação clicado com sucesso")
        
        # Localizar e clicar no botão de enviar
        elemento_botao_enviar = WebDriverWait(chrome, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="icmEbdNacionalForm"]/button'))
        )
        chrome.execute_script("arguments[0].scrollIntoView(true);", elemento_botao_enviar)
        time.sleep(1)
        elemento_botao_enviar.click()
        print("Botão de enviar clicado com sucesso")
        
        # Esperar 10 segundos após enviar o formulário
        time.sleep(10)
    
    except Exception as e:
        print(f"Erro no índice {index}: {e}")
    
    finally:
        # Fechar o navegador
        time.sleep(2)
        chrome.quit()
