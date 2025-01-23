from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import os

# Caminho do arquivo Excel local
excel_file = 'dados.xlsx'  # Substitua pelo nome do seu arquivo

# Verifica se o arquivo existe
if not os.path.exists(excel_file):
    print(
        f"Arquivo '{excel_file}' não encontrado. Certifique-se de que o arquivo está no diretório correto."
    )
    exit()

# Abrir o arquivo Excel
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active  # Usa a primeira aba da planilha

# Obtém os dados das colunas (nomes e links)
nomes = [cell.value for cell in sheet['A'][1:]]  # Ignora o cabeçalho
links = [cell.value for cell in sheet['B'][1:]]  # Ignora o cabeçalho

mensagem_template = """
Olá, aqui é a Contabilizei!\n
O processo de abertura da sua empresa está começando e queremos que você se sinta seguro para iniciar sua jornada com a gente.\n
Nossos especialistas estão trabalhando em todo o processo de cadastro, por isso é importante que você esteja atendo ao seu whatsapp e realize todo o seu cadastro com o apoio do time.\n
Além disso, estamos te convidando para uma Reunião de Boas Vindas, onde vamos te contar rotinas importantes  para sua empresa:\n
• Passos da abertura do seu CNPJ
• Termos contábeis importantes
• Organização do fluxo financeiro da sua empresa
• Calendário Contábil\n
Acesse o link para escolher o melhor horário:
https://calendly.com/seja-bem-vindo/rotinas2?utm_source=wpp_manual&utm_campaign=basico
"""
# Substituir as quebras de linha por '%0A' para serem interpretadas corretamente pelo WhatsApp Web
mensagem_url = mensagem_template.replace("\n", "%0A")
print(mensagem_url)
# Nome do arquivo de log
log_file = "mensagens_log.txt"


# Função para registrar no log
def registrar_log(status, nome, numero, mensagem):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, "a") as log:
        log.write(
            f"{now} - {nome} - {numero} - {mensagem} - {status}\n{'-'*75}\n")

    if status == "SUCESSO":
        print(
            f"[{now}] Mensagem enviada com sucesso para {nome} ({numero}): {mensagem}\n{'-'*75}"
        )
    else:
        print(
            f"[{now}] Falha ao enviar mensagem para {nome} ({numero}): {mensagem}\n{'-'*75}"
        )


# Função para extrair o número de um link do WhatsApp
def extrair_numero(link):
    try:
        numero = link.split('/')[-1].strip()  # Extrai o número
        return numero
    except Exception as e:
        print(f"Erro ao extrair o número do link {link}: {e}")
        return None


# Configuração das opções do Chrome
chrome_options = Options()
chrome_options.add_argument(
    "--start-maximized")  # Inicia o navegador maximizado
chrome_options.add_argument("--no-sandbox")  # Para rodar no ambiente local
chrome_options.add_argument(
    "--disable-dev-shm-usage"
)  # Para evitar problemas com a memória compartilhada

# Usando o webdriver-manager para instalar o ChromeDriver correto
driver_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=driver_service, options=chrome_options)

# Abre o WhatsApp Web
driver.get("https://web.whatsapp.com")

# Espera até que o WhatsApp Web carregue completamente
try:
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//div[@id="pane-side"]')))
    print("Login no WhatsApp Web bem-sucedido!")
except Exception as e:
    print(
        "Erro ao fazer login no WhatsApp Web. Certifique-se de escanear o código QR a tempo."
    )
    input("Por favor, escaneie o QR code e pressione Enter para continuar...")

# Iterar sobre os nomes e links e enviar a mensagem personalizada
for nome, link in zip(nomes, links):  # Ignora o cabeçalho
    numero = extrair_numero(link)

    if numero:
        numero_formatado = f"+55{numero}"  # Ajuste para o seu país
        try:
            # Personaliza a mensagem com o nome
            mensagem = mensagem_template.format(nome=nome)
            # Constrói a URL para o envio de mensagem via WhatsApp Web
            url = f"https://web.whatsapp.com/send?phone={numero_formatado}&text={mensagem_url}"

            # Navega para a URL
            driver.get(url)

            # Aguarda até que o campo de texto da mensagem esteja disponível
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     '//div[@contenteditable="true"][@data-tab="10"]')))

            # Localiza o campo de texto da mensagem
            input_box = driver.find_element(
                By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
            time.sleep(5)
            # Envia a mensagem pressionando Enter
            input_box.send_keys(Keys.ENTER)

            # Log de sucesso
            registrar_log("SUCESSO", nome, numero_formatado, mensagem)

            # Aguarde alguns segundos para evitar problemas com o envio rápido
            time.sleep(5)

        except Exception as e:
            # Log de falha
            registrar_log("FALHA", nome, numero_formatado, mensagem)
    else:
        print(f"Link inválido: {link}")

# Fecha o navegador após o envio de todas as mensagens
driver.quit()
