import openpyxl
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import os
import time
import datetime

# Obtém a hora atual para a saudação dinâmica
hora_atual = datetime.datetime.now().hour

# Define a saudação com base na hora do dia
if 6 <= hora_atual < 12:
    saudacao = "bom dia"
elif 12 <= hora_atual < 18:
    saudacao = "boa tarde"
else:
    saudacao = "boa noite"

# --- Planilha ---
PLANILHA_NOME = "contatos.xlsx"

TEMPO_ENTRE_MENSAGENS = 5 # Tempo de espera entre o envio de cada contato, para evitar ser detectado como spam
TEMPO_ENTRE_PARTES_MENSAGEM = 1 # Tempo de espera entre o envio de cada parte da mensagem para o mesmo contato

# Diretório para armazenar o perfil do Chrome (para manter o login do WhatsApp Web)
USER_DATA_DIR = os.path.join(os.getcwd(), "chrome_profile")

# --- Inicialização do WebDriver (Navegador) ---
try:
    chrome_options = Options()
    # Adiciona o diretório de dados do usuário para persistir o login do WhatsApp Web
    chrome_options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    # Exclui mensagens de log desnecessárias do console do ChromeDriver
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # Inicializa o WebDriver do Chrome
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window() # Maximiza a janela do navegador
    print("Navegador Chrome inicializado (ChromeDriver gerenciado automaticamente).")

    # URL do WhatsApp Web
    driver.get("https://web.whatsapp.com/")

    print("Aguardando o carregamento do WhatsApp Web. Se necessário, escaneie o QR Code.")

    # Espera até que o elemento do painel lateral do WhatsApp Web esteja presente, indicando que a página carregou e o usuário está logado
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="pane-side"]'))
    )
    print("WhatsApp Web carregado com sucesso! Você deve estar logado.")

except Exception as e:
    print(f"Erro ao inicializar o navegador ou carregar o WhatsApp Web: {e}")
    print("Verifique se sua versão do Selenium é 4.6.0 ou superior (`pip install --upgrade selenium`).")
    print("Certifique-se de que o diretório de perfil de usuário está acessível e não corrompido.")
    exit() # Sai do script em caso de erro crítico

# --- Carregar Planilha de Dados ---
try:
    workbook = openpyxl.load_workbook(PLANILHA_NOME)
    contatos = workbook['lista01'] 
    nome_da_planilha_carregada = contatos.title

    print(f"Planilha '{PLANILHA_NOME}' carregada com sucesso.")
except FileNotFoundError:
    print(f"Erro: Planilha '{PLANILHA_NOME}' não encontrada. Verifique o nome e o caminho.")
    driver.quit() # Fecha o navegador antes de sair
    exit()
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    driver.quit() # Fecha o navegador antes de sair
    exit()

# Variáveis para rastrear o status do envio 
mensagens_enviadas = 0
mensagens_falhas = 0
erros_log = [] # Lista para armazenar detalhes de erros

# Itera sobre as linhas da planilha, começando da segunda linha (assumindo que a primeira é o cabeçalho)
for linha_idx, linha in enumerate(contatos.iter_rows(min_row=2), start=2):
    # Extrai os valores das colunas (ajuste os índices se suas colunas forem diferentes)
    nome = linha[0].value
    telefone = linha[1].value
    usuario = "Kevin Mendes"

    # Verifica se os dados essenciais estão presentes
    if not (nome and telefone):
        print(f"Linha {linha_idx}: Dados incompletos. Pulando para a próxima linha.")
        mensagens_falhas += 1
        erros_log.append(f"Linha {linha_idx} - Dados incompletos: nome='{nome}', Telefone='{telefone}'")
        continue # Pula para a próxima iteração do loop

    # Limpa o número de telefone, removendo caracteres não numéricos
    telefone_limpo = ''.join(filter(str.isdigit, str(telefone)))

    # --- Definição das Partes da Mensagem ---
    mensagem_parte1 = (
        f"Olá, {nome}, {saudacao}. Meu nome é {usuario} e sou um dos organizadores (Staffs) do programa Miniempresa da JA Piauí, "
        "do qual seu/sua filho(a) está participando na escola."
    )

    """
    mensagem_parte2 = (
        "A feira de Miniempresa 2025 acontecerá nos dias 27 e 28 deste mês no Teresina Shopping. "
        "Para tornar esse momento ainda mais especial para seu/sua filho(a), gostaríamos que "
        f"você escrevesse uma carta. Nela, você poderia expressar o que sente ao ver o(a) {nome} "
        "se engajando em um projeto tão importante para o futuro dele(a), ou palavras que o(a) "
        "encorajem a continuar e a se dedicar a este evento."
    )
    
    mensagem_parte3 = (
        "*Regras:* "
        "*1ª)* Escolha entre me enviar uma foto da carta escrita à mão, com letra legível, ou o envio da carta no chat do WhatsApp; "
        "*2ª)* Envie o quanto antes para podermos organizar melhor todas as cartas enviadas; "
        "*3ª) Lembre-se de manter essa conversa e o assunto sobre a carta em sigilo do seu filho ou de qualquer outro participante do evento.*"
    )

    mensagem_parte4 = "Podemos contar com você?"
    """

    # Lista contendo todas as partes da mensagem na ordem de envio
    todas_as_partes_mensagem = [
        mensagem_parte1,
        #mensagem_parte2,
        #mensagem_parte3,
        #mensagem_parte4
    ]

    # Constrói o link do WhatsApp Web com o número de telefone
    link_msg_whatsapp = f'https://web.whatsapp.com/send?phone={telefone_limpo}'

    print(f"\n--- Tentando enviar mensagem para: {nome} ({telefone_limpo}) ---")
    try:
        # PRIMEIRO PASSO: Navega o navegador para o link do WhatsApp Web
        driver.get(link_msg_whatsapp)

        # XPath da caixa de texto de mensagem no WhatsApp Web
        message_box_xpath = '//div[@contenteditable="true"][@data-tab="10"]'

        # SEGUNDO PASSO: Espera até 30 segundos para que a caixa de texto da mensagem esteja presente na página
        msg_box = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, message_box_xpath))
        )

        # TERCEIRO PASSO: Envia cada parte da mensagem sequencialmente
        for i, parte in enumerate(todas_as_partes_mensagem):
            msg_box.clear() # Limpa a caixa de texto antes de digitar
            msg_box.send_keys(parte) # Envia a parte da mensagem
            time.sleep(1) # Pequena pausa para processamento
            msg_box.send_keys(Keys.ENTER) # Envia a mensagem simulando o pressionar da tecla ENTER
            print(f"Parte {i+1} da mensagem enviada para {nome}.")
            if i < len(todas_as_partes_mensagem) - 1: # Se não for a última parte, espera um pouco mais
                time.sleep(TEMPO_ENTRE_PARTES_MENSAGEM)

        print(f"Todas as partes da mensagem enviadas com sucesso para {nome}.")
        mensagens_enviadas += 1

        # Pausa entre o envio de mensagens para evitar ser detectado como spam e sobrecarregar o WhatsApp Web
        time.sleep(TEMPO_ENTRE_MENSAGENS)

    except Exception as e:
        print(f"Não foi possível enviar mensagem para {nome} ({telefone_limpo}). Erro: {e}")
        mensagens_falhas += 1
        # Registra os detalhes do erro no log para análise posterior
        erros_log.append(f"Linha {linha_idx} - Falha ao enviar para {nome} ({telefone_limpo}): {e}")

# --- Conclusão ---
print("\n--- Relatório Final ---")
print(f"Total de mensagens processadas: {mensagens_enviadas + mensagens_falhas}")
print(f"Mensagens enviadas com sucesso: {mensagens_enviadas}")
print(f"Mensagens com falha: {mensagens_falhas}")

if erros_log:
    print("\n--- Detalhes das Falhas ---")
    # Salva os detalhes das falhas em um arquivo CSV para facilitar a depuração
    with open("erros_automacao_whatsapp.csv", "w", newline="", encoding="utf-8") as f:
        f.write("Detalhes do Erro\n")
        for erro in erros_log:
            f.write(f"{erro}\n")
            print(erro)
    print("\nOs detalhes das falhas foram salvos em 'erros_automacao_whatsapp.csv'.")
else:
    print("\nTodas as mensagens foram enviadas com sucesso (ou sem erros registrados).")

# Fecha o navegador no final da execução do script
driver.quit()
print("Automação concluída e navegador fechado.")
