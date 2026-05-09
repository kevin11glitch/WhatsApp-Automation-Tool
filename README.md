# 🤖 WhatsApp Automation Tool - Junior Achievement Staff Project

Este projeto nasceu de uma necessidade real identificada durante a minha participação como **Staff no programa Miniempresa da Junior Achievement (2025)**. 

### 💡 O Problema
No decorrer do programa, surgiu o desafio logístico de coletar as cartas de apoio e carinho dos pais dos alunos participantes. O processo manual de envio de lembretes e coleta de dados era extremamente moroso e propenso a falhas de comunicação. Diante disso, desenvolvi esta aplicação para **automatizar o contato direto com os responsáveis**, otimizando o tempo da equipe e garantindo a participação dos alunos.

---

## 🚀 Funcionalidades

-   **Integração com Excel**: Extração automática de nomes e números de telefone diretamente de arquivos `.xlsx`.
-   **Saudação Inteligente**: O script identifica o horário local e define automaticamente se a mensagem começa com "Bom dia", "Boa tarde" ou "Boa noite".
-   **Gestão de Sessão (Chrome Profile)**: Utiliza uma pasta de perfil dedicada para salvar o login do WhatsApp Web, evitando a necessidade de escanear o QR Code em cada execução.
-   **Tratamento de Erros e Logs**: Gera um relatório automático (`erros_automacao_whatsapp.csv`) caso um número seja inválido ou o chat não carregue, permitindo o acompanhamento manual posterior.
-   **Segurança e Antispam**: Implementação de *delays* estratégicos para simular o comportamento humano e evitar bloqueios na plataforma.

## 🛠️ Tecnologias Utilizadas

-   **Python 3**: Linguagem principal do projeto.
-   **Selenium**: Automação do navegador para interação com a interface web do WhatsApp.
-   **Openpyxl**: Biblioteca para leitura e manipulação de planilhas Excel.
-   **Urllib**: Formatação de caracteres especiais para mensagens personalizadas via URL.

## 📂 Estrutura do Projeto

-   `automacao_whatsapp.py`: Script principal com a lógica de automação.
-   `contatos.xlsx`: Planilha base (deve conter a aba `lista01` com Nome na coluna A e Telefone na coluna B).
-   `erros_automacao_whatsapp.csv`: Log de erros gerado automaticamente pelo sistema.
-   `chromedriver.exe`: Driver necessário para o controle do Google Chrome (deve ser compatível com a versão do seu navegador).

## 🛠️ Instalação das Bibliotecas
Para que o script funcione, você precisa instalar as bibliotecas que fazem a ponte entre o Python, o Excel e o Navegador. Siga os passos abaixo no seu terminal:
1. **Instalação Direta:** Execute esse comando para instalar as dependências oficiais: `pip install selenium openpyxl`
2. **O que cada biblioteca faz no projeto?**
     - **Selenium:** É o motor do projeto. Ele permite que o Python "assuma o controle" do Google Chrome, clique em botões e digite as mensagens por você.
     - **Openpyxl:** É a biblioteca que permite ao Python abrir o seu arquivo contatos.xlsx, ler as linhas e extrair os nomes e números sem que você precise abrir o Excel manualmente.
3. **Configuração do WebDriver (Essencial):**
     - Além das bibliotecas, o Selenium precisa do ChromeDriver:
     - Ele deve ser da mesma versão do seu Google Chrome.
     - Deve ser colocado na mesma pasta onde está o seu script automacao_whatsapp.py.
     - Você pode baixá-lo no site oficial: [Chromedriver Downloads](https://developer.chrome.com/docs/chromedriver/downloads?hl=pt-br).

## 📖Como Usar
-   Clone o repositório para sua máquina local com o link: ```https://github.com/kevin11glitch/WhatsApp-Automation-Tool.git```.
-   Preencha a planilha contatos.xlsx com os dados dos responsáveis.
-   Certifique-se de que o chromedriver.exe está na mesma pasta do script.
-   Execute o comando: `python automacao_whatsapp.py`
-   Na primeira vez, escaneie o QR Code. Nas próximas, o sistema entrará automaticamente.

## 🎓 Contexto Acadêmico
Desenvolvido por Kevin Iohan Mendes de Sousa, estudante de Engenharia de Software na Universidade Federal do Ceará (UFC), Campus Russas. Este projeto demonstra a aplicação prática de automação e manipulação de dados para resolver problemas logísticos em organizações do terceiro setor.
