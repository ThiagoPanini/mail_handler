"""
Script responsável por detalhar um exemplo de uso das
ferramentas do módulo xchange_mail construído para 
gerenciamento e envio de e-mails. O código simple.py
se utiliza de tais ferramentas para enviar um report
simples por e-mail com opções de envio de dados no
corpo ou em anexo.

------------------------------------------------------
                        SUMÁRIO
------------------------------------------------------
1. Configuração inicial
    1.1 Importação de bibliotecas
    1.2 Definição das variáveis do projeto
2. Envio de report simples
    2.1 Lendo base e enviando e-mail
"""

# Autor: Thiago Panini
# Data de Criação: 12/04/2021


"""
------------------------------------------------------
-------------- 1. CONFIGURAÇÃO INICIAL ---------------
            1.1 Importação de bibliotecas
------------------------------------------------------ 
"""

# Funções xchange_mail
from xchange_mail.handler import send_mail_mult_files

# Python libs
import os
from dotenv import load_dotenv, find_dotenv
from pandas import read_csv, read_excel, DataFrame
from datetime import datetime


"""
------------------------------------------------------
-------------- 1. CONFIGURAÇÃO INICIAL ---------------
        1.2 Definição de variáveis do projeto
------------------------------------------------------ 
"""

# Lendo variáveis de ambiente
load_dotenv(find_dotenv())

# Definindo variáveis de diretório
PROJECT_PATH = os.getenv('PROJECT_PATH')

# Definindo variáveis de configuração do e-mail
USERNAME = os.getenv('MAIL_FROM')
PWD = os.getenv('PASSWORD')
SERVER = 'outlook.office365.com'
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = [os.getenv('MAIL_TO')]

# Definindo variáveis de formatação do e-mail
SUBJECT = '[SIMPLE xchange_mail] Report HTML por E-mail'
NOW = datetime.now()
TODAY = NOW.strftime('%d/%m/%Y')
if NOW.hour >= 6 and NOW.hour < 12:
    GREETINGS = 'Bom dia'
elif NOW.hour >= 12 and NOW.hour < 18:
    GREETINGS = 'Boa tarde'
else:
    GREETINGS = 'Boa noite'
MAIL_BODY = f'{GREETINGS}, processo realizado com sucesso em {TODAY}! <br><br>'
MAIL_SIGNATURE = '<br>Att,<br>Desenvolvedores xchange_mail'

# Variáveis para anexo de arquivos no e-mail ou no body
CSV_FILEPATH = os.getenv('CSV_FILEPATH')
TXT_FILEPATH = os.getenv('TXT_FILEPATH')
DF_ON_BODY = True
DF_ON_ATTACHMENT = True
FILE_PATHS = [CSV_FILEPATH, TXT_FILEPATH]


"""
------------------------------------------------------
------------- 2. ENVIO DE REPORT SIMPLES -------------
           2.1 Lendo base e enviando e-mail
------------------------------------------------------ 
"""

# Lendo base de dados
dfs = []
for path in FILE_PATHS:
    try:
        dfs.append(read_csv(path))
    except Exception as e:
        dfs.append(read_excel(path))

# Montando arquivo de metadados
inputs = [i for i in range(1, len(FILE_PATHS) + 1)]
names = ['performances.csv', 'requirements.txt']
meta_df = DataFrame({})
meta_df['input'] = inputs
meta_df['name'] = names
meta_df['df'] = dfs
meta_df['flag_body'] = [1, 0]
meta_df['flag_attach'] = [1, 1]

# Enviando email com único anexo
send_mail_mult_files(meta_df=meta_df,
                     username=USERNAME,
                     password=PWD,
                     server=SERVER,
                     mail_box=MAIL_BOX,
                     subject=SUBJECT,
                     mail_body=MAIL_BODY,
                     mail_signature=MAIL_SIGNATURE,
                     mail_to=MAIL_TO)