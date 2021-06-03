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
from xchange_mail.mail import send_simple_mail

# Python libs
import os
from dotenv import load_dotenv, find_dotenv
from pandas import read_csv
from datetime import datetime
import ntpath


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
MAIL_TO = os.getenv('MAIL_TO')
if MAIL_TO.count('@') > 1:
    MAIL_TO = MAIL_TO.split(';')
else:
    MAIL_TO = [MAIL_TO]

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

# Embedding de imagem no corpo de e-mail
IMAGE_ON_BODY = True
IMAGE_LOCATION = os.getenv('IMG_FILEPATH')
IMAGE_FILENAME = ntpath.basename(IMAGE_LOCATION)
IMAGE_HYPERLINK = 'https://www.linkedin.com/in/thiago-panini/'

# Variáveis para anexo de arquivos no e-mail ou no body
CSV_FILEPATH = os.getenv('CSV_FILEPATH')
PDF_FILEPATH = os.getenv('PDF_FILEPATH')
DF_ON_BODY = True
DF_ON_ATTACHMENT = True
FILE_PATHS = [CSV_FILEPATH]


"""
------------------------------------------------------
------------- 2. ENVIO DE REPORT SIMPLES -------------
           2.1 Lendo base e enviando e-mail
------------------------------------------------------ 
"""

# Lendo base de dados
df = read_csv(CSV_FILEPATH)

# Enviando email com único anexo
send_simple_mail(username=USERNAME,
                 password=PWD,
                 server=SERVER,
                 mail_box=MAIL_BOX,
                 subject=SUBJECT,
                 mail_body=MAIL_BODY,
                 mail_signature=MAIL_SIGNATURE,
                 mail_to=MAIL_TO,
                 df=df.head(),
                 df_on_body=DF_ON_BODY,
                 df_on_attachment=DF_ON_ATTACHMENT,
                 attachment_filename='performances.csv',
                 image_on_body=IMAGE_ON_BODY,
                 image_location=IMAGE_LOCATION,
                 image_filename=IMAGE_FILENAME,
                 image_hyperlink=IMAGE_HYPERLINK,
                 local_attachment_path=PDF_FILEPATH)