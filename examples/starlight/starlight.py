"""
Script responsável por detalhar um exemplo de uso das
ferramentas do módulo xchange_mail construído para 
gerenciamento e envio de e-mails. O código starlight.py
se utiliza de tais ferramentas para enviar um report
por e-mail utilizando um template criado a partir
da plataforma online BeeFree

------------------------------------------------------
                        SUMÁRIO
------------------------------------------------------
1. Configuração inicial
    1.1 Importação de bibliotecas
    1.2 Definição das variáveis do projeto
2. Envio de report customizado por e-mail
    2.1 Preparando corpo via template HTML
    2.2 2.2 Realizando envio de e-mail formatado
"""

# Autor: Thiago Panini
# Data de Criação: 11/04/2021


"""
------------------------------------------------------
-------------- 1. CONFIGURAÇÃO INICIAL ---------------
            1.1 Importação de bibliotecas
------------------------------------------------------ 
"""

# Funções xchange_mail
from xchange_mail.handler import connect_exchange, attach_file, \
    format_mail_body, send_simple_mail

# Third-part
import os
from dotenv import load_dotenv, find_dotenv
from pandas import read_csv, read_excel
import codecs
from datetime import datetime


"""
------------------------------------------------------
-------------- 1. CONFIGURAÇÃO INICIAL ---------------
        1.2 Definição de variáveis do projeto
------------------------------------------------------ 
"""

# Definindo variáveis de diretório
PROJECT_PATH = '/home/paninit/workspaces/xchange_mail'
EXAMPLE_PATH = os.path.join(PROJECT_PATH, 'examples/starlight')
HTML_PATH = os.path.join(EXAMPLE_PATH, 'starlight.html')
DEPARA_IMGS = os.path.join(EXAMPLE_PATH, 'depara_imgs.txt')
DEPARA_TAGS = os.path.join(EXAMPLE_PATH, 'depara_tags.txt')

# Lendo variáveis de ambiente
load_dotenv(find_dotenv())

# Definindo variáveis de configuração do e-mail
USERNAME = os.getenv('MAIL_BOX')
PWD = os.getenv('PASSWORD')
SERVER = 'outlook.office365.com'
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = [os.getenv('MAIL_TO')]

# Definindo variáveis de formatação do e-mail
SUBJECT = 'xchange_mail - Starlight - Report HTML por E-mail'
MAIL_SIGNATURE = ''

# Variáveis para anexo de arquivos no e-mail ou no body
CSV_FILEPATH = '/home/paninit/workspaces/sentimentor/ml/performance.csv'
DF_ON_BODY = False
DF_ON_ATTACHMENT = True
FILE_PATHS = [CSV_FILEPATH]


"""
------------------------------------------------------
------ 2. ENVIO DE REPORT CUSTOMIZADO POR EMAIL ------
        2.1 Preparando corpo via template HTML
------------------------------------------------------ 
"""

# Lendo arquivo html
f = codecs.open(HTML_PATH, 'r', 'utf-8')
HTML_BODY = f.read()

# Lendo depara de referências (imagens e tags)
depara_imgs = read_csv(DEPARA_IMGS, sep=';')
depara_tags = read_csv(DEPARA_TAGS, sep=';')

# Iterando sobre cada elemento do depara de imgs
local_img_list = list(depara_imgs['local_img'].values)
hosted_img_list = list(depara_imgs['hosted_img'].values)
for local_img, hosted_img in zip(local_img_list, hosted_img_list):
    HTML_BODY = HTML_BODY.replace(local_img, hosted_img)

# Iterando sobre cada elemento do depara de tags
tag_template_list = list(depara_tags['tag_template'].values)
tag_projeto_list = list(depara_tags['tag_projeto'].values)
for tag_template, tag_projeto in zip(tag_template_list, tag_projeto_list):
    HTML_BODY = HTML_BODY.replace(tag_template, tag_projeto)

# Substituição de indicadores vivos
date_report = datetime.now().strftime('%d/%m/%Y')
HTML_BODY = HTML_BODY.replace('__date_report__', date_report)


"""
------------------------------------------------------
------ 2. ENVIO DE REPORT CUSTOMIZADO POR EMAIL ------
       2.2 Realizando envio de e-mail formatado
------------------------------------------------------ 
"""

# Lendo base de dados a serem enviadas no corpo ou em anexo
dfs = []
for path in FILE_PATHS:
    try:
        dfs.append(read_csv(path))
    except:
        dfs.append(read_excel(path))

# Enviando email com único anexo
send_simple_mail(username=USERNAME,
                 password=PWD,
                 server=SERVER,
                 mail_box=MAIL_BOX,
                 subject=SUBJECT,
                 mail_body=HTML_BODY,
                 mail_signature=MAIL_SIGNATURE,
                 mail_to=MAIL_TO,
                 df=dfs[0],
                 df_on_body=DF_ON_BODY,
                 df_on_attachment=DF_ON_ATTACHMENT,
                 attachment_filename='performances.csv')