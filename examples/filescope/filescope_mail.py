"""
Exemplo de envio de e-mail via Exchange utilizando
um report de controle de diretórios extraído da
biblioteca filescope. A partir desse script, será
consolidado um fluxo de extração de indicadores e 
envio de report via e-mail utilizando as funcionalidades
da biblitoeca xchange_mail

------------------------------------------------------
                        SUMÁRIO
------------------------------------------------------
1. Configuração inicial
    1.1 Importação de bibliotecas
    1.2 Definição das variáveis do projeto
"""

# Autor: Thiago Panini
# Data de Criação: 21/04/2021


"""
------------------------------------------------------
-------------- 1. CONFIGURAÇÃO INICIAL ---------------
            1.1 Importação de bibliotecas
------------------------------------------------------ 
"""

# Funções xchange_mail
from xchange_mail.handler import send_simple_mail

# Python libs
import os
from dotenv import load_dotenv, find_dotenv
from pandas import read_csv, read_excel
import codecs
from datetime import datetime
from filescope.manager import controle_de_diretorio, convert_kb_into_str


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
EXAMPLE_PATH = os.path.join(PROJECT_PATH, 'examples/filescope')
HTML_PATH = os.path.join(EXAMPLE_PATH, 'filescope.html')

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
SUBJECT = '[FILESCOPE xchange_mail] Report Analítico por E-mail'
MAIL_SIGNATURE = ''

# Variáveis para anexo de arquivos no e-mail ou no body
REPORT_DIR = os.getenv('FILESCOPE_PATH')
DF_ON_BODY = False
DF_ON_ATTACHMENT = False

# Função auxiliar para conversão de grandezas
def convert_scale(value):
    if value > 1000000:
        return str(round(value / 1000000, 1)) + ' M'
    elif value > 1000:
        return str(round(value / 1000, 1)) + ' K'
    else:
        return str(round(value, 1))


"""
------------------------------------------------------
------ 2. ENVIO DE REPORT CUSTOMIZADO POR EMAIL ------
        2.1 Preparando corpo via template HTML
------------------------------------------------------ 
"""

# Lendo base de dados
df = controle_de_diretorio(root=REPORT_DIR)

# Lendo arquivo html
f = codecs.open(HTML_PATH, 'r', 'utf-8')
HTML_BODY = f.read()

# Extraindo indicadores da base
total_space = convert_kb_into_str(df['tamanho_kb'].sum())
qtd_files = convert_scale(len(df))
avg_age = str(int(round(df['dias_desde_criacao'].mean(), 0)))
avg_score = str(round(df.iloc[:100, :]['filescope_score'].mean(), 1))

# Substituindo indicadores
real_inds = [total_space, qtd_files, avg_age, avg_score]
template_inds = ['IND01', 'IND02', 'IND03', 'IND04']
for real, template in zip(real_inds, template_inds):
    HTML_BODY = HTML_BODY.replace(template, real)

# Substituindo lista de top arquivos
top_files = df.head(5)['arquivo'].values
template_filelist = ['nome_arquivo_0' + str(i) for i in range(1, 6)]
for filename, template in zip(top_files, template_filelist):
    HTML_BODY = HTML_BODY.replace(template, filename)

# Enviando email com único anexo
send_simple_mail(username=USERNAME,
                 password=PWD,
                 server=SERVER,
                 mail_box=MAIL_BOX,
                 subject=SUBJECT,
                 mail_body=HTML_BODY,
                 mail_signature=MAIL_SIGNATURE,
                 mail_to=MAIL_TO,
                 df=df,
                 df_on_body=DF_ON_BODY,
                 df_on_attachment=DF_ON_ATTACHMENT,
                 attachment_filename='controle_diretorio.csv')