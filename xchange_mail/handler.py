# Importando bibliotecas
from dotenv import load_dotenv
import os
from exchangelib import Credentials, Account, Configuration, Message, DELEGATE, \
                        FileAttachment, HTMLBody
import pandas as pd
import io
from pretty_html_table import build_table

# Definindo função de conexão com o servidor
def connect_exchange(username, password, server, mail_box, auto_discover=False, access_type=DELEGATE):
    """
    Realiza a conexão com o servidor através da conta exchange
    
    Parâmetros
    ----------
    :param username: e-mail do usuário com acesso de envio de e-mail do endereço smtp [type: string]
    :param password: senha do usuário com acesso de envio de e-mail do endereço smtp [type: string]
    :param server: servidor responsável por gerenciar o transporte [type: string]
    :param mail_box: endereço primário associado a conta de envio [type: string]
    :param auto_discover: flag para apontar ao EWS utilizando protocolo específico [type: bool, default=False]
    :param access_type: tipo de acesso relacionado as credenciais [type: obj, default=DELEGATE]
    
    Retorno
    -------
    :return account: objeto contendo informações compiladas de uma conta de usuário
    """
    
    # Definindo e configurando credenciais
    creds = Credentials(username=username, password=password)
    config = Configuration(server=server, credentials=creds)
    account = Account(primary_smtp_address=mail_box, credentials=creds,
                      autodiscover=auto_discover, access_type=access_type, config=config)
    
    return account

# Criando função para envio de arquivos anexo
def attach_file(name, df):
    """
    Armazena DataFrames em buffers e transforma o conteúdo em bytes para envio em anexo
    
    Parâmetros
    ----------
    :param name: nome/referência do arquivo com a extensão [type: string]
    :param df: base de dados em formato DataFrame [type: pd.DataFrame]
    
    Retorno
    -------
    :return attachment_list: lista contendo o nome e o conteúdo em bytes do arquivo anexo [type: list]
    """
    
    # Criando buffer para armazenamento de bytes
    buffer = io.BytesIO()
    
    # Resgatando extensão do arquivo
    file_name, file_ext = os.path.splitext(name)
    
    # Salvando arquivo no buffer de acordo com extensão
    if file_ext in ['.csv', '.txt']:
        df.to_csv(buffer)
    elif file_ext == '.xlsx':
        df.to_excel(buffer)
    else:
        print('Extensão inválida. Opções: "csv", "txt" e "xlsx"')
        
    # Lendo buffer e retornando referências do anexo
    buffer_content = buffer.getvalue()
    
    return [name, buffer_content]

# Função para formatação de código html em corpo de e-mail
def format_mail_body(string_mail_body, mail_signature='', **kwargs):
    """
    Função desenvolvida para formatação do corpo do e-mail em formato HTML.
    Opcionalmente, é possível enviar DataFrames em formato de tabela pré formatada
    utilizando a biblioteca pretty_html_table.
    
    Parâmetros
    ----------
    :param string_mail_body: corpo de e-mail em formato de string [type: string]
        *pode conter código html para devida formatação e conversão
    :param **kwargs: argumentos adicionais da função para formatação do corpo
        :arg df: base de dados a ser enviada no corpo [type: pd.DataFrame]
        :arg color: configuração de cor disponibilizada pelo formatador [type: string, default='blue_light']
        :arg font_size: tamanho da fonte [type: string, default='medium']
        :arg font_family: tipo da fonte [type: string, default='Century Gothic']
        :arg text_align: alinhamento do texto [type: string, default='left']
        
    Retorno
    -------
    :return HTMLBody(string): corpo de e-mail devidamente formatado [type: string]
    """
    
    # Extraindo parâmetros
    df = kwargs['df'] if 'df' in kwargs else None
    color = kwargs['color'] if 'color' in kwargs else 'blue_light'
    font_size = kwargs['font_size'] if 'font_size' in kwargs else 'medium'
    font_family = kwargs['font_family'] if 'font_family' in kwargs else 'Century Gothic'
    text_align = kwargs['text_align'] if 'text_align' in kwargs else 'left'
    
    if df is not None:
        html_df = build_table(df, 
                              color=color, 
                              font_size=font_size, 
                              font_family=font_family, 
                              text_align=text_align)
        
        return HTMLBody(string_mail_body + html_df + mail_signature)
    else:
        return HTMLBody(string_mail_body + mail_signature)
    
# Definindo função capaz de encapsular a mensagem
def send_mail_mult_files(meta_df, username, password, server, mail_box, subject, mail_body, 
                         mail_to, mail_signature='', auto_discover=False, access_type=DELEGATE):
    """
    Função desenvolvida para o gerenciamento de múltiplos arquivos DataFrame a serem
    enviados no e-mail configurado, seja em anexo ou no corpo de e-mail. Como principal
    parâmetro, essa função utiliza um DataFrame informativo chamado de "meta_df" contendo
    as instruções necessárias relacionadas aos DataFrames de entrada do código.
    
    Parâmetros
    ----------
    :param meta_df: base contendo parâmetros informativos de ação [type: pd.DataFrame]
                    a base meta_df contém um DataFrame distinto por linha e deve ser formada por
        :col input: índice relacionado a base de entrada
        :col name: nome com extensão da base a ser enviada em anexo ou no corpo do e-mail
        :col df: base de entrada em formato DataFrame em cada linha
        :col flag_body: flag para envio da base no corpo do e-mail
        :col flag_attach: flag para envio da base em anexo
    :param username: e-mail do usuário com acesso de envio de e-mail do endereço smtp [type: string]
    :param password: senha do usuário com acesso de envio de e-mail do endereço smtp [type: string]
    :param server: servidor responsável por gerenciar o transporte [type: string]
    :param mail_box: endereço primário associado a conta de envio [type: string]
    :param subject: título do e-mail a ser enviado [type: string]
    :param mail_to: lista de recipientes do e-mail [type: list]
    :param mail_signature: assinatura a ser colocada no final do e-mail [type: string or HTMLBody]    
    :param auto_discover: flag para apontar ao EWS utilizando protocolo específico [type: bool, default=False]
    :param access_type: tipo de acesso relacionado as credenciais [type: obj, default=DELEGATE]
 
    Retorno
    -------
    Essa função não possui retorno, além do envio do e-mail com as especificações configuradas
    """
    
    # Configurando conta de envio
    account = connect_exchange(username=username, password=password, server=server, mail_box=mail_box,
                               auto_discover=auto_discover, access_type=access_type)

    # Verificando dados a serem enviados no corpo
    meta_df_body = meta_df.query('flag_body == 1')
    if len(meta_df_body) > 0:
        html_df = pd.DataFrame(meta_df_body['df']).iloc[0, :]['df']
        html_body = format_mail_body(mail_body, df=html_df, mail_signature=mail_signature)
    else:
        html_body = format_mail_body(mail_body, mail_signature=mail_signature)
    
    # Preparando mensagem
    m = Message(account=account,
                subject=subject,
                body=html_body,
                to_recipients=mail_to)
    
    # Verificando anexos e preparando estruturas
    meta_df_attach = meta_df.query('flag_attach == 1')
    file_names = list(meta_df_attach['name'])
    file_dfs = list(meta_df_attach['df'])

    attach_dict = {file_names.index(name) + 1: {'name': name, 'df': df} for name, df in zip(file_names, file_dfs)}
    attachments = [attach_file(inner_dict['name'], inner_dict['df']) for idx, inner_dict in attach_dict.items()]

    # Anexando arquivos
    for name, content in attachments or []:
        file = FileAttachment(name=name, content=content)
        m.attach(file)

    # Enviando mensagem
    m.send_and_save()

# Definindo função capaz de encapsular a mensagem
def send_mail_one_file(username, password, server, mail_box, subject, mail_body, mail_to, mail_signature='',
                       auto_discover=False, access_type=DELEGATE, df=None, df_on_body=False, 
                       df_on_attachment=False, attachment_filename='file.csv'):

    """
    Função desenvolvida para o gerenciamento de envio de e-mails independente da presença
    de uma base de dados em formato DataFrame a ser enviada em anexo ou no corpo.
    
    Parâmetros
    ----------
    :param username: e-mail do usuário com acesso de envio de e-mail do endereço smtp [type: string]
    :param password: senha do usuário com acesso de envio de e-mail do endereço smtp [type: string]
    :param server: servidor responsável por gerenciar o transporte [type: string]
    :param mail_box: endereço primário associado a conta de envio [type: string]
    :param subject: título do e-mail a ser enviado [type: string]
    :param mail_to: lista de recipientes do e-mail [type: list]
    :param mail_signature: assinatura a ser colocada no final do e-mail [type: string or HTMLBody]    
    :param auto_discover: flag para apontar ao EWS utilizando protocolo específico [type: bool, default=False]
    :param access_type: tipo de acesso relacionado as credenciais [type: obj, default=DELEGATE]
    :param df: base de dados opcionalmente enviada em anexo ou no corpo [type: pd.DataFrame]
    :param df_on_body: flag para envio da base de dados no corpo do e-mail [type: bool, default=False]
    :param df_on_attachment: flag para envio da base de dados em anexo [type: bool, default=False]
    :param attachment_filename: nome do arquivo com extensão a ser enviado em anexo [type: string, default='file.csv']
 
    Retorno
    -------
    Essa função não possui retorno, além do envio do e-mail com as especificações configuradas
    """
    
    # Configurando conta de envio
    account = connect_exchange(username=username, password=password, server=server, mail_box=mail_box,
                               auto_discover=auto_discover, access_type=access_type)

    # Verificando dados a serem enviados no corpo
    if df_on_body and df is not None:
        html_body = format_mail_body(mail_body, df=df, mail_signature=mail_signature)
    else:
        html_body = format_mail_body(mail_body, mail_signature=mail_signature)

    # Preparando mensagem
    m = Message(account=account,
                subject=subject,
                body=html_body,
                to_recipients=mail_to)
    
    # Verificando anexos e preparando estruturas
    if df_on_attachment and df is not None:
        attachments = [attach_file(name=attachment_filename, df=df)]

        # Anexando arquivos
        for name, content in attachments or []:
            file = FileAttachment(name=name, content=content)
            m.attach(file)

    # Enviando mensagem
    m.send_and_save()
