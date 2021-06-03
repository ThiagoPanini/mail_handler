"""
---------------------------------------------------
------------------- MODULE: Mail -------------------
---------------------------------------------------
This module allocates useful functions for sending
mails using exchangelib with simple configuration
steps

Table of Contents
---------------------------------------------------
1. Initial setup
    1.1 Importing libraries
2. Sending mails through exchange
    2.1 Auxiliar functions
    2.2 Mail sending functions
---------------------------------------------------
"""

# Author: Thiago Panini
# Date: 26/05/2021


"""
---------------------------------------------------
---------------- 1. INITIAL SETUP -----------------
             1.1 Importing libraries
---------------------------------------------------
"""

# Exchangelib classes
from exchangelib import Credentials, Account, Configuration, Message, DELEGATE, \
                        FileAttachment, HTMLBody

# Standard python libraries
import os
import ntpath
from dotenv import load_dotenv
import pandas as pd
import io
from pretty_html_table import build_table


"""
---------------------------------------------------
-------- 2. SENDING MAILS THROUGH EXCHANGE --------
              2.1 Auxiliar functions
---------------------------------------------------
"""

# Connecting to the server
def connect_exchange(username, password, server, mail_box, auto_discover=False, access_type=DELEGATE):
    """
    Connects to Exchange server and generates an Account object
    
    Parameters
    ----------
    :param username: user mail with rights for sending mails through the mail box provided [type: string]
    :param password: user passwords smtp [type: string]
    :param server: server for managing the mail sending [type: string]
    :param mail_box: primary address associated to the user account [type: string]
    :param auto_discover: flag for pointing to EWS using a specific protocol [type: bool, default=False]
    :param access_type: access type associated to the credentials provided [type: obj, default=DELEGATE]
    
    Return
    ------
    :return account: exchange object with user account information [type: Account]
    """
    
    # Setting up credentials, configuration and returning account
    creds = Credentials(username=username, password=password)
    config = Configuration(server=server, credentials=creds)
    account = Account(primary_smtp_address=mail_box, credentials=creds,
                      autodiscover=auto_discover, access_type=access_type, config=config)
    
    return account

# Function for streaming DataFrame objects and attaching it to the mail
def buffer_dataframe(name, df):
    """
    Stores DataFrames object on buffers and transform the content on bytes for sending attached
    
    Parameters
    ----------
    :param name: filename with extension (csv or xlsx) [type: string]
    :param df: DataFrame object to be attached [type: pd.DataFrame]
    
    Return
    ------
    :return attachment_list: list with name [0] and DataFrame content on bytes [1] of the DataFrame provided [type: list]
    """
    
    # Creating a buffer for storing bytes
    buffer = io.BytesIO()
    
    # Returning file extension
    file_name, file_ext = os.path.splitext(name)
    
    # Saving file on buffer according to its extension
    try:
        if file_ext in ['.csv', '.txt']:
            df.to_csv(buffer)
        elif file_ext == '.xlsx':
            df.to_excel(buffer)
        else:
            print('Invalid extension. Options: "csv", "txt" e "xlsx"')

        # Reading buffer content
        buffer_content = buffer.getvalue()
        
    except TypeError as te:
        # Bytes buffer was not supported. Trying string bytes
        buffer = io.StringIO()

        if file_ext in ['.csv', '.txt']:
            df.to_csv(buffer)
        elif file_ext == '.xlsx':
            df.to_excel(buffer)
        else:
            print('Invalid extension. Options: "csv", "txt" e "xlsx"')
        
        #  Reading buffer content
        buffer_content = buffer.getvalue().encode()
    
    return [name, buffer_content]

# Formatting html mail body and customizing DataFrames if applicable
def format_html_body(string_mail_body, mail_signature='', **kwargs):
    """
    Formats a mail string body using HTMLBody class. In addition, the function
    can receive a DataFrame object and transform it using pretty_html_table package
    for customizing the source object in a custom table to be sent on mail body.
    
    Parameters
    ----------
    :param string_mail_body: raw string mail body [type: string]
        *can have html code for be transformed on HTMLBody class
    :param **kwargs: additional parameters
        :arg df: DataFrame object to be sent on mail body as a custom table [type: pd.DataFrame]
        :arg color: color configuration from pretty_html_table [type: string, default='blue_light']
        :arg font_size: font size for html table built from DataFrame [type: string, default='medium']
        :arg font_family: font family for html table built from DataFrame [type: string, default='Century Gothic']
        :arg text_align: text allign for html table built from DataFrame [type: string, default='left']
        
    Return
    ------
    :return HTMLBody(string): mail body in a html format [type: HTMLBody]
    """
    
    # Extracting parameters from kwargs
    df = kwargs['df'] if 'df' in kwargs else None
    color = kwargs['color'] if 'color' in kwargs else 'blue_light'
    font_size = kwargs['font_size'] if 'font_size' in kwargs else 'medium'
    font_family = kwargs['font_family'] if 'font_family' in kwargs else 'Century Gothic'
    text_align = kwargs['text_align'] if 'text_align' in kwargs else 'left'
    
    # Building a html table from DataFrame if applicable
    if df is not None:
        html_df = build_table(df, 
                              color=color, 
                              font_size=font_size, 
                              font_family=font_family, 
                              text_align=text_align)
        
        return HTMLBody(string_mail_body + html_df + mail_signature)
    else:
        # There is no DataFrame on argument. Transforming just body and signature html strings
        return HTMLBody(string_mail_body + mail_signature)

# Sending a simple mail with useful customization
def send_simple_mail(username, password, server, mail_box, subject, mail_to, mail_body='', mail_signature='',
                     auto_discover=False, access_type=DELEGATE, df=None, df_on_body=False, 
                     df_on_attachment=False, attachment_filename='file.csv', image_on_body=False, 
                     image_location=None, image_filename='image.png', image_hyperlink=None, 
                     local_attachment_path=None, **kwargs):
    """
    Handles the mail sending of a simple mail. Things that this function can do:
        * Send a mail with simple mail subject, body and signature for one or more recipients
        * Send a mail with a custom HTML body or template for one or more recipients
        * Send a mail with a DataFrame attached or even on body using pretty_html build_table function
        * Send a mail with an image attached or even on body using cid
    
    Parameters
    ----------
    :param username: user mail with rights for sending mails through the mail box provided [type: string]
    :param password: user passwords smtp [type: string]
    :param server: server for managing the mail sending [type: string]
    :param mail_box: primary address associated to the user account [type: string]
    :param subject: mail subject [type: string]
    :param mail_body: body raw string or html code [type: string]
    :param mail_to: recipients list [type: list]
    :param mail_signature: raw string or html code to be put at the end of body [type: string, default='']
    :param auto_discover: flag for pointing to EWS using a specific protocol [type: bool, default=False]
    :param access_type: access type associated to the credentials provided [type: obj, default=DELEGATE]
    :param df: DataFrame object that can be sent attached or on mail body [type: pd.DataFrame, default=None]
    :param df_on_body: flag for sending DataFrame on mail body as a custom table [type: bool, default=False]
    :param df_on_attachment: flag for sending DataFrame file attached [type: bool, default=False]
    :param attachment_filename: filename for attached DataFrame [type: string, default='file.csv']
    :param image_on_body: flag for sending an image on mail body [type: bool, default=False]
    :param image_location: location of image stored on disk [type: string, default=None]
    :param image_filename: filename for attached image [type: string, default='image.png']
    :param image_hyperlink: hyperlink to be put on image body [type: string, default=None]
    :param local_attachment_path: path to file to be attached [type: string, default=None]
    :param **kwargs: additional parameters
        :arg df: DataFrame object to be sent on mail body as a custom table [type: pd.DataFrame]
        :arg color: color configuration from pretty_html_table [type: string, default='blue_light']
        :arg font_size: font size for html table built from DataFrame [type: string, default='medium']
        :arg font_family: font family for html table built from DataFrame [type: string, default='Century Gothic']
        :arg text_align: text allign for html table built from DataFrame [type: string, default='left']
 
    Return
    ------
    This function returns anything besides the mail sending
    """
    
    # Creating and configuring account using function parameters
    account = connect_exchange(username=username, password=password, server=server, mail_box=mail_box,
                               auto_discover=auto_discover, access_type=access_type)

    # Extracting kwargs
    color = kwargs['color'] if 'color' in kwargs else 'blue_light'
    font_size = kwargs['font_size'] if 'font_size' in kwargs else 'medium'
    font_family = kwargs['font_family'] if 'font_family' in kwargs else 'Century Gothic'
    text_align = kwargs['text_align'] if 'text_align' in kwargs else 'left'

    # Formatting html to be sent on body. If df is passed, it builds a custom html table
    if df_on_body and df is not None:
        html_body = format_html_body(mail_body, df=df, mail_signature=mail_signature, color=color,
                                     font_size=font_size, font_family=font_family, text_align=text_align)
    else:
        html_body = format_html_body(mail_body, mail_signature=mail_signature, color=color,
                                     font_size=font_size, font_family=font_family, text_align=text_align)

    # Creating a message object
    m = Message(account=account,
                subject=subject,
                body=html_body,
                to_recipients=mail_to)
    
    # Validating attachments
    if df_on_attachment and df is not None:
        attachments = [buffer_dataframe(name=attachment_filename, df=df)]

        # Attaching a DataFrame
        for name, content in attachments or []:
            file = FileAttachment(name=name, content=content)
            m.attach(file)

    # Putting image on body if applicable
    if image_on_body and image_location is not None:
        
        # Opening local image and creating the attachment content
        with open(image_location, 'rb') as f:
            img = FileAttachment(
                name=image_filename, content=f.read(),
                is_inline=True, content_id=image_location
            )

        # Attaching content and building a new HTMLBody with image
        m.attach(img)
        html_image_body = f'<img src="cid:{image_location}">'
        if image_hyperlink is not None:
            html_image_body = f'<a href={image_hyperlink}>' + html_image_body + '</a>'
        
        # Adding initial body and signature
        html_image_body = mail_body + html_image_body + mail_signature

        m.body = HTMLBody(html_image_body)

    # Verifying the need to attach a local file
    local_attachments = []
    if local_attachment_path is not None:
        try:
            with open(local_attachment_path, 'rb') as f:
                content = f.read()
            local_attachments.append((ntpath.basename(local_attachment_path), content))

            # Attaching to email
            for name, content in local_attachments or []:
                file = FileAttachment(name=name, content=content)
                m.attach(file)
        except Exception as e:
            print(f'Error on reading file {local_attachment_path}. Exception: {e}')

    # Sending message
    m.send_and_save()

# Sending a mail using a meta_df data for handling multiple DataFrames and actions
def send_mail_mult_files(meta_df, username, password, server, mail_box, subject, mail_body, 
                         mail_to, mail_signature='', auto_discover=False, access_type=DELEGATE):
    """
    Handles multiple DataFrames object using a meta_df DataFrame that guides actions for each object.
    The mailing proccess uses this meta_df for attaching, sending DataFrames on body and more.
    
    Parâmetros
    ----------
    :param meta_df: DataFrame object with informative paramters for guiding actions [type: pd.DataFrame]
        *the meta_df object must a have one DataFrame per line. The columns of meta_df are:
        :col input: numerical index for each DataFrame
        :col name: name with extension of each DataFrame
        :col df: DataFrame object
        :col flag_body: flag for sending the DataFrame on mail body
        :col flag_attach: flag for sending the DataFrame attached
    :param username: user mail with rights for sending mails through the mail box provided [type: string]
    :param password: user passwords smtp [type: string]
    :param server: server for managing the mail sending [type: string]
    :param mail_box: primary address associated to the user account [type: string]
    :param subject: mail subject [type: string]
    :param mail_body: body raw string or html code [type: string]
    :param mail_to: recipients list [type: list]
    :param mail_signature: raw string or html code to be put at the end of body [type: string, default='']
    :param auto_discover: flag for pointing to EWS using a specific protocol [type: bool, default=False]
    :param access_type: access type associated to the credentials provided [type: obj, default=DELEGATE]
 
    Return
    ------
    This function returns anything besides sending the configured mail
    """
    
    # Setting up account
    account = connect_exchange(username=username, password=password, server=server, mail_box=mail_box,
                               auto_discover=auto_discover, access_type=access_type)

    # Filtering and formating DataFrames to be sent on body
    meta_df_body = meta_df.query('flag_body == 1')
    if len(meta_df_body) > 0:
        html_df = pd.DataFrame(meta_df_body['df']).iloc[0, :]['df']
        html_body = format_html_body(mail_body, df=html_df, mail_signature=mail_signature)
    else:
        html_body = format_html_body(mail_body, mail_signature=mail_signature)
    
    # Creating a message object
    m = Message(account=account,
                subject=subject,
                body=html_body,
                to_recipients=mail_to)
    
    # Filtering and preparing DataFrames to be sent attached
    meta_df_attach = meta_df.query('flag_attach == 1')
    file_names = list(meta_df_attach['name'])
    file_dfs = list(meta_df_attach['df'])

    attach_dict = {file_names.index(name) + 1: {'name': name, 'df': df} for name, df in zip(file_names, file_dfs)}
    attachments = [buffer_dataframe(inner_dict['name'], inner_dict['df']) for idx, inner_dict in attach_dict.items()]

    # Attaching files
    for name, content in attachments or []:
        file = FileAttachment(name=name, content=content)
        m.attach(file)

    # Sending message
    m.send_and_save()

