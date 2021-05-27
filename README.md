<h1 align="center">
  <a href="https://pypi.org/project/xchange-mail/"><img src="https://i.imgur.com/ISexIyT.png" alt="xchange_mail logo"></a>
</h1>

<div align="center">
  <strong>📧 Sending emails with basic formatting using exchangelib 📧</strong>
</div>
<br/>


<div align="center">  
 
  [![PyPI](https://img.shields.io/pypi/v/xchange_mail?color=blueviolet)](https://pypi.org/project/xchange-mail/)
  ![PyPI - Python Version](https://img.shields.io/pypi/pyversions/xchange_mail?color=green)
  ![PyPI - Status](https://img.shields.io/pypi/status/xchange_mail)

</div>


<div align="center">  
  
  ![Downloads](https://img.shields.io/pypi/dm/xchange_mail?color=darkblue)
  ![Downloads](https://img.shields.io/pypi/dw/xchange_mail?color=blue)
  ![Downloads](https://img.shields.io/pypi/dd/xchange_mail?color=lightblue)

</div>
<br/>

## Table of contents

- [About xchange_mail](#about-xchange_mail)
- [Package Structure](#package-structure)
- [Installing the Package](#installing-the-package)
- [Examples](#examples)
- [Contribution](#contribution)
- [Social Media](#social-media)

___

## About xchange_mail

This python package was build for making the mail sending processing through exchangelib a little bit easier. The idea is to create some custom functions for a limited use cases, so the user won't need to configure or define details on Account or Config exchangelib classes, but rather execute basic functions for sending basic emails.

The **examples** section will clarify some use cases of `xchange_mail` package for helping users to send basic emails through MS Exchange. Keep watching this documentation.

## Package Structure

At this time, the package is built around just one module called `mail`. This module contains some functions for helping users connecting with Exchange server and also sending basic mails with plain text or HTML body messages. The table below has the explanation of the main componentes of this `mail` module.

| Function                | Short Description                                                                         |
| :---------------------: | :---------------------------------------------------------------------------------------: |
| `connect_exchange()`    | Receives some user credentials for connecting to Exchange and returning an Account object |
| `attach_file()`         | Stores a pandas DataFrame object on buffers and returns a two-elements list containing the attach name and the attach object |
| `format_mail_body()`    | Creates a HTMLBody object. If a DataFrame is passed as an argument, it uses `pretty_html_table` package for customizing a table before creating the HTMLBody |
| `send_simple_mail()`    | Sends a simple mail through exchange with possibilities for attaching one file, sending a DataFrame object on mail body, sending an image on mail body or attached or using html code for customizing mail |
| `send_mail_mult_files()` | Can send multiple files attached or multiple DataFrames on body |

Biblioteca python construída para facilitar o gerenciamento e envio de e-mails utilizando a biblioteca `exchangelib` como ORM da caixa de e-mails Exchange.

___

## Installing the Package

The latest version of `xchange_mail` package are published and available on [PyPI repository](https://pypi.org/project/xchange-mail/)

> :pushpin: **Note:** as a good practice for every Python project, the creation of a <a href="https://realpython.com/python-virtual-environments-a-primer/">virtual environment</a> is needed to get a full control of dependencies and third part packages on your code. By this way, the code below can be used for creating a new venv on your OS.
> 

```bash
# Creating and activating venv on Linux
$ python -m venv <path_venv>/<name_venv>
$ source <path_venv>/<nome_venv>/bin/activate

# Creating and activating venv on Windows
$ python -m venv <path_venv>/<name_venv>
$ <path_venv>/<nome_venv>/Scripts/activate
```

With the new venv active, all you need is execute the code below using pip for installing the package (upgrading pip is optional):

```bash
$ pip install --upgrade pip
$ pip install xchange_mail
```

The xchange_mail package is built in an upper layer above some other python packages like exchangelib and pandas. So, when installing mlcomposer, the pip utility will also install all dependencies linked to the package.

## Examples

After introducing the package, it's time to explain it in a deeper way: through examples. On this Github repository, it's possible to find some good uses of xchange_mail on `examples/` folder. In practice, for sending a basic email it's possible to execute the `send_simple_mail()` function with few parameter configuration as seen below:

```python
from xchange_mail.mail import send_simple_mail

# Extracting environment variables from a .env file (optional)
USERNAME = os.getenv('MAIL_FROM')
PWD = os.getenv('PASSWORD')
SERVER = 'outlook.office365.com'
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = os.getenv('MAIL_TO')

# Sending a basic mail
send_simple_mail(username=USERNAME,
                 password=PWD,
                 server=SERVER,
                 mail_box=MAIL_BOX,
                 subject='This is a xchange_mail test',
                 mail_body='Testing the package by sending a simple mail',
                 mail_signature='Regards, xchange_mail developers',
                 mail_to=MAIL_TO)
```

Done! Almost all other package features are built around this `send_simple_mail()` function and the other one called `send_mail_mult_files()`. Just to clarify, there are some parameters that can be set on the function above for sending a pandas DataFrame attached on mail body, for example. There is also a feature for sending an image embedding on mail body. The code below is an example of sending a simple mail with a DataFrame object attached, on mail body with an image saved locally.

```python
import pandas as pd
from xchange_mail.mail import send_simple_mail

# Extracting environment variables from a .env file (optional)
USERNAME = os.getenv('MAIL_FROM')
PWD = os.getenv('PASSWORD')
SERVER = 'outlook.office365.com'
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = os.getenv('MAIL_TO')

# Sending a basic mail
send_simple_mail(username=USERNAME,
                 password=PWD,
                 server=SERVER,
                 mail_box=MAIL_BOX,
                 subject='This is a xchange_mail test',
                 mail_body='Testing the package by sending a simple mail',
                 mail_signature='Regards, xchange_mail developers',
                 mail_to=MAIL_TO,
                 df_on_body=True,
                 df_on_attachment=True,
                 df=df,
                 attachment_filename='pandas_dataframe.csv',
                 image_on_body=True,
                 image_location='/home/user/image_dir/image.png')
```

For new use cases, please take a look at `examples/` folder on this repository.


## Contribution

The xchange_mail python package is an open source implementation and the more people use it, the more happy the developers will be. So if you want to contribute with xchange_mail, please feel free to follow the best practices for implementing coding on this github repository through creating new branches, making merge requests and pointig out whenever you think there is a new topic to explore or a bug to be fixed.

Thank you very much for reaching this and it will be a pleasure to have you as xchange_mail user or developer.

___

## Social Media

* Follow me on LinkedIn: https://www.linkedin.com/in/thiago-panini/
* See my other Python packages: https://github.com/ThiagoPanini
