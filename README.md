<div align="center">
  <strong>Envio de E-mails via Exchange</strong>
</div>
<br/>

Biblioteca python construída para facilitar o gerenciamento e envio de e-mails utilizando a biblioteca `exchangelib` como ORM da caixa de e-mails Exchange.

## Instalação via git clone

Para utilização das funcionalidades encapsuladas neste repositório, é possível executar o comando `git clone` na máquina local.

```bash
$ git clone https://github.com/ThiagoPanini/xchange_mail.git 
```

Nesse cenário, o usuário pode usufruir de scripts disponibilizados no diretório `examples` para terem uma base de utilização das ferramentas deste repositório. Um dos exemplos de uso está presente em `examples/starlight`, sendo seu conteúdo é dado por:

```bash
└── starlight
    ├── depara_imgs.txt
    ├── depara_tags.txt
    ├── starlight.html
    └── starlight.py
```

O script `starlight.py`, desde configurado previamente com as credenciais do usuário executar do código, permite o envio de um e-mail altamente customizado com um template HTML com o tema Starlight criado a partir da plataforma <a href="https://beefree.io/">Bee Free</a>.

<h1 align="center">
  <img src="https://i.imgur.com/SeNZiJ6.png" alt="starlight_template">
</h1>

## Instalação via pip

Adicionalmente, é possível realizar a instalação do pacote diretamente do <a href="https://pypi.org/project/xchange-mail/0.0.1/">PyPI</a>. Para tal, é recomendada a criação de um ambiente virtual para o projeto ao qual exige a utilização das funcionalidades aqui expostas a partir da sequência de comandos abaixo:

```bash
# Criando e ativando ambiente virtual
$ python3 -m venv <nome_venv>
$ source ~/<nome_venv>/bin/activate

# Instalando biblioteca
pip install xchange_mail
```

Assim, em qualquer script Python, será possível importar os módulos da biblioteca da seguinte forma

```bash
from xchange_mail.handler import send_simple_mail
```
