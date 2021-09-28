"""
---------------------------------------------------
-------------- TESTS: exchange_tests --------------
---------------------------------------------------


Table of Contents
---------------------------------------------------
1. Configurações iniciais
    1.1 Importando bibliotecas
    1.2 Definindo variáveis do projeto
2. Testando funcionalidades
    2.1 Criação básica de conta e mensagem
    2.2 Anexo de arquivos locais 
    2.3 Anexo de arquivo após leitura em memória
    2.4 Leitura e anexo de arquivo direto em memória
    2.5 Enviando DataFrames no corpo do e-mail
    2.6 Envio de e-mail utilizando função única
    2.7 Utilizando função única com anexos
    2.8 Utilizando função única com imagem no body
---------------------------------------------------
"""

# Author: Thiago Panini
# Date: 27/09/2021


"""
---------------------------------------------------
------------ 1. CONFIGURAÇÕES INICIAIS ------------
            1.1 Importando bibliotecas
---------------------------------------------------
"""

# Funcionalidades
import jaiminho.exchange as jex
from exchangelib.errors import UnauthorizedError

# Bibliotecas padrão
import os
from dotenv import find_dotenv, load_dotenv
import pandas as pd
import requests


"""
---------------------------------------------------
------------ 1. CONFIGURAÇÕES INICIAIS ------------
        1.2 Definindo variáveis do projeto
---------------------------------------------------
"""

# Lendo variáveis de ambiente
load_dotenv(find_dotenv())

# Definindo variávies de diretório
WORK_PATH = os.getenv('WORK_PATH')
PROJECT_PATH = os.getcwd()

# Coletando variáveis
MAIL_USERNAME = os.getenv('MAIL_USERNAME')
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = os.getenv('MAIL_TO')
MAIL_TO = [MAIL_TO] if MAIL_TO.count('@') == 1 else MAIL_TO.split(';')
SERVER = 'outlook.office365.com'


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
      2.1 Criação básica de conta e mensagem
---------------------------------------------------
"""

# Conectando ao servidor e obtendo conta
try:
    acc = jex.connect_to_exchange(
        username=MAIL_USERNAME,
        password=os.getenv('PASSWORD'),
        server=SERVER,
        mail_box=MAIL_BOX
    )
except UnauthorizedError as ue:
    print(f'Erro de autorização ao realizar login no servidor')
    raise ue

# Gerando mensagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [1] - Mensagem Simples',
    body='Criando e enviando mensagem simples por e-mail via connect_exchange() e create_message()',
    to_recipients=MAIL_TO
)
m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
           2.2 Anexo de arquivos locais 
---------------------------------------------------
"""

# Gerando mensagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [2] - Anexo Local',
    body='Enviando e-mail anexando arquivo local via attach_file()',
    to_recipients=MAIL_TO
)

# Anexando arquivo local à imagem
LOCAL_FILENAME = os.path.join(PROJECT_PATH, 'README.md')
m = jex.attach_file(
    message=m,
    file=LOCAL_FILENAME,
    attachment_name=os.path.basename(LOCAL_FILENAME)
)
assert len(m.attachments) >= 1, 'Anexo de arquivo falhou. Mensagem não possui elementos anexos mesmo após chamada da função'

# Definição de caminhos para diferentes arquivos a serem anexados
IMG_FILENAME = os.getenv('IMG_FILENAME')
CSV_FILENAME = os.getenv('CSV_FILENAME')
TXT_FILENAME = os.getenv('TXT_FILENAME')
PYTHON_FILENAME = os.getenv('PYTHON_FILENAME')
PDF_FILENAME = os.getenv('PDF_FILENAME')
PPT_FILENAME = os.getenv('PPT_FILENAME')

# Gerando lista completa e iterando para anexos individuais
ATTACHMENTS = [IMG_FILENAME, CSV_FILENAME, TXT_FILENAME, PYTHON_FILENAME, PDF_FILENAME, PPT_FILENAME]
for file in ATTACHMENTS:
    m = jex.attach_file(
    message=m,
    file=file,
    attachment_name=os.path.basename(file)
)
assert len(m.attachments) == len(ATTACHMENTS) + 1, f'Anexo de múltiplos arquivos falhou. Total de anexos ({len(m.attachments)}) difere do esperado ({len(ATTACHMENTS) + 1})'
m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
    2.3 Anexo de arquivo após leitura em memória
---------------------------------------------------
"""

# Leitura de imagem e DataFrame em memória para anexo
with open(IMG_FILENAME, 'rb') as f:
    img = f.read()
df = pd.read_csv(CSV_FILENAME)
MEM_ATTACHMENTS = [img, df]
PATHS = [IMG_FILENAME, CSV_FILENAME]

# Gerando mensagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [3] - Objetos em Memória',
    body='Enviando e-mail após leitura e anexo de DataFrame e imagem em memória',
    to_recipients=MAIL_TO
)

# Anexando arquivos em memória
for name, file in zip(PATHS, MEM_ATTACHMENTS):
    m = jex.attach_file(
    message=m,
    file=file,
    attachment_name=os.path.basename(name)
)
assert len(m.attachments) == len(MEM_ATTACHMENTS), f'Anexo de dados em memória falou. Total de anexos ({len(m.attachments)}) difere do esperado ({len(MEM_ATTACHMENTS)})'
m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
 2.4 Leitura e anexo de arquivo direto em memória
---------------------------------------------------
"""

# Aplicando requests online de imagem
url = 'https://ak-static.cms.nba.com/wp-content/uploads/headshots/nba/latest/260x190/203081.png'
r = requests.get(url)
img = r.content

# Gerando mensagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [4] - Imagem Baixada e Anexada',
    body='Enviando e-mail após requisição online de imagem e anexo direto da memória do conteúdo em bytes',
    to_recipients=MAIL_TO
)

# Anexando arquivo
m = jex.attach_file(
    message=m,
    file=img,
    attachment_name='Damian Lillard.png'
)
assert len(m.attachments) == 1, f'Anexo de dados em memória falou. Total de anexos ({len(m.attachments)}) deveria ser apenas 1'
m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
    2.5 Enviando DataFrames no corpo do e-mail
---------------------------------------------------
"""

# Preparando e transformando DataFrame
df_filter = df.loc[:5, ['player_name', 'player_team', 'game_date', 'matchup', 'wl']]
df_html = jex.df_to_html(df=df_filter)

# Gerando imagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [5] - DataFrame no Body',
    body='Envio de e-mail utilizando DataFrame formatado no body\n\n' + df_html,
    to_recipients=MAIL_TO
)
m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
    2.6 Envio de e-mail utilizando função única
---------------------------------------------------
"""

# Enviando e-mail sem nenhum tipo de anexo
jex.send_mail(
    username=MAIL_USERNAME,
    password=os.getenv('PASSWORD'),
    server=SERVER,
    mail_box=MAIL_BOX,
    mail_to=MAIL_TO,
    subject='[Jaiminho] exchange_tests.py [6] - Método Consolidado',
    body='Enviando e-mail simples a partir de método consolidado send_mail()',
    send=True
)


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
      2.7 Utilizando função única com anexos
---------------------------------------------------
"""

# Preparando zip de anexo para envio (nomes e arquivos)
FILENAMES = ['Damian Lillard.png', 'nba_players_stats.csv']
CONTENTS = [img, CSV_FILENAME]
attachments = zip(FILENAMES, CONTENTS)

# Enviando e-mail com anexos
jex.send_mail(
    username=MAIL_USERNAME,
    password=os.getenv('PASSWORD'),
    server=SERVER,
    mail_box=MAIL_BOX,
    mail_to=MAIL_TO,
    subject='[Jaiminho] exchange_tests.py [7] - Método Consolidado com Anexo',
    body='Enviando e-mail a partir de método consolidado send_mail() com anexo em memória e local + DataFrame no body\n\n' + df_html,
    zip_attachments=attachments,
    send=True
)

"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
   2.8 Utilizando função única com imagem no body
---------------------------------------------------
"""

# Enviando e-mail com anexos
jex.send_mail(
    username=MAIL_USERNAME,
    password=os.getenv('PASSWORD'),
    server=SERVER,
    mail_box=MAIL_BOX,
    mail_to=MAIL_TO,
    subject='[Jaiminho] exchange_tests.py [8] - Método Consolidado com Imagem no Body',
    body='Enviando e-mail a partir de método consolidado send_mail() com imagem e DataFrame no body:\n\n<img src="cid:Damian Lillard.png">\n\n' + df_html,
    zip_attachments=zip(['Damian Lillard.png'], [img]),
    send=True
)