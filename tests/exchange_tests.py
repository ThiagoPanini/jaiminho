"""
---------------------------------------------------
-------------- TESTS: exchange_tests --------------
---------------------------------------------------


Table of Contents
---------------------------------------------------
1. Configurações iniciais
    1.1 Importando bibliotecas
    1.2 Configurando logs
    1.3 Definindo variáveis do projeto
2. Gerenciando buckets e objetos
    2.1 Criando e configurando bucket
    2.2 Realizando upload de objetos
    2.3 Realizando o download de objetos
---------------------------------------------------
"""

# Author: Thiago Panini
# Date: 29/08/2021


"""
---------------------------------------------------
------------ 1. CONFIGURAÇÕES INICIAIS ------------
            1.1 Importando bibliotecas
---------------------------------------------------
"""

# Códigos úteis AWS
from exchangelib.properties import HTMLBody
import jaiminho.exchange as jex
from exchangelib.errors import UnauthorizedError

# Bibliotecas padrão
import os
from dotenv import find_dotenv, load_dotenv
import pandas as pd


"""
---------------------------------------------------
------------ 1. CONFIGURAÇÕES INICIAIS ------------
        1.2 Definindo variáveis do projeto
---------------------------------------------------
"""

# Definindo variávies de diretório
ROOT_PATH = os.path.expanduser('~')
WORK_PATH = os.path.join(ROOT_PATH, 'workspaces')
PROJECT_PATH = os.getcwd()

# Lendo variáveis de ambiente
load_dotenv(find_dotenv())

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
    print(f'Erro de autorização ao realizar login no servidor. Exception: {ue}')
    exit()

# Gerando imagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [1]',
    body='1º teste de envio de e-mails com Jaiminho',
    to_recipients=MAIL_TO
)
#m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
      2.2 Anexo de arquivos locais e em memória
---------------------------------------------------
"""

# Gerando imagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [2]',
    body='2º teste de envio de e-mails com Jaiminho',
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
IMG_FILENAME = os.path.join(WORK_PATH, 'nbaflow/dev/data/images/players/Damian Lillard.png')
CSV_FILENAME = os.path.join(WORK_PATH, 'nbaflow/dev/data/backup/2020-21_gamelog.csv')
TXT_FILENAME = os.path.join(PROJECT_PATH, 'requirements.txt')
PYTHON_FILENAME = os.path.join(PROJECT_PATH, 'setup.py')
PDF_FILENAME = os.path.join(WORK_PATH, 'voice-unlocker/auxiliar/tg_reconhecimento_voz_thiago_panini.pdf')
PPT_FILENAME = os.path.join(WORK_PATH, 'voice-unlocker/auxiliar/ppt/modelos/Computer Science Proposal by Slidesgo.pptx')

# Gerando lista completa e iterando para anexos individuais
ATTACHMENTS = [IMG_FILENAME, CSV_FILENAME, TXT_FILENAME, PYTHON_FILENAME, PDF_FILENAME, PPT_FILENAME]
for file in ATTACHMENTS:
    m = jex.attach_file(
    message=m,
    file=file,
    attachment_name=os.path.basename(file)
)
assert len(m.attachments) == len(ATTACHMENTS) + 1, f'Anexo de múltiplos arquivos falhou. Total de anexos ({len(m.attachments)}) difere do esperado ({len(ATTACHMENTS) + 1})'
#m.send_and_save()

# Leitura de imagem e DataFrame em memória para anexo
with open(IMG_FILENAME, 'rb') as f:
    img = f.read()
df = pd.read_csv(CSV_FILENAME)
MEM_ATTACHMENTS = [img, df]
PATHS = [IMG_FILENAME, CSV_FILENAME]

# Gerando imagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [3]',
    body='3º teste de envio de e-mails com Jaiminho',
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
#m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
    2.3 Enviando DataFrames no corpo do e-mail
---------------------------------------------------
"""

# Preparando e transformando DataFrame
df_filter = df.loc[:5, ['player_name', 'player_team', 'game_date', 'matchup', 'wl']]
df_html = jex.df_to_html(df=df_filter)

# Gerando imagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [4]',
    body='4º teste de envio de e-mails com Jaiminho\n\n' + df_html,
    to_recipients=MAIL_TO
)
#m.send_and_save()


"""
---------------------------------------------------
----------- 2. TESTANDO FUNCIONALIDADES -----------
    2.4 Envio de e-mail utilizando função única
---------------------------------------------------
"""

# Enviando e-mail sem nenhum tipo de anexo
jex.send_mail(
    username=MAIL_USERNAME,
    password=os.getenv('PASSWORD'),
    server=SERVER,
    mail_box=MAIL_BOX,
    mail_to=MAIL_TO,
    subject='[Jaiminho] exchange_tests.py [5]',
    body='5º teste de envio de e-mails com Jaiminho'
)

# Preparando zip de anexo para envio (nomes e arquivos)
attachments = zip(PATHS, MEM_ATTACHMENTS)

# Enviando e-mail com anexos
jex.send_mail(
    username=MAIL_USERNAME,
    password=os.getenv('PASSWORD'),
    server=SERVER,
    mail_box=MAIL_BOX,
    mail_to=MAIL_TO,
    subject='[Jaiminho] exchange_tests.py [6]',
    body='6º teste de envio de e-mails com Jaiminho',
    zip_attachments=attachments
)

