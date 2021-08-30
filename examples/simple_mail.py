"""
---------------------------------------------------
-------------- SCRIPTS: simple_mail ---------------
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
import jaiminho.exchange as jex

# Bibliotecas padrão
import os
from dotenv import find_dotenv, load_dotenv

# Logging
import logging


"""
---------------------------------------------------
------------ 1. CONFIGURAÇÕES INICIAIS ------------
               1.2 Configurando log
---------------------------------------------------
"""

# Definindo função para configurar objeto de log do código
def log_config(logger, level=logging.DEBUG, 
               log_format='%(levelname)s;%(asctime)s;%(filename)s;%(module)s;%(lineno)d;%(message)s',
               log_filepath=os.path.join(os.getcwd(), 'exec_log/execution_log.log'),
               flag_file_handler=False, flag_stream_handler=True, filemode='a'):
    """
    Função que recebe um objeto logging e aplica configurações básicas ao mesmo
    
    Parâmetros
    ----------
    :param logger: objeto logger criado no escopo do módulo [type: logging.getLogger()]
    :param level: level do objeto logger criado [type: level, default=logging.DEBUG]
    :param log_format: formato do log a ser armazenado [type: string]
    :param log_filepath: caminho onde o arquivo .log será armazenado 
        [type: string, default='exec_log/execution_log.log']
    :param flag_file_handler: define se será criado um arquivo de armazenamento de log
        [type: bool, default=False]
    :param flag_stream_handler: define se as mensagens de log serão mostradas na tela
        [type: bool, default=True]
    :param filemode: tipo de escrita no arquivo de log [type: string, default='a' (append)]
    
    Retorno
    -------
    :return logger: objeto logger pré-configurado
    """

    # Setting level for the logger object
    logger.setLevel(level)

    # Creating a formatter
    formatter = logging.Formatter(log_format, datefmt='%Y-%m-%d %H:%M:%S')

    # Creating handlers
    if flag_file_handler:
        log_path = '/'.join(log_filepath.split('/')[:-1])
        if not os.path.isdir(log_path):
            os.makedirs(log_path)

        # Adding file_handler
        file_handler = logging.FileHandler(log_filepath, mode=filemode, encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    if flag_stream_handler:
        # Adding stream_handler
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)    
        logger.addHandler(stream_handler)

    return logger

# Instanciando e configurando objeto de log
logger = logging.getLogger(__file__)
logger = log_config(logger)


"""
---------------------------------------------------
------------ 1. CONFIGURAÇÕES INICIAIS ------------
        1.3 Definindo variáveis do projeto
---------------------------------------------------
"""

# Lendo variáveis de ambiente
load_dotenv(find_dotenv())

# Coletando variáveis
USERNAME = os.getenv('USERNAME')
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = os.getenv('MAIL_TO')
MAIL_TO = [MAIL_TO] if MAIL_TO.count('@') == 1 else MAIL_TO.split(';')