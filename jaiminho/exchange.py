"""
---------------------------------------------------
---------------- MÓDULO: exchange -----------------
---------------------------------------------------
Dentro da proposta do pacote jaiminho, este módulo
tem por objetivo consolidar os principais elementos
para o gerenciamento de e-mails a partir do servidor
Exchange da Microsoft utilizando, como principal
ferramenta, a biblioteca exchangelib do Python.
Aqui, o usuário poderá encontrar diversas
funcionalidades responsáveis por encapsular boa 
parte do trabalho de construção de e-mails, envio
de anexo, preparação de HTML, entre outras features
comumente utilizadas no gerenciamento de e-mails.

Table of Contents
---------------------------------------------------
1. Configurações iniciais
    1.1 Importando bibliotecas
    1.2 Configurando logs
2. Encapsulando envio de e-mails
    2.1 Funções auxiliares
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

# Classes da biblioteca exchangelib
from exchangelib import Credentials, Account, Configuration, Message, \
                        FileAttachment, HTMLBody

# Bibliotecas gerais
from pandas import DataFrame
from io import BytesIO


"""
---------------------------------------------------
-------- 2. ENCAPSULANDO O ENVIO DE EMAILS --------
               2.1 Funções auxiliares
---------------------------------------------------
"""

# Conectando ao servidor Exchange
def connect_to_exchange(username, password, server, mail_box):
    """
    Providencia o retorno de uma conta configurada da Exchange
    a partir da utilização de credenciais válidas fornecidas
    pelo usuário. Na prática, esta função realiza a criação
    sequencial dos objetos Credentials, Configuration e 
    Account da biblioteca exchangelib, permitindo assim uma
    maior facilidade no manuseio dos elementos básicos de
    criação e configuração de conta no processo de gerenciamento
    de e-mails.

    Parâmetros
    ----------
    :param username:
        Usuário de e-mail com permissões válidas de envio de
        e-mails a partir da caixa genérica (mail_box) fornecida.
        Eventualmente, este parâmetro pode ser preenchido com
        o usuário de domínio Windows utilizado na autenticação
        do sistema de envio de e-mails pela Microsoft.
        [type: string]

    :param password:
        Senha referente ao usuário "username" fornecido para
        a devida autenticação no envio de e-mails.
        [type: string]

    :param server:
        Servidor de gerenciamento de e-mails utilizado nas
        operações a serem realizadas pelo objeto de conta
        criado. Um exemplo prático do servidor associado ao
        outlook do office 365 é: "outlook.office365.com"
        [type: string]

    :param mail_box:
        Caixa de e-mail a ser utilizada nas operações a serem
        designadas, sejam estas de leitura ou envio de e-mails.
        Basicamente, este parâmetro é equivalente ao parâmetro
        "primary_smtp_address" da classe Account e é a partir
        dele que as ações de e-mail são vinculadas.
        [type: string]

    Retorno
    -------
    :return account:
        Objeto do tipo Account considerando as credenciais
        fornecidas e uma configuração previamente estabelecida
        com o servidor e o endereço de SMTP primário também
        fornecidos.
        [type: Account]
    """

    # Configurando credenciais do usuário
    creds = Credentials(
        username=username, 
        password=password
    )

    # Configurando servidor com as credenciais fornecidas
    config = Configuration(
        server=server, 
        credentials=creds
    )

    # Criando objeto de conta com todo o ambiente já configurado
    account = Account(
        primary_smtp_address=mail_box, 
        credentials=creds, 
        config=config
    )
    
    return account

# Criando objeto de mensagem
def create_message(account, subject, body, to_recipients):
    """
    Consolida os elementos mais básicos para a criação de
    um objeto de mensagem a ser gerenciado externamente
    pelo usuário ou por demais funções neste módulo.
    Em linhas gerais, o código neste bloco instancia a 
    classe Message() da biblioteca exchangelib com os
    argumentos fundamentais para a construção de uma
    simples mensagem.

    Parâmetros
    ----------
    :param account:
        Objeto do tipo Account considerando as credenciais
        fornecidas e uma configuração previamente estabelecida
        com o servidor e o endereço de SMTP primário também
        fornecidos.
        [type: Account]

    :param subject:
        Título da mensagem ser enviada por e-mail.
        [type: string]

    :param body:
        Corpo da mensagem a ser enviada por e-mail. Este 
        argumento é automaticamente transformado em uma 
        string HTML a partir da aplicação da classe HTMLBody()
        antes da consolidação na classe Message().
        [type: string]

    Retorno
    -------
    :return m:
        Mensagem construída a partir da inicialização da classe
        Message() da biblioteca exchangelib. Tal mensagem 
        representa as configurações mais básicas de um e-mail 
        contendo uma conta configurada, um titulo, um corpo html 
        e uma lista válida de destinatários.
        [type: Message]
    """

    m = Message(
        account=account,
        subject=subject,
        body=HTMLBody(body),
        to_recipients=to_recipients
    )

    return m

def attach_file_to_message(message, file, attachment_name, is_inline=True):
    """
    """

    if type(file) is str:
        # Leitura de arquivo local em bytes
        with open(file, 'rb') as f:
            content = f.read()

    elif type(file) is DataFrame:
        # Salva objeto DataFrame em buffer para posterior leitura
        buffer = BytesIO()
        file.to_csv(buffer)
        content = buffer.getvalue()

    elif type(file) is bytes:
        content = file
    
    else: # Formato do anexo inválido
        return message

    # Criando objeto de anexo e incluindo na mensagem
    file = FileAttachment(
        name=attachment_name,
        content=content,
        is_inline=is_inline,
        content_id=attachment_name
    )
    message.attach(file)
    
    return message