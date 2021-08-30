<h1 align="center">
  <img src="https://i.imgur.com/Y5cCw3T.png", alt="jaiminho logo">
</h1>

## Table of content

- [Sobre o jaiminho](#sobre-o-jaiminho)
- [Instalação do Pacote](#instalação-do-pacote)
- [Funcionalidades Disponíveis](#funcionalidades-disponíveis)
- [Utilização Prática](#utilização-prática)
- [Contatos](#contatos)

___

## Sobre o jaiminho

Muito mais do que um pacato cidadão de [Tangamandápio](https://pt.wikipedia.org/wiki/Tangamand%C3%A1pio), o querido carteiro Jaiminho agora inspira implementações da comunidade de desenvolvimento em diversas frentes envolvendo envio de e-mails, entrega de pacotes, sistemas de pagamentos, sistemas de notificações, entre outras ferramentas. Neste repositório, a solução _jaiminho_ tem como principal objetivo facilitar o envio de e-mails em Python a partir da criação de módulos e funções encapsuladas com base em bibliotecas já existentes. É como o bordão clássico de evitar a fadiga e proporcionar ao usuário uma série de elementos já codificados para agilizar a construção de ferramentas que necessitem dessa troca de mensagens.

Hoje, a última versão do pacote _jaiminho_ já pode ser encontrada no repositório [PyPI](https://pypi.org/project/jaiminho/) e detalhes sobre novas funcionalidades poderão ser visualizadas neste repositório do GitHub. Adicionalmente, exemplos de utilização serão fornecidos para um melhor entendimento dos usuários que queiram consumir o pacote.

___

## Instalação do Pacote

Com o [ambiente virtual python](https://realpython.com/python-virtual-environments-a-primer/) ativo, para a instalação do pacote _jaiminho_ via pip, basta executar o comando abaixo:

```bash
pip install jaiminho
```

Com isso, todo o ferramental disponível na última versão do pacote poderá ser usufruído. Vale citar que o pacote _jaiminho_ para envio de e-mails possui algumas dependências associadas que são gerenciadas automaticamente no ato de sua instalação, sendo elas:
* `exchangelib`: client python para envio de e-mails utilizando o servidor Exchange da Microsoft
* `pandas`: poderosa ferramenta para a manipulação de dados em python
* `pretty-html-table`: módulo responsável por transformar objetos DataFrame do pandas em tabelas customizadas em HTML

___

## Funcionalidades Disponíveis

Até o presente momento, o pacote _jaiminho_ conta com o módulo `exchange.py` responsável por consolidar as principais operações de envio de e-mails utilizando, como base fundamental, a biblioteca `exchangelib` definida acima. Em sua versão mais recente, o referido módulo está estruturado em um formato de funções e entrega, a princípio, as seguintes funcionalidades:

| Função                      | Descrição                                                                                             |
| :-------------------------: | :---------------------------------------------------------------------------------------------------: |         
| `connect_to_exchange()`     | Realiza a conexão com o servidor Exchange a partir de credenciais fornecidas pelo usuário             |
| `create_message()`          | Utiliza uma conta conectada ao servidor Exchange para criar uma mensagem básica                       |
| `attach_file()`             | Gerencia o processo de anexação de arquivos a uma mensagem criada                                     |
| `df_to_html()`              | Transforma um objeto DataFrame em uma tabela HTML pré formatada a partir do pacote pretty-html-table  |
| `send_mail()`               | Encapsula os processos de criação de conta, mensagem, anexo (opcional) e envia o e-mail solicitado    |

Cada uma das funções acima listadas possuem uma documentação completa e que pode ser acessada diretamente no respectivo módulo.

___

## Utilização Prática

Visando propor um melhor entendimento sobre algumas das principais funcionalidades do pacote _jaiminho_, o código abaixo foi desenvolvido para utilizar credenciais fornecidas pelo usuário afim de conectar ao servidor Exchange e enviar uma mensagem simples a um destinatário.

```python
# Importando bibliotecas
import jaiminho.exchange as jex
import os
from dotenv import find_dotenv, load_dotenv

# Lendo variáveis de ambiente
load_dotenv(find_dotenv())

# Coletando variáveis
MAIL_USERNAME = os.getenv('MAIL_USERNAME')
MAIL_BOX = os.getenv('MAIL_BOX')
MAIL_TO = os.getenv('MAIL_TO')
MAIL_TO = [MAIL_TO] if MAIL_TO.count('@') == 1 else MAIL_TO.split(';')
SERVER = 'outlook.office365.com'

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
    
# Gerando e enviando mensagem simples
m = jex.create_message(
    account=acc,
    subject='[Jaiminho] exchange_tests.py [1]',
    body='1º teste de envio de e-mails com Jaiminho',
    to_recipients=MAIL_TO
)
m.send_and_save()
```

Como informado na seção anterior, o anexo de arquivos a um e-mail é uma das principais funcionalidades disponíveis para o pacote e, dessa forma, o código abaixo referencia o caminho de um arquivo qualquer, salvo localmente no sistema operacional, para anexação à mensagem a partir da função `attach_file()`:

```python
# Arquivo README.md salvo localmente
PROJECT_PATH = os.getcwd()
LOCAL_FILENAME = os.path.join(PROJECT_PATH, 'README.md')

# Anexando arquivo à mensagem criada pela função create_message()
m = jex.attach_file(
    message=m,
    file=LOCAL_FILENAME,
    attachment_name=os.path.basename(LOCAL_FILENAME)
)
m.send_and_save()
```

Por fim, propondo um encapsulamento e abstração ainda maior aos usuários finais, no exemplo abaixo será proposta um envio simples de e-mail a partir da função `send_mail()` que, por sua vez, realiza chamadas internas às demais funções do pacote para criação de conta, criação da mensagem, anexo de arquivos (opcional) e envio do e-mail configurado.

```python
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
```

Para mais informações, o arquivo `tests/exchange_tests.py` contempla os mesmos exemplos acima e outros adicionais para um melhor detalhamento das funcionalidades disponíveis. Seu consumo é recomendado para extrair poder máximo do pacote _jaiminho_.

___

## Contatos

* LinkedIn: https://www.linkedin.com/in/thiago-panini/
* Outros pacotes desenvolvidos: https://github.com/ThiagoPanini
