# Bot MASSA Whatsapp

Este é um bot para enviar mensagens em massa via Whatsapp (realmente é massa) para você que precisa encaminhar uma mensagem personalizada a muitos números, seja para sua empresa, familia, amigos ou faculdade

## Como funciona?

Esse é um  [RPA](https://www.totvs.com/blog/inovacoes/rpa/), um robô que faz um processo automático no computador, assumindo controle do navegador, teclado e mouse do usuário, mas pode ficar tranquilo pois você ainda pode utilizar o computador enquanto nosso Bot Massa está em execução.
A base do funcionamento é a ferramenta [Selenium](https://www.selenium.dev), uma biblioteca de teste automatizado. 

### Chromedriver

Confira a versão do seu google chrome antes de iniciar a execução, caso a versão do chromedriver seja diferente não será possível fazer o controle.

Baixe a versão correta do chromedriver no site oficial da Google: https://googlechromelabs.github.io/chrome-for-testing/

## Utilização

Parametrização da planilha

- O arquivo precisa ser em Excel
- Deve existir uma coluna chamada CELULAR com os números de whatsapp
- Deve conter apenas uma aba
- Não pode conter filtros
- Deixe na planilha apenas as linhas que serão enviadas

Envio das mensagens

Primeiro, escreva a mensagem que será enviada ou importe de um 
arquivo de texto.

Aqui você pode personalizar as mensagens da seguinte forma:

Para inserir um valor dinâmico dentro da mensagem, escreva a tag 
com o mesmo nome da coluna na sua planilha
dentro da marcação {{% nome da coluna %}}. Exemplo:

Olá {{%NOME%}}! Sua nota na prova foi {{%NOTA%}}

Desta forma a mensagem irá ser montada com o valor da respectiva coluna.

Também é possível anexar uma imagem à mensagem, selecionando no local indicado

Após configurar a mensagem, escolha a planilha em excel com os dados
e clique em Iniciar

O navegador será aberto no Whatsapp Web. Siga as instruções do site para 
conectar um dispositivo, e após a conexão o bot começará a ser executado

## Executável

Para gerar o executável, recomendo utilizar a biblioteca [PyInstaller](https://pyinstaller.org/en/stable/). O arquivo spec está disponível no repositório, assim como o comando para execução pelo Prompt

### Com SPEC
`pyinstaller BotMassaWhatsapp.spec`

### Command prompt

` pyinstaller --onefile --name BotMassaWhatsapp.exe .\bot_whatsapp.py --icon data\botwhatsapp.ico --noconsole --add-data=data\chromedriver.exe:data`