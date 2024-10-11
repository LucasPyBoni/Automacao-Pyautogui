## Sobre Projeto:

Automação utilizando pyautogui, onde o arquivo excel é lido e para cada produto dentro desse arquivo, as informações suas colocadas dentro do Fakturama.


### Principais Bibliotecas:

Pandas, Pillow, Time, Pyautogui, SubProcess, entre outras.

### Linha de raciocínio

O Fakturama é aberto pelo subprocess.popen, enquanto ele não abre, o time sleep roda para não travar o código. São criadas funções p/ não repetir código, a encontrar_imagem procura a imagem pelo pyautogui.locateOnScreen e retorna a posição, a função escrever_texto recebe um texto, copia e cola por hotkeys p/ casos de palavras com ç e acento que não tem no teclado americano, já a função direita recebe uma posição mas clica direita dela, geralmente são onde fica o campo para digitar os textos, assim evitando de ficar pegando posição de tudo.
