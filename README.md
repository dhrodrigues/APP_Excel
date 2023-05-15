## Desafio trabalho

Este código foi desenvolvido base em um VBA utilizado para transferências de dados de uma planilha Excel para um arquivo Word. O desafio era utilizar as mesmas características do VBA, porém deveria rodar em todos os sistemas operacionais.
Este código conta com a utilização de tkinter, que é uma biblioteca que possibilita a utilização de uma interface gráfica onde o usuário poderá selecionar onde está seu arquivo para realização de tal transferência e salva-lo após.
A biblioteca tqdm mostra o progresso do processo em um todo.
Utilizei a biblioteca Pyinstaller onde posso criar um executável do código sem precisar abrir o terminal toda vez que vou precisar executar essa aplicação.

## Título e Descrição
Aplicação de transferência Excel para Word. Esta aplicação foi baseada em um VBA utilizado. Porém, devido à falta de extensão do VBA em outros sistemas operacionais como MacOS e Linux, esta aplicação foi criada para dar maior agilidade e poder de processamento em todos os O.S.

## Instalação
## Para instalação, basta utilizar:
pip freeze > requirements.txt

Assim todas as bibliotecas serão instaladas para execução. Após utilizar o Pyinstaller para criar o arquivo executável da aplicação.


## Utilização


Esta aplicação é utilizada quando há necessidade de transferência de dados que encontram-se em um arquivo Excel para um arquivo Word que possuem tabelas iguais.

![Exemplo do Excel](excel.png)
![Exemplo do Word](word.png)
