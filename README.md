# Conversor de arquivo Excel para arquivo TXT para importação de dados no programa RPC da ANS

### O objetivo desse projeto é facilitar no processo de preenchimento dos dados para enviou do RPC junto a ANS, onde era efetuado o preenchimento de forma manualmente através de dados que foram previamente coletos e armazenadas em um arquivo xlsx, onde eram inseridos um à um no programa. O conversor irá usar esse mesmo arquivo xlsx para criar um arquivo no formato TXT com layout previsto pelo manual da ANS e que poderá ser importado no programa, gerando um ganho operacional enorme.
<br>

* A biblioteca pandas foi utilizada para trabalhar melhor o relatório que será importado.
* Foi utilizado o Pyinstaller para criar um arquivo excecutável, e assim, não será necessário ter o Python instalado no computador para poder utilizar o programa.
* Os campos do Header (registro C1) são imputados pelo operador, facilitando assim a utilização.
* O registro C4 está desabilitado, pois as informações contidas nele não eram preenchidas no processo manual.
* Algumas linhas do código precisam de melhorias, a princípio o código foi escrito no Jupyter Notebook e posteriormente transferido para o Pycharm.
<br>
As informações usadas para criar o layout do arquivo txt foram retiradas no manual do RPC, disponibilizado pela ANS (versão 3.1.4 de 31/08/2016) que está disponível no site:https://www.gov.br/ans/pt-br/centrais-de-conteudo/manuais-do-portal-operadoras/reajuste-de-planos-coletivos-rpc .

<br><br>

<b> Layout do arquivo TXT </b> <br><br>

![Layout_1](https://user-images.githubusercontent.com/86494924/148312778-59342f24-38c2-45de-ad6a-a1ca918cf0a0.jpg)

![Layout_2](https://user-images.githubusercontent.com/86494924/148312976-b2af5a9a-4c36-4247-abee-d4f4b3a8025b.png)

![Layout_3](https://user-images.githubusercontent.com/86494924/148313029-de2b3679-966f-48f2-a678-956f6949c1b3.png)

![Layout_4](https://user-images.githubusercontent.com/86494924/148313097-27f0da9d-3a49-4e94-a06d-ea2ca49639ea.png)

![Layout_5](https://user-images.githubusercontent.com/86494924/148313140-0f39797f-0b11-44d5-8850-22cc16473292.png)

![Layout_6](https://user-images.githubusercontent.com/86494924/148313179-f472c4e8-6d4a-49c8-beb8-4979734fa33e.png)

![Layout_7](https://user-images.githubusercontent.com/86494924/148313208-a716e098-b654-457c-997b-f03c3c9abcf1.png)

<br><br>

<b>Modelo arquivo xlsx</b> <br><br>

[modelo.xlsx](https://github.com/MarceloPalumbo/rpc/files/7818806/modelo.xlsx)

<br><br>
<b>Visualização do arquivo excecutável</b> <br><br>

![exe](https://user-images.githubusercontent.com/86494924/148314486-96627ec7-b04d-4bd2-a55a-ec1c254a2127.png)


