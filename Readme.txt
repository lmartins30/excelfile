The script has comments to help but they are in Portuguese ( Se você falar português, sem problemas, é até melhor) so this file is to help if any command line commentary is necessary.

English: 
- Every variable with the '_d' in the end is for the original file and with the '_u' is for the new file you will create, if you want to 
For Example:
  - 'caminho_d' is the path of the origin excel file and 'caminho_u' is the path to the file you will create
  - 'planilha_d' is the name of the sheet on the origin file and 'planilha_u' is the path to the sheet you will create on the new file
Lines 24 - 28:
  - Takes the values from the excel file, in this case, the file have 4 column, you can have how many you need to.
Lines 31 - 35:
  - Writes the data in the new file, as said above, you can have how many you need. In this case we just copied from one to another with an extra column.
The function 'checa_data' is a date check function. The date is the first column of the file and we check it based in the date the script is running.
  - hoje means today
  - delta is the difference between today and the date on the file.
  - The first line is to take just the date and not the time, to make operations easier and faster.

Português:
- As variáveis que contém '_d' são do arquivo base e as que tem '_u' são do arquivo que será criado, se necessário.
- Linhas 24 a 28 pegam os dados das colunas e linhas 31 a 35 escrevem os dados.

Script desenvolvido por Luiz Antonio Martins,
qualquer dúvida ou sugestão: luizantoniomartins30@gmail.com
