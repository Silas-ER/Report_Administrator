## Construção de relatórios através de script
### Building reports through script

Script gerado para criação de relatório especifico solicitado pelo Ministério da Agricutura e Pecuária, com o objetivo específico de otimizar o processo de escrita que era feito manualmente, facilitando a construção e padronização para envio. 

Script generated to create a specific report requested by the Ministry of Agriculture and Livestock, with the specific objective of optimizing the writing process that was done manually, facilitating construction and standardization for sending.

Por padrão com os outros scripts já gerados na empresa e pela familiaridade com a linguagem, optei por utilizar o Python. Utilizei as libs: 
<ul>
  <li>datetime (Coleta de data, utilizada no código para consulta e nome de arquivos na exportação dos dados);</li>
  <li>openpyxl (Criação de planilhas, formatação e padronização das mesmas);</li>
  <li>smtplib (Para conexão com o protocolo SMTP);</li>
  <li>email.mime (Criação de email personalizado e anexo de arquivos);</li>
  <li>pyodbc (Utilizada para a conexão com o banco de dados, nesse caso o MS-SQL);</li>
  <li>pandas (Foi utilizada para leitura dos dados do SQL).</li>
</ul>

<p>
  Para criação do executavél utilizei a lib Cx_freeze que está configurada através do setup.py e programei a execução com o agendador de tarefas do windows.
  Por fim tentei ao máximo modularizar o código para facilitar a manutenção. 
</p>

