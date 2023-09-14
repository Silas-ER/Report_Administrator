## Construção de relatórios através de script
### Building reports through script
<p>
Script gerado para criação de <a href="https://github.com/Silas-ER/Report_Administrator/blob/main/Consulta%20e%20envio%20de%20relat%C3%B3rios%20por%20e-mail/relatorios">relatório</a> especifico solicitado pelo Ministério da Agricutura e Pecuária, com o objetivo específico de otimizar o processo de escrita que era feito manualmente, facilitando a construção e padronização para envio. 
</p>

<p>
Script generated to create a specific <a href="https://github.com/Silas-ER/Report_Administrator/blob/main/Consulta%20e%20envio%20de%20relat%C3%B3rios%20por%20e-mail/relatorios">report</a> requested by the Ministry of Agriculture and Livestock, with the specific objective of optimizing the writing process that was done manually, facilitating construction and standardization for sending.
</p>

<p>
Por padrão com os outros scripts já gerados na empresa e pela familiaridade com a linguagem, optei por utilizar o Python. Utilizei as libs: 
</p>

<ul>
  <li>datetime (Coleta de data, utilizada no código para consulta e nome de arquivos na exportação dos dados);</li>
  <li>openpyxl (Criação de planilhas, formatação e padronização das mesmas);</li>
  <li>smtplib (Para conexão com o protocolo SMTP);</li>
  <li>email.mime (Criação de email personalizado e anexo de arquivos);</li>
  <li>pyodbc (Utilizada para a conexão com o banco de dados, nesse caso o MS-SQL);</li>
  <li>pandas (Foi utilizada para leitura dos dados do SQL).</li>
</ul>

<p> 
  By default with the other scripts already generated in the company and due to my familiarity with the language, I chose to use Python. I used the libs:
</p>

<ul>
  <li>datetime (Collection of date, used in the code for querying and file names when exporting data);</li>
  <li>openpyxl (Creation of spreadsheets, formatting and standardization);</li>
  <li>smtplib (For connection to the SMTP protocol);</li>
  <li>email.mime (Creation of personalized email and file attachments);</li>
  <li>pyodbc (Used to connect to the database, in this case MS-SQL);</li>
  <li>pandas (It was used to read SQL data).</li>
</ul>

<p>
  Para criação do executavél utilizei a lib Cx_freeze que está configurada através do setup.py e programei a execução com o agendador de tarefas do windows.
  <br>
  Por fim tentei ao máximo modularizar o código para facilitar a manutenção. 
</p>

<p>
  To create the executable I used the Cx_freeze lib that is configured through setup.py and programmed the execution with the Windows task scheduler.
  <br>
  Finally, I tried my best to modularize the code to facilitate maintenance.
</p>