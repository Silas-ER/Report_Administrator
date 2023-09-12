import sys
from cx_Freeze import setup, Executable

# Inclua aqui as bibliotecas que você está usando
build_exe_options = {
    "packages": ["datetime", "openpyxl", "smtplib", "email.mime", "pyodbc", "pandas"],
    "include_files": ["relatorios/", "cabecalho.py", "conexao_banco.py", "envio_email.py", "estilizar.py", "formulas.py"],
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Relatorio de lagostas para o ministerio da pesca",
    version="0.1",
    description="Cria Relatorios de entradas de lagostas atraves da conexão com o banco de dados sql fazendo filtros com scripts, importando para o excel e enviando por e-mail",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base)]
)
