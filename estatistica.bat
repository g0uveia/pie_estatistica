@echo off
set /p filename="Nome do arquivo: "

python src/1.py "%filename%"
python src/2.py "%filename%"
python src/3.py "%filename%"
python src/4.py "%filename%"
python src/5.py "%filename%"

