@echo off
color e0
echo Subindo com os arquivos alterados parea o Git ...
Git add .
echo Informe a descricao das alteracoes
@set /p var=
Git commit –m "Incluir teste"
Git push
echo on