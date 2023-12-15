@echo off
color e0
echo Subindo com os arquivos alterados parea o Git ...
Git add .
echo Informe a descricao das alteracoes
echo.
@set /p var=
Git commit –m "%var%"
Git push
echo on