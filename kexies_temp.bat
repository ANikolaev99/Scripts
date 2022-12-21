@echo off
:: Этот блок обьявляет переменные
set "d1=C:\Works\Temp\scripts\log.txt"
set "d2=C:\Works\Temp\scripts\Temp\"

:: Этот блок записывает в файл результат выпонлнения
>result.txt echo Скопировано %date%_%time:~,8% в "%d2%":

:: Этот блок перемещает log файл
move "%d1%" "%d2%%date%.txt"

pause& exit
