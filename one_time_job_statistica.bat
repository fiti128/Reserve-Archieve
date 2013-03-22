REM Программа резервного копирования АС "Статистический учет капитальных вложений"
REM Описание: Программа архивирует и копирует указанные папки, файлы с удаленной машины. 
REM Перед копированием посылает сообщение пользователю о закрытии программы. Через указанное время
REM закрывает ее. По завершению работы удаляет архивы старше 6 дней.
REM В случае ошибки назначает одноразовое задание на выполение программы еще раз на следующий день.
REM Весь процесс логируется в файл. Учитываются високосные года.
REM ----------------------------------------------------------------------------
REM Устанавливаем кодировку для возможности копирования файлов с кириллицей в названии.
chcp 1251
@echo off
cls
REM ----------------------------------------------------------------------------
REM Задаем переменные
set ip=10.1.3.167
set /A waiting_time=30
set msg=Через %waiting_time% секунд начнется резервное копирование! Программа будет принудительно закрыта
 REM externalUser и externalPassword - относятся к компьютеру с которого производиться копирование.
set externalUser=10.1.3.167\pasha
set externalPassword=Gfif1234
 REM local - относятся к компьютеру, с которого запускается бат файл
set local_user=IRCM-TEST\Administrator 
set local_password=Aa6677
set program_to_close=CityInfo.exe
  REM Задаем файлы и папки для копирования Z - это диск С, N - это диск D
set _in=N:\vokssta N:\noksstam N:\noksstan Z:\copsta Z:\copstaM Z:\copstaN Z:\stat Z:\statmr Z:\stands
set _7-zip_path=c:\Program Files\7-zip
set archive_path=c:\archive\statistica
set error_job_name=one_time_job_statistica
set log_file=log_statistica.txt

REM ----------------------------------------------------------------------------
REM Дальше не трогать, если не знаете че делаете
REM Подчищаем за виндой
schtasks /delete /TN %error_job_name% /f
REM Проверяем ай-пи, логин и пароль, одновременно записывая во временный кэш логин и пароль
tasklist.exe /S %ip% /U %externalUser% /P %externalPassword%
if errorlevel 1 goto failConnection
REM Включаем службу Windows messenger
sc \\%ip% config messenger start= auto
sc \\%ip% start messenger
echo Даем время службе включиться.
ping -n 15 localhost
REM Посылаем сообщение пользователю
net send %ip% Через %waiting_time% секунд начнется резервное копирование! Программа будет принудительно закрыта будет закрыта
if errorlevel 1 goto failmsg
REM Отключаем службу Windows messenger
sc \\%ip% stop messenger
sc \\%ip% config messenger start= disabled
echo Ждем %waiting_time%  секунд. 
ping -n %waiting_time% localhost
REM Закрываем программу
taskkill.exe /S %ip% /U %externalUser% /P %externalPassword% /IM %program_to_close% /F

REM СОБСТВЕННО, АРХИВИРОВАНИЕ, КОПИРОВАНИЕ ФАЙЛОВ В СОЗДАННУЮ ПАПКУ, С ПАРАМЕТРАМИ: БЕЗ ВЫВОДА ЗАПРОСА О ПЕРЕЗАПИСИ, С ПОДКАТАЛОГАМИ, РЕЖИМ ПЕРЕЗАПУСКА, ИГНОР ОШИБОК, + СИСТЕМНЫЕ И СКРЫТЫЕ ФАЙЛЫ
 REM Форматируем дату
 set nd=%date:~6,4%-%date:~3,2%-%date:~0,2%
 REM Подключаем сетевые диски
 net use Z: \\%ip%\C$ /USER:%externalUser% %externalPassword%
 if errorlevel 1 goto failConnection
 net use N: \\%ip%\D$ /USER:%externalUser% %externalPassword%
 if errorlevel 1 goto failConnection
 set _out=Z:\%nd%
 REM Архивируем
 echo Запущенна архивация... Пожалуйста подождите!
 for %%i in (%_in%) do "%_7-zip_path%\7z" a "%_out%\%%~ni.zip" "%%i"
 echo Архивация завершенна!
 md %archive_path%\%nd%
 REM Копируем
 echo Началось копирование данных... Пожалуйста подождите!
 xcopy %_out% %archive_path%\%date:~6,4%-%date:~3,2%-%date:~0,2% /d /y /e /c /h
 if errorlevel 1 goto failCopy
 echo Копирование прошло успешно!
 REM Подчищаем и отключаем сетевые диски
 rd Z:\%nd% /s /q /s /q
 net use Z: /d
 net use N: /d

goto done 


 :failCopy
echo ОШИБКА --- %date% %time% : Не удалось скопировать файлы с %ip%. Компьютер стал недоступен или есть программы использующие данные файлы >> %log_file%
echo Не удалось скопировать файлы с %ip%. Компьютер стал недоступен или есть программы использующие данные файлы
goto schtasks
 :failConnection
echo ОШИБКА --- %date% %time% : Не удалось подключится к %ip%. Возможно неверное имя пользователя или пароль >> %log_file%
echo Не удалось подключится к %ip%. Возможно неверное имя пользователя или пароль 
goto schtasks
 :failmsg
echo ОШИБКА --- %date% %time% : Не удалось доставить сообщение о закрытии программы >> %log_file%
echo Не удалось доставить сообщение о закрытии программы
goto schtasks
 :schtasks
set This_Date=%date%
 REM Парсим дату в инты дня, месяца и года для арифметических вычислений
 IF %This_DATE:~0,1%==0 (
 SET /A This_DAY=%This_DATE:~1,1%) else (
 SET /A This_DAY=%This_DATE:~0,2%)
 IF %This_DATE:~3,1%==0 (
 SET /A This_MONTH=%This_DATE:~4,1%) else (
 SET /A This_MONTH=%This_DATE:~3,2%)
 SET /A This_YEAR=%This_DATE:~-4,4%
 
 REM В зависимости от високосного года определяем dif(Days in February)
 set /A feb=This_YEAR%%4
 set /A dif=28
 if %feb%==0 set /A dif=29
 set /A feb=This_YEAR%%100
 if %feb%==0 set /A dif=28
 set /A feb=This_YEAR%%400
 if %feb%==0 set /A dif=29
 
 REM Задаем переменную инту DIM(Days in Month), которая показывает сколько дней в этом месяце
 if %This_MONTH%==1 Set /A DIM=31
 if %This_MONTH%==2 Set /A DIM=%dif%
 if %This_MONTH%==3 Set /A DIM=31
 if %This_MONTH%==4 Set /A DIM=30
 if %This_MONTH%==5 Set /A DIM=31
 if %This_MONTH%==6 Set /A DIM=30
 if %This_MONTH%==7 Set /A DIM=31
 if %This_MONTH%==8 Set /A DIM=31
 if %This_MONTH%==9 Set /A DIM=30
 if %This_MONTH%==10 Set /A DIM=31
 if %This_MONTH%==11 Set /A DIM=30
 if %This_MONTH%==12 Set /A DIM=31 
 REM Собственно алгоритм определения завтрашнего дня
 SET /A checking_day=%DIM%-%This_DAY%-1
 SET /A tomorrow_year = %This_YEAR%
 if %checking_day% GEQ 0 SET /A tomorrow_day = %This_DAY%+1
 if %checking_day% GEQ 0 SET /A tomorrow_month = %This_MONTH%
 if %checking_day% GEQ 0 SET /A tomorrow_year = %This_YEAR%
 if %checking_day% LSS 0 SET /A tomorrow_day = 0 - %checking_day%
 if %checking_day% LSS 0 SET /A tomorrow_month = %This_MONTH% +1
 if %tomorrow_month% == 13 SET /A tomorrow_year = %This_YEAR%+1
 if %tomorrow_month% == 13 SET /A tomorrow_month = 1
 if %tomorrow_day% LSS 10 SET tomorrow_DAY=0%tomorrow_day%
 if %tomorrow_month% LSS 10 SET tomorrow_MONTH=0%tomorrow_month%
set tomorrow_date=%tomorrow_day%/%tomorrow_month%/%tomorrow_year%
set path_to_file=%0
REM Создаем новый таск на следующий день
schtasks /create /SC ONCE /SD %tomorrow_date% /ST 13:30:00 /TR %path_to_file% /TN one_time_job_statistica /RU %local_user% /RP %local_password%
echo -- Работа программы завершилась со сбоем --- Созданна повторная задача на %tomorrow_date% >> %log_file%
 goto exit
 :success
echo УСПЕХ --- %date% %time% : Задача резервного копирования АС "Статистический учет капитальных вложений" выполненна УСПЕШНО! >> %log_file%
echo УСПЕХ --- %date% %time% : Архивы доступны по адресу %archive_path%\%nd% >> %log_file%
echo Задача по копированию файлов выполненна УСПЕШНО!
goto exit

REM Ошибок нет, значит можно удалить старый бэкап.
 :done
REM СЧИТАЕМ И УДАЛЯЕМ КАТАЛОГ С БЭКАПОМ, СОЗДАННЫЙ НЕДЕЛЮ НАЗАД
 set /a counter=5
 :loop 
set /a counter+=1

set T_Date=%DATE%
  echo %DATE%
 IF %T_DATE:~0,1%==0 (
 SET /A T_DAY=%T_DATE:~1,1%) else (
 SET /A T_DAY=%T_DATE:~0,2%)

 IF %T_DATE:~3,1%==0 (
 SET /A T_MONTH=%T_DATE:~4,1%) else (
 SET /A T_MONTH=%T_DATE:~3,2%)
 SET /A T_YEAR=%T_DATE:~6,4%

REM ТУТ МЕНЯЕМ ЧИСЛО, ЗА СКОЛЬКО ДНЕЙ НАЗАД БУДЕМ ПОКАЗЫВАТЬ ДАТУ (до 28 дней, т.к. скрипт, при "отрицательном" дне всегда делает - 1 месяц)
 SET /A T_DAY=%T_DAY%-%counter%

 IF %T_DAY% LEQ 0 SET /A T_MONTH=%T_MONTH%-1
 IF %T_MONTH%== 0 SET /A T_YEAR=%T_YEAR%-1
 IF %T_MONTH%== 0 SET /A T_MONTH=12
 if %T_MONTH%==1 Set /A DIM=31
 if %T_MONTH%==2 Set /A DIM=28
 if %T_MONTH%==3 Set /A DIM=31
 if %T_MONTH%==4 Set /A DIM=30
 if %T_MONTH%==5 Set /A DIM=31
 if %T_MONTH%==6 Set /A DIM=30
 if %T_MONTH%==7 Set /A DIM=31
 if %T_MONTH%==8 Set /A DIM=31
 if %T_MONTH%==9 Set /A DIM=30
 if %T_MONTH%==10 Set /A DIM=31
 if %T_MONTH%==11 Set /A DIM=30
 if %T_MONTH%==12 Set /A DIM=31
 IF %T_DAY% LEQ 0 SET /A T_DAY=%T_DAY%+%DIM%
 IF %T_DAY% LSS 10 SET T_DAY=0%T_DAY%
 IF %T_MONTH% LSS 10 SET T_MONTH=0%T_MONTH%

REM Удалям, собственно...
 rd %archive_path%\%T_YEAR%-%T_MONTH%-%T_DAY% /s /q
 if not %counter%==22 goto loop

goto success

goto exit

 :exit


