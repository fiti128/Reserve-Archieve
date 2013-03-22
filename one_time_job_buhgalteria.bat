REM ��������� ���������� ����������� �� "������������� ���� ����������� ��������"
REM ��������: ��������� ���������� � �������� ��������� �����, ����� � ��������� ������. 
REM ����� ������������ �������� ��������� ������������ � �������� ���������. ����� ��������� �����
REM ��������� ��. �� ���������� ������ ������� ������ ������ 6 ����.
REM � ������ ������ ��������� ����������� ������� �� ��������� ��������� ��� ��� �� ��������� ����.
REM ���� ������� ���������� � ����. ����������� ���������� ����.
REM ----------------------------------------------------------------------------
REM ������������� ��������� ��� ����������� ����������� ������ � ���������� � ��������.
chcp 1251
@echo off
cls
REM -----------------------------------------------------------------------------
REM ������ ����������
set ip=10.1.3.167
set /A waiting_time=30
set msg=����� %waiting_time% ������ �������� ��������� �����������! ��������� ����� ������������� �������
 REM externalUser � externalPassword - ��������� � ���������� � �������� ������������� �����������.
set externalUser=10.1.3.167\pasha
set externalPassword=Gfif1234
 REM local - ��������� � ����������, � �������� ����������� ��� ����
set program_to_close=CityInfo.exe
  REM ������ ����� � ����� ��� ����������� Z - ��� ���� �, N - ��� ���� D
set _in=N:\foxpro N:\mt N:\mtdbf N:\mtnoks
set _7-zip_path=c:\Program Files\7-zip
set archive_path=c:\archive\buhgalteria
set log_file=log_buhgaleria.txt

REM ----------------------------------------------------------------------------
REM ������ �� �������, ���� �� ������ �� �������
REM ��������� ��-��, ����� � ������, ������������ ��������� �� ��������� ��� ����� � ������
tasklist.exe /S %ip% /U %externalUser% /P %externalPassword%
if errorlevel 1 goto failConnection
REM �������� ������ Windows messenger
sc \\%ip% config messenger start= auto
sc \\%ip% start messenger
echo ���� ����� ������ ����������.
ping -n 15 localhost
REM �������� ��������� ������������
net send %ip% ����� %waiting_time% ������ �������� ��������� �����������! ��������� ����� ������������� ������� ����� �������
if errorlevel 1 goto failmsg
REM ��������� ������ Windows messenger
sc \\%ip% stop messenger
sc \\%ip% config messenger start= disabled
echo ���� %waiting_time%  ������. 
ping -n %waiting_time% localhost
REM ��������� ���������
taskkill.exe /S %ip% /U %externalUser% /P %externalPassword% /IM %program_to_close% /F

REM ����������, �������������, ����������� ������ � ��������� �����, � �����������: ��� ������ ������� � ����������, � �������������, ����� �����������, ����� ������, + ��������� � ������� �����
 REM ����������� ����
 set nd=%date:~6,4%-%date:~3,2%-%date:~0,2%
 REM ���������� ������� �����
 net use Z: \\%ip%\C$ /USER:%externalUser% %externalPassword%
 if errorlevel 1 goto failConnection
 net use N: \\%ip%\D$ /USER:%externalUser% %externalPassword%
 if errorlevel 1 goto failConnection
 set _out=Z:\%nd%
 REM ����������
 echo ��������� ���������... ���������� ���������!
 for %%i in (%_in%) do "%_7-zip_path%\7z" a "%_out%\%%~ni.zip" "%%i"
 echo ��������� ����������!
 md %archive_path%\%nd%
 REM ��������
 echo �������� ����������� ������... ���������� ���������!
 xcopy %_out% %archive_path%\%date:~6,4%-%date:~3,2%-%date:~0,2% /d /y /e /c /h
 if errorlevel 1 goto failCopy
 echo ����������� ������ �������!
 REM ��������� � ��������� ������� �����
 rd Z:\%nd% /s /q /s /q
 net use Z: /d
 net use N: /d

goto done 

 :failCopy
echo ������ --- %date% %time% : �� ������� ����������� ����� � %ip%. ��������� ���� ���������� ��� ���� ��������� ������������ ������ ����� >> %log_file%
echo �� ������� ����������� ����� � %ip%. ��������� ���� ���������� ��� ���� ��������� ������������ ������ �����
goto exit
 :failConnection
echo ������ --- %date% %time% : �� ������� ����������� � %ip%. �������� �������� ��� ������������ ��� ������ >> %log_file%
echo �� ������� ����������� � %ip%. �������� �������� ��� ������������ ��� ������ 
goto exit
 :failmsg
echo ������ --- %date% %time% : �� ������� ��������� ��������� � �������� ��������� >> %log_file%
echo �� ������� ��������� ��������� � �������� ���������
goto exit
 
 :success
echo ����� --- %date% %time% : ������ ���������� ����������� �� "������������� ���� ����������� �������� �������! >> %log_file%
echo ����� --- %date% %time% : ������ �������� �� ������ %archive_path%\%nd% >> %log_file%
echo ������ �� ����������� ������ ���������� �������!
goto exit

REM ������ ���, ������ ����� ������� ������ �����.
 :done
REM ������� � ������� ������� � �������, ��������� ������ �����
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

REM ��� ������ �����, �� ������� ���� ����� ����� ���������� ���� (�� 28 ����, �.�. ������, ��� "�������������" ��� ������ ������ - 1 �����)
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

REM ������, ����������...
 rd %archive_path%\%T_YEAR%-%T_MONTH%-%T_DAY% /s /q
 if not %counter%==22 goto loop

goto success

goto exit

 :exit


