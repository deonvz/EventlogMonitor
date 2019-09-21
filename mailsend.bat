@Rem Scripted by Deon van Zyl
@echo ---Start--  >> sendmail.txt

@Rem -----Set Enviroment Vars---

@SET required_space=Threshold
@SET source_email=websystems@someplace.com   
@SET destination_email=Deonvz@someplace.com   
@SET mailserver=10.0.0.1
@SET domain=someplace.com
@Rem ----------------------------

@date /t >> sendmail.txt
@time /t >> sendmail.txt
MAILSEND -d %domain% -smtp %mailserver% -t %destination_email% -f %source_email% -sub "Eventlog Errors on %computername%" -v < msg.txt >> sendmail.txt
@echo ---ENd--  >> sendmail.txt
@echo -----------  >> sendmail.txt
pause
exit