; The following example demonstrates how to automate FTP uploading using the operating 
; system's built-in FTP command. This script has been tested on Windows XP and 98se.

FTPCommandFile = %A_ScriptDir%\FTPCommands.txt
FTPLogFile = %A_ScriptDir%\FTPLog.txt
FileDelete %FTPCommandFile%  ; In case previous run was terminated prematurely.

FileAppend,
(
open ftp.spiketrade.com
spikeworld@spiketrade.com
world007
binary
put C:\Users\Vla\Pictures\NH-NL\Imagens\AUS-D-VladimirDietrichkeitRightsReserved.gif
ls -l
quit
), %FTPCommandFile%

RunWait %comspec% /c ftp.exe -s:"%FTPCommandFile%" >"%FTPLogFile%"
FileDelete %FTPCommandFile%  ; Delete for security reasons.
Run %FTPLogFile%  ; Display the log for review.