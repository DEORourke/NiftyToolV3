@echo off
ECHO Pinging %1
ECHO Saving the log to %3
ECHO.
ECHO Press ctrl+C or close this window to end the test.
::	The below command continuously pings the address passed to the script in the first variable with a packet size specified by the second variable and for each ping prepends a timestamp to the output before writing to a file at the path passed in the third variable.
ping -t %1 -l %2 |cmd /q /v /c "(pause&pause)>nul & for /l %%a in () do (set /p "data=" && echo(!time! !data!)&ping -n 2 localhost>nul" > %3
::	Currently not able to find a way to run this command from a function in the main JS file. Just too many layers of quotes. This will hopefully be corrected in the next update.
:EOF