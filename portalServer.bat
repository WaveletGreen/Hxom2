@echo off
rem 
rem TPR is short for TC_PORTAL_ROOT to reduce command line length
title Teamcenter Rich Client
set TPR=C:\Siemens\Teamcenter11\portal
if not defined FMS_HOME set FMS_HOME=C:\TC\Siemens\Teamcenter11\tccs
set PATH=C:\TC\Siemens\Teamcenter11\tccs\bin;C:\TC\Siemens\Teamcenter11\tccs\lib;%TPR%;%PATH%
set JAVA_HOME=%C:\TC\Java\jre7%
set JRE_HOME=%C:\TC\Java\jre7%
start "TAO ImR" /min cmd /c "C:\TC\Siemens\Teamcenter11\iiopservers\start_imr.bat"