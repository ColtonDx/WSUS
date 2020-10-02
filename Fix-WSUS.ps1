net stop CryptSvc 
net stop BITS 
net stop dosvc
net stop wuauserv
ren %windir%\SoftwareDistribution SoftwareDistribution.old
reg Delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientId /f 
reg Delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientIDValidation /f
net start wuauserv
net Start BITS 
net start CryptSvc 
wuauclt /resetauthorization /detectnow
wuauclt /detectnow /reportnow
pause