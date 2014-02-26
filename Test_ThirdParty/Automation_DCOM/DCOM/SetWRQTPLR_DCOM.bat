SET logfile="C:\Temp\RunDCOM.log"

@echo off
echo Working on default DCOM permissions  >%logfile%
echo>>%logfile%
echo>>%logfile%
echo    DCOM Access Limit Permissions   >>%logfile%
echo  >>%logfile%
dcomperm -ma set Administrators permit level:r,l  >>%logfile%
dcomperm -ma set Administrator permit level:r,l  >>%logfile%
dcomperm -ma set "Authenticated Users" permit level:r,l  >>%logfile%
dcomperm -ma set "Anonymous Logon" permit level:r,l  >>%logfile%
dcomperm -ma set Everyone permit level:r,l   >>%logfile%
dcomperm -ma set Interactive permit level:r,l  >>%logfile%
dcomperm -ma set Network permit level:r,l    >>%logfile%
dcomperm -ma set System permit level:r,l    >>%logfile%
echo
echo    DCOM Access Permissions  >>%logfile%
echo
dcomperm -da set Administrators permit level:r,l  >>%logfile%
dcomperm -da set Administrator permit level:r,l   >>%logfile%
dcomperm -da set "Authenticated Users" permit level:r,l  >>%logfile%
dcomperm -da set "Anonymous Logon" permit level:r,l   >>%logfile%
dcomperm -da set Everyone permit level:r,l   >>%logfile%
dcomperm -da set Interactive permit level:r,l   >>%logfile%
dcomperm -da set Network permit level:r,l  >>%logfile%
dcomperm -da set System permit level:r,l   >>%logfile%
echo  >>%logfile%
echo     DCOM Launch Permissions   >>%logfile%
echo   >>%logfile%
dcomperm -ml set Administrators permit level:rl,ll,la,ra   >>%logfile%
dcomperm -ml set Administrator permit level:rl,ll,la,ra  >>%logfile%
dcomperm -ml set "Authenticated Users" permit level:rl,ll,la,ra   >>%logfile%
dcomperm -ml set "Anonymous Logon" permit level:rl,ll,la,ra   >>%logfile%
dcomperm -ml set Everyone permit level:rl,ll,la,ra    >>%logfile%
dcomperm -ml set Interactive permit level:rl,ll,la,ra   >>%logfile%
dcomperm -ml set Network permit level:rl,ll,la,ra    >>%logfile%
dcomperm -ml set System permit level:rl,ll,la,ra   >>%logfile%
echo   >>%logfile%
echo     DCOM Launch Permissions   >>%logfile%
echo    >>%logfile%
dcomperm -dl set Administrators permit level:rl,ll,la,ra   >>%logfile%
dcomperm -dl set Administrator permit level:rl,ll,la,ra   >>%logfile%
dcomperm -dl set "Authenticated Users" permit level:rl,ll,la,ra   >>%logfile%
dcomperm -dl set "Anonymous Logon" permit level:rl,ll,la,ra  >>%logfile%
dcomperm -dl set Everyone permit level:rl,ll,la,ra   >>%logfile%
dcomperm -dl set Interactive permit level:rl,ll,la,ra   >>%logfile%
dcomperm -dl set Network permit level:rl,ll,la,ra   >>%logfile%
dcomperm -dl set System permit level:rl,ll,la,ra   >>%logfile%
echo   >>%logfile%
echo  >>%logfile%
echo Working on AQT Remote Agent permissions  >>%logfile%
echo  >>%logfile%
dcomperm  -runas {25E8BB22-5C86-11D4-90DA-00104B3E51B1} "Interactive User"   >>%logfile%
dcomperm  -al {25E8BB22-5C86-11D4-90DA-00104B3E51B1} Default  >>%logfile%
dcomperm  -aa {25E8BB22-5C86-11D4-90DA-00104B3E51B1} Default   >>%logfile%
echo  >>%logfile%
echo >>%logfile%
echo Working on QTP Automation Server   >>%logfile%
echo >>%logfile%
dcomperm  -runas {A67EB23A-1B8F-487D-8E38-A6A3DD150F0B} "Interactive User"  >>%logfile%
dcomperm  -al {A67EB23A-1B8F-487D-8E38-A6A3DD150F0B} Default  >>%logfile%
dcomperm  -aa {A67EB23A-1B8F-487D-8E38-A6A3DD150F0B} Default  >>%logfile%
echo  >>%logfile%
echo  >>%logfile%
echo Working on Tulip Remote Server >>%logfile%
echo  >>%logfile%
dcomperm  -runas {70396405-BE62-11D2-8F0B-00104B3E51B1} "Interactive User"  >>%logfile%
dcomperm  -al {70396405-BE62-11D2-8F0B-00104B3E51B1} Default  >>%logfile%
dcomperm  -aa {70396405-BE62-11D2-8F0B-00104B3E51B1} Default  >>%logfile%
echo  >>%logfile%
echo  >>%logfile%
echo Working on WinRunner Remote Agent  >>%logfile%
echo  >>%logfile%
dcomperm  -runas {0B171F02-F204-11D0-9398-0080C837F11F} "Interactive User"  >>%logfile%
dcomperm  -al {0B171F02-F204-11D0-9398-0080C837F11F} Default   >>%logfile%
dcomperm  -aa {0B171F02-F204-11D0-9398-0080C837F11F} Default   >>%logfile%
echo  >>%logfile%
echo >>%logfile%
echo Working on WinRunner Document object  >>%logfile%
echo  >>%logfile%
dcomperm  -runas {CD70EDCE-7777-11D2-9509-0080C82DD192} "Interactive User"  >>%logfile%
dcomperm  -al {CD70EDCE-7777-11D2-9509-0080C82DD192} Default  >>%logfile%
dcomperm  -aa {CD70EDCE-7777-11D2-9509-0080C82DD192} Default  >>%logfile%
echo  >>%logfile%
echo  >>%logfile%
echo Working on Vapi-XP object  >>%logfile%
echo  >>%logfile%
dcomperm  -runas {FCB69899-EC52-4A7A-86DB-3655E9FDBA58} "Interactive User"  >>%logfile%
dcomperm  -al {FCB69899-EC52-4A7A-86DB-3655E9FDBA58} Default   >>%logfile%
dcomperm  -aa {FCB69899-EC52-4A7A-86DB-3655E9FDBA58} Default  >>%logfile%
echo   >>%logfile%
echo   >>%logfile%
echo Working on Business Process Testing object   >>%logfile%
echo   >>%logfile%
dcomperm  -runas {6108A56C-6239-41F6-8C0F-94D9CE0D4B61} "Interactive User"   >>%logfile%
dcomperm  -al {6108A56C-6239-41F6-8C0F-94D9CE0D4B61} Default   >>%logfile%
dcomperm  -aa {6108A56C-6239-41F6-8C0F-94D9CE0D4B61} Default   >>%logfile%
echo  >>%logfile%
echo   >>%logfile%
echo Working on System Test Remote Agent  >>%logfile%
echo  >>%logfile%
dcomperm  -runas {1B78CAE4-A6A8-11D5-9D7A-000102E1A2A2} "Interactive User"  >>%logfile%
dcomperm  -al {1B78CAE4-A6A8-11D5-9D7A-000102E1A2A2} Default  >>%logfile%
dcomperm  -aa {1B78CAE4-A6A8-11D5-9D7A-000102E1A2A2} Default  >>%logfile%
echo  >>%logfile%
echo   >>%logfile%
echo Working on LoadRunner Specific Settings   >>%logfile%
echo   >>%logfile%
dcomperm  -runas {E933439A-81A1-11D4-8EEE-0050DA6171E8} "Interactive User"  >>%logfile%
dcomperm  -al {E933439A-81A1-11D4-8EEE-0050DA6171E8} Default   >>%logfile%
dcomperm  -aa {E933439A-81A1-11D4-8EEE-0050DA6171E8} Default  >>%logfile%



