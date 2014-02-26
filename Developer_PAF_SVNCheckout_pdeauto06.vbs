'run the pstools to check out the latest code in pdeauto06 host ,so you can easily check out the code remotely
'the default workspace in pdeauto06 is 

strShareFolder="\\pdeauto06.fc.hp.com\PAF_Automation"
strSVNServer="https://pdeauto06.fc.hp.com/svn/PAF_Automation_QTP"    'the SVN Server URL for QTP repository
strSVNUser="Alter"      ' the SVN log user
strSVNPassword="Alter"  ' the SVN log password
strSVNRevision="head"     'can be any revision number you need to test
strLocalLocation=" C:\PAF_Automation"   ' the local copy workspace from SVN server

strPsexec="\Test_ThirdParty\PSTools\psExec.exe"

Set ws=CreateObject("wscript.shell")
'Set fso=CreateObject("scripting.filesystemobject")
ws.Popup "Now is running to check out code from SVN to pdeauto06 Server,This maybe meeds about 2 minutes.......",6,"SVN Code Check Out Tool",0+64
'Set shareFolder=fso.GetFolder(strShareFolder)
'For Each subfolder In shareFolder.SubFolders
   ' If InStr(subfolder.Name,".svn")<=0 Then 
       ' subfolder.Delete True 
   ' End If 
'Next 
'Set shareFolder=Nothing
'Set fso=nothing
'get the pstools directory we need to use blow 
strPsPath=ws.CurrentDirectory+strPsexec

' careful with the parameter i ,which means the session id in the remote host
strPscommand=strPsPath&" \\pdeauto06.fc.hp.com -u pdeauto06\qatest -p L0ngh)rn -i 0 -d -w "&strLocalLocation&" cmd.exe "
strSVNDeleteCommand="""/k svn  delete *.*"""
strDeletesvn=" /k FOR /R . %f IN (.svn) DO RD /s /q %f"
strCheckoutCommand="""/k svn checkout "&strSVNServer&" . -r "&strSVNRevision&"  --force  --username "&strSVNUser&" --password "&strSVNPassword&""""
strCommand=strPscommand&strCheckoutCommand
' first delete the directory ,clean up
strComamndDelete=strPscommand&strSVNDeleteCommand
strComamndDelete2=strPscommand&strDeletesvn

ws.Popup  "will Complete Check out all the latest codes into pdeauto06,DONOT close the command line windows!!!!",8,"SVN Code Check Out Tool",0+64

'use the svn delete to delete the SVN copy workspace
ws.Run strComamndDelete,1,True
'use the for loop to delete the .svn folder
ws.Run strComamndDelete2,1,true
ws.Run strCommand,1,true



ws.Popup  "Congratulation ,All codes checked out in pdeauto06 !!! you can enjoy the test now !!!",3,"SVN Code Check Out Tool",0+64

Set ws=Nothing
