' this file is used for developer to develop this framework code easily
Call qtp_loader()
Sub qtp_loader()
       ' On Error Resume Next 
        
		Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
		Dim qtTestRecovery 'As QuickTest.Recovery ' Declare a Recovery object variable
		Dim intIndex ' Declare an index variable
		
		Dim arrayaddins
		functiondir=Trim(InputBox("Input the local directory:","QTP Settings"))
		' Open UFT and prepare objects variables
		Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
	
	
		'add ins
		arrayaddins=Array("Web","ActiveX")
		isloaded=qtApp.SetActiveAddins(arrayaddins,errordesc)
		If Not isloaded Then
		   MsgBox "QTP cannot load the add ins:"+errordesc
		   Exit sub
		End If 
	
			
		If Not  qtApp.Launched Then 
		       qtApp.Launch ' Start UFT
		End If 
		qtApp.Visible = True ' Make the UFT application visible
		qtApp.New ' Open a new test
		
	  
	    If qtApp.Test.IsNew Then
             qtApp.Test.SaveAs "C:\Temp\TempTest"
        end if
		'configuration the run
		qtApp.ShowPaneScreen "ActiveScreen", false ' Display the Active Screen pane
		qtApp.ShowPaneScreen "DataTable", true ' Hide the Data Table Explorer pane
		qtApp.ShowPaneScreen "DebugViewer", True ' Display the Debug Watch pane
		qtApp.WindowState = "Maximized" ' Maximize the UFT window

	
		' configuration the settings
		Set qtSettings=qtApp.Test.Settings
		
		addinsloaded=qtApp.Test.ValidateAddins()
		If Not addinsloaded Then
		   MsgBox "Validate the add in loading correctly:"&addinsloaded
		   Exit sub
		End If 
		'set run settings
		Set qtRun=qtSettings.Run
		qtRun.IterationMode = "rngIterations"
		qtRun.ObjectSyncTimeOut=0
		qtRun.DisableSmartIdentification=true
		
		Set qtWeb=qtSettings.Web
		qtWeb.BrowserNavigationTimeout=40
		
		'launcher
		qtSettings.Launchers("Web").Active=false
		qtSettings.Launchers("Windows Applications").Active = false
		qtSettings.Launchers("Windows Applications").Applications.RemoveAll
	    qtSettings.Launchers("Windows Applications").RecordOnQTDescendants = True
		qtSettings.Launchers("Windows Applications").RecordOnExplorerDescendants = False
		qtSettings.Launchers("Windows Applications").RecordOnSpecifiedApplications = True
		
		'resources settings
		Set qtLibraries = qtSettings.Resources.Libraries 
		
		If qtLibraries.Count>0 Then 
		    qtLibraries.RemoveAll
		End If  
		
		'set function loaded
		Set objFunctions=CreateObject("scripting.filesystemobject")
	
		
		strUtilitydir=functiondir+"\Test_Libraries"
		strPageActiondir=functiondir+"\Test_Page_Action"
		strRecoverydir=functiondir+"\Test_Recovery_Scenarios"
		strTestsuitedir=functiondir+"\Test_Suites"
		
		'utility directory
		If objFunctions.FolderExists(strUtilitydir) Then 
			Set utilitydir=objFunctions.GetFolder(strUtilitydir)		
			For Each utilityfile In utilitydir.Files
			     If Right(utilityfile.Name,3)="qfl" Or Right(utilityfile.Name,3)="vbs" then
			        qtLibraries.Add utilityfile.Path
			     End If 
			Next 
	     End If 
		'page action directory
		If objFunctions.FolderExists(strPageActiondir) Then 
			Set pagedir=objFunctions.GetFolder(strPageActiondir)		
			For Each pagefile In pagedir.Files
			     If Right(pagefile.Name,3)="qfl"  then
			        qtLibraries.Add pagefile.Path
			     End If 
			Next 
	     End If 
	     
	     
	     'recovery directory
	    ' Enable the recovery mechanism (with default, on errors, setting)
		
	     If objFunctions.FolderExists(strRecoverydir) Then 
			Set recoverydir=objFunctions.GetFolder(strRecoverydir)		
			For Each recoveryfile In recoverydir.Files
			     If Right(recoveryfile.Name,3)="qfl"  then
			        qtLibraries.Add recoveryfile.Path
			     End If 
			Next 
	     End If 
	     'test suite files
	      'recovery directory
	     If objFunctions.FolderExists(strTestsuitedir) Then 
			Set testsuitedir=objFunctions.GetFolder(strTestsuitedir)		
			For Each testsuitefile In testsuitedir.Files
			     If Right(testsuitefile.Name,3)="qfl"  then
			        qtLibraries.Add testsuitefile.Path
			     End If 
			Next 
	     End If 
		
		'recovery settings	
		
		Set qtTestRecovery = qtSettings.Recovery ' Return the Recovery object for the current test
		qtTestRecovery.Enabled = True
		If qtTestRecovery.Count > 0 Then ' If there are any default scenarios specified for the test
		    qtTestRecovery.RemoveAll ' Remove them
		End If
		
		' Add recovery scenarios
		If objFunctions.FolderExists(strRecoverydir) Then 
			Set recoveryfiledir=objFunctions.GetFolder(strRecoverydir)		
			For Each recoverysfile In recoveryfiledir.Files
			     If Right(recoverysfile.Name,3)="qrs"  then
			        qtTestRecovery.Add recoverysfile.Path,"Recovery_RunAllRecoveryScenarios"
			     End If 
			Next 
	     End If 
		
		' Enable all scenarios
		For intIndex = 1 To qtTestRecovery.Count ' Iterate the scenarios
		    qtTestRecovery.Item(intIndex).Enabled = True ' Enable each Recovery Scenario (Note: the 'Item' property is default and can be omitted)
		Next
		
		
		'Ensure that the recovery mechanism is set to be activated only after errors
		qtTestRecovery.SetActivationMode "OnError"
		'OnError is the default, the other option is "OnEveryStep".
		
		
		Set optionrun=qtApp.Options.Run
		optionrun.ViewResults=false
		optionrun.RunMode="Fast"
		
	
		
		Set qtWebOptions =qtApp.Options.Web
		
		qtWebOptions.EnableBrowserResize = False ' Set to open the browser to its default size
		qtWebOptions.RunUsingSourceIndex = True ' Set to use the source index property (for better performance)
		qtWebOptions.UseAutoXPathIdentifiers = True ' Instruct UFT to learn and use an XPath property during a run session
		qtWebOptions.RunOnlyClick = True ' Set to run click events as MouseDown, MouseUp and Click
		qtWebOptions.BrowserCleanup = false ' Set to close all open browsers when test/iteration finishes
		qtWebOptions.RecordByWinMouseEvents = "OnClick OnMouseDown" ' Indicate for which events to use standard Windows events
		qtWebOptions.RecordAllNavigations = True ' Set to record navigation each time the URL changes

        qtApp.Options.DisplayKeywordView=False
        
        
		
		Set qtApp = Nothing ' Release the Application object
		Set objFunctions=nothing
		Set qtTestRecovery = Nothing ' Release the Recovery object
End sub
