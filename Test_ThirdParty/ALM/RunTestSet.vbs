'VBScript Document
Option Explicit
'WScript.Interactive=false

'QC Paramters
Dim Server, UserName, Password, QCDomain, QCProject, QCTestSetPath, QCTestSetName
Dim strExecutionLogFile

'Server = "http://qc1d.atlanta.hp.com/qcbin/"
Server="http://qc1d.atlanta.hp.com/qcbin/"
UserName = "chang-yuan.hu_hp.com"
Password = "gu.chan-1026"
QCDomain = "TS_DEV"
QCProject = "Support_Automation"
QCTestSetPath = "Root\techassessments\PAF_Automation"
'QCTestSetName = "PAF_E2E_SmokeTesting_pdeauto06.fc.hp.com"
QCTestSetName="PAF_SmokeTesting_pdeauto06.fc.hp.com"

strExecutionLogFile="C:\alm.log"

''' <summary>
''' Loads an instance of QCRunTestSet and initiates the scheduling controller
''' </summary>
''' <param name="Server">QC Server</param>
''' <param name="UserName"></param>
''' <param name="Password"></param>
''' <param name="Domain">Domain name</param>
''' <param name="Project">Project name</param>
''' <param name="QCTestSetPath">Path to the TestSet (Root\FolderName\SubFolderName)</param>
''' <param name="QCTestSetName">Name of the test set</param>
''' <remarks></remarks>
Public Sub RunTestSet(Server, UserName, Password, QCDomain, QCProject, QCTestSetPath, QCTestSetName)
	Dim QCTestSetExec, sErr, arrArgs, ix, arg, bExit
	
	sErr = "Unable to execute RunTestSet. Please provide the "
	arrArgs = Array("Server", "UserName", "Password", "QCDomain", "QCProject", "QCTestSetPath", "QCTestSetName")
	bExit = False
	
	For ix = LBound(arrArgs) To UBound(arrArgs)
		Execute "arg = " & arrArgs(ix)

		If arg = "" Then
			logs sErr & arrArgs(ix) & "."&"Error!"
			
			bExit = True
		End If
	Next
	
	If bExit Then Exit Sub
	
	Set QCTestSetExec = New QCRunTestSet
	logs "Test is begin to run ...."
	With QCTestSetExec
		.Server = Server
		.UserName = UserName
		.Password = Password
		.QCDomain = QCDomain
		.QCProject = QCProject
		.QCTestSetPath = QCTestSetPath
		.QCTestSetName = QCTestSetName
		logs "Log ALM Server URL:"+Server+",With Username:"+UserName+",With Password:"+Password+",in the ALM Domain:"+QCDomain&_
		     ",ALM Project:"+QCProject+",Test Set Parent Path:"+QCTestSetPath+",Test Set name is:"+QCTestSetName
		logs "Run this test set:"+QCTestSetName+" now ...."
		.Run
		logs "Run From ALM Server Completed...."
	End With
	
	Set QCTestSetExec = Nothing
End Sub

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Class QCRunTestSet
	
'Public Variables
	
	''' <summary>
	''' QC Server URL (string)
	''' </summary>
	''' <remarks></remarks>
	Public Server
	
	''' <summary>
	''' UserName (string)
	''' </summary>
	''' <remarks></remarks>
	Public UserName
	
	''' <summary>
	''' Password (string)
	''' </summary>
	''' <remarks></remarks>
	Public Password
	
	''' <summary>
	''' Quality Center Domain (string)
	''' </summary>
	''' <remarks></remarks>
	Public QCDomain
	
	''' <summary>
	''' QC Project (string)
	''' </summary>
	''' <remarks></remarks>
	Public QCProject
	
	''' <summary>
	''' QC TestSet Folder Path (string)
	''' </summary>
	''' <remarks>Root\TestSetFolder\TestSetSubFolder</remarks>
	Public QCTestSetPath
	
	''' <summary>
	''' Target TestSet Name (string)
	''' </summary>
	''' <remarks></remarks>
	Public QCTestSetName	
	
	''' <summary>
	''' Recipient list from QC's Automation tab
	''' </summary>
	''' <remarks></remarks>
	Public EMailTo
	
	''' <summary>
	''' TSTestFactory manages test instances (TSTest objects) in a test set
	''' </summary>
	''' <remarks></remarks>
	Public TSTestFactory
	
	''' <summary>
	''' Number of blocked tests after completion of scheduler (integer)
	''' </summary>
	''' <remarks></remarks>
	Public iBlocked
	
	''' <summary>
	''' Number of failed tests after completion of scheduler (integer)
	''' </summary>
	''' <remarks></remarks>
	Public iFailed
	
	''' <summary>
	''' Number of N/A tests after completion of scheduler (integer)
	''' </summary>
	''' <remarks></remarks>
	Public iNA
	
	''' <summary>
	''' Number of NoRun tests after completion of scheduler (integer)
	''' </summary>
	''' <remarks></remarks>
	Public iNoRun
	
	''' <summary>
	''' Number of NotCompleted tests after completion of scheduler (integer)
	''' </summary>
	''' <remarks></remarks>
	Public iNotCompleted
	
	''' <summary>
	''' Number of Passed tests after completion of scheduler (integer)
	''' </summary>
	''' <remarks></remarks>
	Public iPassed
	
	''' <summary>
	''' DateTime stamp at the start of the Scheduling session (DateTime)
	''' </summary>
	''' <remarks></remarks>
	Public dtStartTime

	
'Private Variables
	
	''' <summary>
	''' QuickTest.Application object
	''' </summary>
	''' <remarks></remarks>
	Private qtApp
	
	''' <summary>
	''' TDApiOle object
	''' </summary>
	''' <remarks></remarks>
	Private TDConnection
	
	''' <summary>
	''' TSScheduler object returned by the StartExecution method
	''' </summary>
	''' <remarks></remarks>
	Private TSScheduler
	
	''' <summary>
	''' TestSet Folder object
	''' </summary>
	''' <remarks></remarks>
	Private TSFolder
	
	''' <summary>
	''' Executes the scheduler
	''' </summary>
	''' <remarks></remarks>
	Public Default Sub Run()
		Dim bStatus, dtStartTime
		
		'@see isQCConnected()
		bStatus = isQCConnected
		
		'@see isQTPInstalled()
		If bStatus Then bStatus = isQTPInstalled
		
		If Not bStatus Then Exit Sub
		logs " QC had connected with Server and QTP is installed in the host ,so we can run the test later......"
		Dim TSTreeManager, QCTestSetPath, TSList, QCTestSetName, TestSet, qtTest, sEnvironment, TSReport, EMailTo, ExecutionStatus

		'TestSetTreeManager manages the test set tree and its related test set folders
		Set TSTreeManager = TDConnection.TestSetTreeManager
		logs " Now going to find  the test set in  ALM Test Lab......"
		QCTestSetPath = Me.QCTestSetPath
		'Return the test set tree node from the specified tree path
		logs " the Test set path in test lab is :"&QCTestSetPath
		Set TSFolder = TSTreeManager.NodeByPath(QCTestSetPath)
		logs "the test set name is:"+QCTestSetPath+",going to find this test set from Test Lab..."
		QCTestSetName = Me.QCTestSetName
		'Returns the list of test sets contained in the folder that match the specified pattern. 
		Set TSList = TSFolder.FindTestSets(QCTestSetName)
			
		If TSList.Count = 0 Then
			logs " The TestSet '" & QCTestSetName & "' was not found."&"TSFolder.FindTestSets Exception!"
			Exit Sub
		End If
		
		For Each TestSet in TSList
		     logs  " Found current Test set name is:"&TestSet.Name
			 If LCase(TestSet.Name) = LCase(QCTestSetName) Then	
			     logs  " We found the test set:"+QCTestSetName+" we need to run from test Lab."		   
				 Exit For
			End If
		Next
		
		'This enables database to update immediately when the field value changes
		TestSet.AutoPost = True
		
		'TSTestFactory manages test instances (TSTest objects) in a test set
		Set TSTestFactory = TestSet.TSTestFactory
		Set Me.TSTestFactory = TSTestFactory
		
		'TSTestFactory.NewList("") creates a list of objects according to the specified filter
		For Each qtTest in TSTestFactory.NewList("")
			'Change test status to N/A
			'We do this to ensure all tests have 'not run' before starting execution
			'If the execution errors out, we can keep track of the tests that were not run
			logs  "Reset current Test case status to No Run ,test case name:"+qtTest.Name
			qtTest.Field("TC_STATUS") = "No Run"
			qtTest.Post
		Next
		
		'Refresh TS and TSFolder
		TestSet.Refresh : TSFolder.Refresh
		logs  " Refresh the test suite status with latest No Run status......."
		'TSReport represents the execution report settings for the current test set
		Set TSReport = TestSet.ExecutionReportSettings
		TSReport.Enabled = True
		
		'This retrieves the EMail list from the recipients list from QC
		EMailTo = TSReport.EMailTo : Me.EMailTo = EMailTo
		logs  "retrieves the ALM Automation Tab EMail list from the recipients list in ALM... "
		On Error Resume Next
			'TestSet.StartExecution returns the TSScheduler object and starts the Execution controller
			Set TSScheduler = TestSet.StartExecution("")
            logs "Begin to run ALM Test later ...."
			If Err.Number <> 0 Then
					logs  Err.Description & vbNewLine & vbNewLine & "Unable to create the TSScheduler" & _
						"object. Please ensure the ALM Client Registration is complete before " & _
						"executing RunTestSet."&"RunTestSet.Run->TSScheduler Exception!"
	
					On Error Goto 0
					Exit Sub
			End If
		On Error Goto 0
		
		'Run all tests on localhost
	'	TSScheduler.RunAllLocally ="pdeauto06.fc.hp.com"
		'Logging enabled
		TSScheduler.LogEnabled = True
		logs "Enabled the test execution with log report in ALM ."
		dtStartTime = Now : Me.dtStartTime = dtStartTime
		logs "Set  the test start time is:"+CStr(dtStartTime)
		logs  "Now is running the QC Test suite ,please wait for few minutes to complete......"
		'Start testSet run
		TSScheduler.Run
		
		'ExecutionStatus represents the execution status of the scheduler
		Set ExecutionStatus = TSScheduler.ExecutionStatus
		
		'Wait until all tests are complete running
		WaitWhileTestRunning ExecutionStatus
		
		Set ExecutionStatus = Nothing
	End Sub
	
	
'Private Methods
	
	''' <summary>
	''' Loops and retrieves Scheduler's Finished property 
	''' until all tests have completed running
	''' </summary>
	''' <remarks></remarks>
	Private Sub WaitWhileTestRunning(ByVal ExecutionStatus)
		Dim RunFinished: RunFinished = False
		Dim EventsList,ExecEventInfoObj,TestExecStatusObj
		Dim iterator:iterator=0
		Dim i
		While RunFinished = False
		    iterator=iterator+1
		    logs  "Iteration : "+cstr(iterator)+", Wait the test set status to be finished....,current Finished status is true or not:"&RunFinished			
			ExecutionStatus.RefreshExecStatusInfo "all", True
			RunFinished = ExecutionStatus.Finished
			Set EventsList = ExecutionStatus.EventsList
            For Each ExecEventInfoObj In EventsList
				logs "Event: " & ExecEventInfoObj.EventDate & " " & _
						ExecEventInfoObj.EventTime & " " & _
						"Event Type: " & ExecEventInfoObj.EventType & " [Event types: " & _
						"1-fail, 2-finished, 3-env fail, 4-timeout, 5-manual]"
            Next

            logs ExecutionStatus.Count & " exec status"
			For i = 1 To ExecutionStatus.Count
				Set TestExecStatusObj = ExecutionStatus.Item(i)
				logs "Iteration " & iterator & " Status: " & _
							" Test " & TestExecStatusObj.testID & _
							" ,Test instance " & TestExecStatusObj.TestInstance & _
							" ,order " & TestExecStatusObj.TSTestID & " " & _
							TestExecStatusObj.Message & ", status=" & _
							TestExecStatusObj.Status
			Next
      

			WScript.Sleep(60000)
		Wend
		logs  "All the test cases had been finished now,wrapped all execution log and send email ......."
	End Sub
	
	''' <summary>
	''' Returns true if QC Connection was successful with the supplied input
	''' </summary>
	''' <remarks></remarks>
	Private Function isQCConnected()
		isQCConnected = False
		logs  "Now going to connect to ALM Server......"
		Dim UserName, Password
 		
 		UserName = Me.UserName
 		Password = Me.Password

		On Error Resume Next
			Set TDConnection = CreateObject("TDApiOle80.TDConnection")
			
			If Err.Number <> 0 Then
				logs  " Unable to create an instance of the TestDirector API " & _
						"OLE (TestDirector Connection) ObjectTDConnection Exception!"
				Err.Clear : Exit Function
			End If

			With TDConnection
				'Create a connection with the QC Server
				.InitConnectionEx Server

				If Err.Number <> 0 Then
					logs  Err.Description&"TDConnection.InitConnectionEx Exception!"
					Exit Function
				End If

				'Login to QC
				.Login UserName, Password
				logs "Login ALM with username:"+UserName+",Password:"+Password

				If Err.Number <> 0 Then
					logs  Err.Description&"TDConnection.Login Exception!"
					Exit Function
				ElseIf Not .LoggedIn Then
					logs  " Unable to login to Quality Center. Please verify your login " & _
							"credentials."&"TDConnection.Login Exception!"
					Exit Function
				End If

				'Connect to QC Project
				.Connect QCDomain, QCProject
				 logs "Login ALM  with Domain:"+QCDomain+",Project name is:"+QCProject

				If Err.Number <> 0 Then
					logs  Err.Description&"TDConnection.Connect Exception!"
					Exit Function
				ElseIf Not .ProjectConnected Then
					logs  "Unable to connect to '" & QCDomain & "/" & QCProject & "'."&"TDConnection.Connect Exception!" 
					Exit Function
				End If
				logs "Login ALM process completed......"
				isQCConnected = True
			End With

		On Error Goto 0
	End Function
	
	''' <summary>
	''' Returns true if QTP is installed on the target machine
	''' </summary>
	''' <remarks></remarks>
	Private Function isQTPInstalled()
	    logs  "Now is checking the QTP is installed in the host or not,also will lauch the QTP application......"
		isQTPInstalled = False
		
		On Error Resume Next
			Set qtApp = GetObject("", "QuickTest.Application")
			
			If Err.Number<>0 Then
				logs  Err.Description&"QuickTest.Application Exception!QTP not installed in the host ,exit function......"
				Exit Function
			Else
			    logs "Lanched QTP now ..."
				qtApp.Launch()
				qtApp.Visible = True
				logs "Had Checked QTP installed in the host ,just exit QTP now ..."
				qtApp.Quit 
				isQTPInstalled = True
			End If
			Set qtApp=nothing
		On Error Goto 0
	End Function
	
	''' <summary>
	''' Returns a HTML log of all tests from the executed TestSet
	''' This HTML is exactly the HTML that is used by Quality Center
	''' * please change it as per your needs
	''' </summary>
	''' <remarks></remarks>
	Private Function get_TSExecutionLog()
		Dim color, style, TSTestFactory, TSList, ix, html, sTest, sStatus, sTester, sActualTester, dtDate, dtExecTime
		
		'color = green;red
		color = "46D44B;D41743"
		
		'default html style
		style = "font-size: 11px; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid #eee;"
		
		Set TSTestFactory = Me.TSTestFactory
		Set TSList = TSTestFactory.NewList("")
		
		'Loop through all tests in the TestSet list and retrieve their status
		For ix = 1 To TSList.Count
			html = html & "<tr>"
			
			'Test Name
			html = html & "<td style='" & style & "'>" & TSList.item(ix).field("TSC_NAME") & "</td>"
			
			'Status
			sStatus = TSList.item(ix).field("TS_EXEC_STATUS")
			Select Case sStatus
				Case "Passed"	: html = html & "<td style='color: #" & Trim(Split(color, ";")(0)) & ";" & style & "'>" & sStatus & "</td>"
				Case "Failed"	: html = html & "<td style='color: #" & Trim(Split(color, ";")(1)) & ";" & style & "'>" & sStatus & "</td>"
				Case Else		: html = html & "<td style='" & style & "'>" & sStatus & "</td>"
			End Select
			
			'Tester
			html = html & "<td style='" & style & "'>" & TSList.item(ix).field("TC_TESTER_NAME") & "</td>"
			
			'Actual Tester
			html = html & "<td style='" & style & "'>" & TSList.item(ix).field("TC_ACTUAL_TESTER") & "</td>"
			
			'DateTime stamp
			html = html & "<td style='" & style & "'>" & Date & "</td>"
			
			'Execution Time
			html = html & "<td style='" & style & "'>" & TSList.item(ix).field("TC_EXEC_TIME") & "</td>"
			html = html & "</tr>"
		Next
		
		get_TSExecutionLog = html
	End Function
	
	''' <summary>
	''' Returns the number of tests passed, failed and not completed
	''' </summary>
	''' <remarks></remarks>
	Private Sub load_ExecutionRunStatus()
        Dim TSTestFactory, TSList, ix, iBlocked, iFailed, iNA, iNoRun, iNotCompleted, iPassed
		
		Set TSTestFactory = Me.TSTestFactory
        Set TSList = TSTestFactory.NewList("")
        logs  " Loop to find all the test suites status as blow......"
		'Loop through all tests in the testSet list and retrieve status
        For ix = 1 To TSList.Count
            Select Case LCase(TSList.item(ix).Field("TS_EXEC_STATUS"))
                Case "blocked" : iBlocked = iBlocked + 1
                Case "failed" : iFailed = iFailed + 1
                Case "n/a" : iNA = iNA + 1
                Case "no run" : iNoRun = iNoRun + 1
                Case "not completed" : iNotCompleted = iNotCompleted + 1
                Case "passed" : iPassed = iPassed + 1
            End Select
        Next

        If iBlocked = "" Then iBlocked = 0
        If iFailed = "" Then iFailed = 0
        If iNA = "" Then iNA = 0
        If iNoRun = "" Then iNoRun = 0
        If iNotCompleted = "" Then iNotCompleted = 0
        If iPassed = "" Then iPassed = 0

        Me.iBlocked = iBlocked
        Me.iFailed = iFailed
        Me.iNA = iNA
        Me.iNoRun = iNoRun
        Me.iNotCompleted = iNotCompleted
        Me.iPassed = iPassed	
	End Sub

	''' <summary>
	''' Sends an email to the distribution list
	''' </summary>
	''' <remarks></remarks>
	Private Sub TDSendMail()
 		Dim EMailTo : EMailTo = Me.EMailTo
		Dim QCTestSetName : QCTestSetName = Me.QCTestSetName
		
		If EMailTo = "" Then Exit Sub
		
		load_ExecutionRunStatus()
 		
		TDConnection.SendMail EMailTo, "", "PAF Smoke Testing ALM Execution Report-" & QCTestSetName&" "&now(), sHTML
	End Sub
	
	
'Private Properties	

	''' <summary>
	''' This HTML is exactly the same HTML as used by Quality Center!
	''' Please modify it according to your needs
	''' </summary>
	''' <remarks></remarks>
	Private Property Get sHTML()
		Dim Server 			: Server = Me.Server
		Dim QCTestSetName		: QCTestSetName = Me.QCTestSetName
		Dim QCDomain			: QCDomain = Me.QCDomain
		Dim QCProject		: QCProject = Me.QCProject
		Dim dtStartTime		: dtStartTime = Me.dtStartTime
		Dim iBlocked			: iBlocked = Me.iBlocked
        Dim iFailed			: iFailed = Me.iFailed
        Dim iNA				: iNA = Me.iNA
        Dim iNoRun			: iNoRun = Me.iNoRun
        Dim iNotCompleted		: iNotCompleted = Me.iNotCompleted
        Dim iPassed			: iPassed = Me.iPassed
		
		sHTML = 	"<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & _
				"<HTML>" & _
					"<HEAD>" & _
						"<TITLE>QC TestSet Execution - Custom Report</TITLE>" & _
						"<STYLE>" & _
							".textfont {font-weight: normal; font-size: 12px; color: #000000; font-family: verdana, arial, helvetica, sans-serif }" & _
							".owner {width:100%; border-right: #6d7683 1px solid; border-top: #6d7683 1px solid; border-left: #6d7683 1px solid; border-bottom: #6d7683 1px solid; background-color: #a3a9b1; padding-top: 3px; padding-left: 3px; padding-right: 3px; padding-bottom: 10px; }" & _
							".product {color: white; font-size: 22px; font-family: Calibri, Arial, Helvetica, Geneva, Swiss, SunSans-Regular; background-color: #54658c; padding: 5px 10px; border-top: 5px solid #a9b2c5; border-right: 5px solid #a9b2c5; border-bottom: #293f6f; border-left: 5px solid #a9b2c5;}" & _
							".rest {color: white; font-size: 24px; font-family: Calibri, Arial, Helvetica, Geneva, Swiss, SunSans-Regular; background-color: white; padding: 10px; border-right: 5px solid #a9b2c5; border-bottom: 5px solid #a9b2c5; border-left: 5px solid #a9b2c5 }" & _
							".chl {font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; width: 17%; height: 20px; border-bottom: 1px solid white }" & _
							"a {color: #336 }" & _
							"a:hover {color: #724e6d }" & _
							".ctext {font-size: 11px; padding-right: 5px; padding-left: 5px; width: 80%; height: 20px; border-bottom: 1px solid #eee }" & _
							".hl {color: #724e6d; font-size: 12px; font-weight: bold; background-color: white; height: 20px; border-bottom: 2px dotted #a9b2c5 }" & _
							".space {height: 10px;}" & _
							"h3 {font-weight: bold; font-size: 11px; color: white; font-family: verdana, arial, helvetica, sans-serif;}" & _
						"</STYLE>" & _
						"<META content='MSHTML 6.00.2800.1106'>" & _
					"</HEAD>" & _
					"<body leftmargin='0' marginheight='0' marginwidth='0' topmargin='0'>" & _
						"<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & _
							"<tr>" & _
								"<td class='product'>Quality Center</td>" & _
							"</tr>" & _
							"<tr>" & _
								"<td class='rest'>" & _
									"<table class='space' width='100%' border='0' cellspacing='0' cellpadding='0'>" & _
										"<tr>" & _
											"<td></td>" & _
										"</tr>" & _
									"</table>" & _
									"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
										"<tbody>" & _
											"<tr>" & _
												"<th class='hl' align='left'>Test Set Execution Summary</th>" & _
											"</tr>" & _
											"<tr>" & _
												"<th class='space' align='left'></th>" & _
											"</tr>" & _
										"</tbody>" & _
									"</table>" & _
									"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
										"<tbody>" & _
											"<tr>" & _
												"<td>" & _
													"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
														"<tbody>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>Server</td>" & _
																"<td class='ctext'>" & Server & "</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>Domain Name</td>" & _
																"<td class='ctext'>" & QCDomain & "</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>Project Name</td>" & _
																"<td class='ctext'>" & QCProject & "</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>TestSet</td>" & _
																"<td class='ctext'>" & QCTestSetName & "</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>Started</td>" & _
																"<td class='ctext'>" & dtStartTime & "</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>Finished</td>" & _
																"<td class='ctext'>" & Now & "</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' width='20%'>Duration</td>" & _
																"<td class='ctext'>" & DateDiff("n", dtStartTime, Now) & " minutes " & DateDiff("s", dtStartTime, Now) Mod 60 & " seconds</td>" & _
															"</tr>" & _
														"</tbody>" & _
													"</table>" & _
												"</td>" & _
											"</tr>" & _
											"<tr>" & _
												"<td class='space'></td>" & _
											"</tr>" & _
										"</tbody>" & _
									"</table>" & _
									"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
										"<tbody>" & _
											"<tr>" & _
												"<td>" & _
													"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
														"<tbody>" & _
															"<tr>" & _
																"<td class='chl' style='width:17%'>Blocked</td>" & _
																"<td class='chl' style='width:17%'>Failed</td>" & _
																"<td class='chl' style='width:17%'>N/A</td>" & _
																"<td class='chl' style='width:17%'>No Run</td>" & _
																"<td class='chl' style='width:17%'>Not Completed</td>" & _
																"<td class='chl' style='width:17%'>Passed</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='ctext' style='width:17%'>" & iBlocked & "</td>" & _
																"<td class='ctext' style='width:17%'>" & iFailed & "</td>" & _
																"<td class='ctext' style='width:17%'>" & iNA & "</td>" & _
																"<td class='ctext' style='width:17%'>" & iNoRun & "</td>" & _
																"<td class='ctext' style='width:17%'>" & iNotCompleted & "</td>" & _
																"<td class='ctext' style='width:17%'>" & iPassed & "</td>" & _
															"</tr>" & _
														"</tbody>" & _
													"</table>" & _
												"</td>" & _
											"</tr>" & _
											"<tr>" & _
												"<td class='space'></td>" & _
											"</tr>" & _
										"</tbody>" & _
									"</table>" & _
									"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
										"<tbody>" & _
											"<tr>" & _
												"<td>" & _
													"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
														"<tbody>" & _
															"<tr>" & _
																"<td colspan='4' class='owner'><h3>Test Set - " & QCTestSetName & "</h3></td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' style='width:15%'>Name:</td>" & _
																"<td class='ctext' style='width:35%'><a href=''>" & QCTestSetName & "</a></td>" & _
																"<td class='chl' style='width:15%'>Test Set ID</td>" & _
																"<td class='ctext' style='width:35%'>4</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' style='width:15%'>Test Set Folder</td>" & _
																"<td class='ctext' style='width:35%'>" & QCTestSetPath & "</td>" & _
																"<td class='chl' style='width:15%'>Type</td>" & _
																"<td class='ctext' style='width:35%'>Default</td>" & _
															"</tr>" & _
															"<tr>" & _
																"<td class='chl' style='width:15%'>Target Cycle</td>" & _
																"<td class='ctext' style='width:35%'></td>" & _
																"<td class='chl' style='width:15%'>Baseline</td>" & _
																"<td class='ctext' style='width:35%'></td>" & _
															"</tr>" & _
														"</tbody>" & _
													"</table>" & _
												"</td>" & _
											"</tr>" & _
											"<tr>" & _
												"<td class='space'></td>" & _
											"</tr>" & _
										"</tbody>" & _
									"</table>" & _
									"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
										"<tbody>" & _
											"<tr>" & _
												"<th class='hl' align='left'>Executed Tests</th>" & _
											"</tr>" & _
											"<tr>" & _
												"<th class='space' align='left'></th>" & _
											"</tr>" & _
										"</tbody>" & _
									"</table>" & _
									"<table class='textfont' cellspacing='0' cellpadding='0' width='100%' align='center' border='0'>" & _
										"<tbody>" & _
											"<tr>" & _
												"<td style='font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid white;'>Test</td>" & _
												"<td style='font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid white;'>Status</td>" & _
												"<td style='font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid white;'>Responsible Tester</td>" & _
												"<td style='font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid white;'>Tester</td>" & _
												"<td style='font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid white;'>Exec Date</td>" & _
												"<td style='font-size: 10px; font-weight: bold; background-color: #eee; padding-right: 5px; padding-left: 5px; height: 20px; border-bottom: 1px solid white;'>Exec Time</td>" & _
											"</tr>" & _
											get_TSExecutionLog() & _
										"</tbody>" & _
									"</table>" & _
								"</td>" & _
							"</tr>" & _
						"</table>" & _
					"</body>" & _
				"</HTML>"
	End Property

	
'Class Initialize & Terminate
	
	''' <summary>
	''' Releases connections and sends mail after TSScheduler execution
	''' </summary>
	''' <remarks></remarks>
	Private Sub Class_Terminate()
		If IsObject(TSScheduler) Then
			If Not TSScheduler Is Nothing Then
				TSFolder.Refresh : WScript.Sleep(5000)
		        logs  "Now go to send the execution email..."
				'Send an email to the distribution list
				TDSendMail()
				
				Set TSScheduler = Nothing
			End If
		End If
		
		On Error Resume Next
			'Disconnect TD session
			TDConnection.Disconnect
			logs "Disconnect with ALM Server successfully!"
			'Disconect and quit QTP
			If IsObject(qtApp) Then
				If qtApp.TDConnection.IsConnected Then qtApp.TDConnection.Disconnect
				qtApp.Quit
			End If
		On Error Goto 0
		logs "Completed all the execution over ,Done !"
		Set qtApp = Nothing
		Set TDConnection = Nothing
	End Sub

End Class ''RunTestSet
sub PromptRun()
        dim WshShell,fso,objOutputFile
        Set WshShell = WScript.CreateObject("WScript.Shell") 
		Set fso = CreateObject("Scripting.FileSystemObject") 
		Set objOutputFile = fso.OpenTextFile("c:\temp.vbs", 2, -1) 
		objOutputFile.WriteLine "MsgBox ""PAF Daily Smoke Testing is running in the host now ,please do not interactive with this host .......""" 
		objOutputFile.Close 
		WshShell.Run "c:\temp.vbs", 1, False
		Set objOutputFile=Nothing
		Set WshShell=Nothing
		Set fso=nothing
end sub
Sub logs(logline)
	        Dim fso,objOutputFile
			Set fso = CreateObject("Scripting.FileSystemObject") 
		'	If fso.FileExists(strExecutionLogFile) Then
		'	   fso.DeleteFile strExecutionLogFile
		'	End If 
			Set objOutputFile = fso.OpenTextFile(strExecutionLogFile,8, true) 
			objOutputFile.WriteLine Now()&"  "& logline
			objOutputFile.Close 		
			Set objOutputFile=Nothing
			Set fso=Nothing
End Sub
Sub deleteLogFile()
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject") 
   If fso.FileExists(strExecutionLogFile) Then
	  fso.DeleteFile strExecutionLogFile
   End If 
   Set fso=nothing
End sub
call PromptRun()
Call deleteLogFile()
Call RunTestSet(Server, UserName, Password, QCDomain, QCProject, QCTestSetPath, QCTestSetName)