'#########################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution	##
'##																						##
'## September 18, 2013																	##
'##																						##
'## Version 1.0																			##
'#########################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 7 Then 
	ShowError(3) 
End If
'Set Culture - en-us
SetLocale(1033)

Dim Host, MetricState, TargetIDList, SiteList, Username, Password, Domain, strSite
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
TargetIDList = WScript.Arguments(2)
SiteList = WScript.Arguments(3)
Username = WScript.Arguments(4)
Password = WScript.Arguments(5)
Domain = WScript.Arguments(6)

Dim arrSites, arrTargetsIDs, arrMetrics
arrSites = Split(SiteList,",")
arrTargetsIDs = Split(TargetIDList,",")
arrMetrics = Split(MetricState,",")

Dim objSWbemLocator, objSWbemServices, colItems

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
	
Dim Counter, FullUserName,objItem
Counter = 0
	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If

	If Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
	End If
	if Err.Number = -2147024891 Then
		CALL ShowError(2, strComputer)
	End If
	If Err Then CALL ShowError(1, Host)

	objSWbemServices.Security_.ImpersonationLevel = 3
	
	'Status IIS Services ( IISADMIN , W3SVC )
	Dim colStatus, objStatus, Status
	Set colStatus = objSWbemServices.ExecQuery("SELECT State from Win32_Service Where Name ='W3SVC'",,16)
			If colStatus.Count <> 0 Then
				For Each objStatus In colStatus
					If objStatus.State <> "Running" Then
					Status = 0
						For Each strSite In arrSites
						If arrMetrics(18)=1 Then _
							Call Output("37:9",arrTargetsIDs(Counter),Status,strSite)	
						Counter = Counter + 1
						Next
					Else

	For Each strSite In arrSites
		If IsEmpty(strSite) = False Then
		Set colItems = objSWbemServices.ExecQuery("Select AnonymousUsersPersec,BytesReceivedPerSec,BytesSentPerSec,BytesTotalPerSec,ConnectionAttemptsPerSec,CurrentAnonymousUsers,CurrentConnections,CurrentNonAnonymousUsers,LogonAttemptsPersec,Name,ServiceUptime,TotalbytesReceived,TotalBytesSent,TotalBytesTransferred,TotalFilesreceived,TotalFilesSent,TotalFilesTransferred,TotalLogonAttempts,TotalMethodRequests,TotalMethodRequestsPerSec from Win32_PerfFormattedData_W3SVC_WebService where Name='" & strSite & "'",,16) 
			If IsEmpty(colItems) = False Then
			For Each objItem in colItems
					'Status
						'Dim Status
						If objItem.ServiceUptime = 0 Then 
							Status = 0
							If arrMetrics(18)=1 Then _
								Call Output("37:9",arrTargetsIDs(Counter),Status,strSite)
						Else
							Status = 1
							If arrMetrics(18)=1 Then _
								Call Output("37:9",arrTargetsIDs(Counter),Status,strSite)
						End If
						
					'AnonymousUsersPersec
					If (arrMetrics(0)=1 and Status = 1) Then _
					Call Output("161:4",arrTargetsIDs(Counter),objItem.AnonymousUsersPersec,strSite)
					'kBytesReceivedPerSec
					If (arrMetrics(1)=1 and Status = 1) Then _
					CALL Output("104:4",arrTargetsIDs(Counter),FormatNumber(objItem.BytesReceivedPerSec/1024),strSite)
					'kBytesSentPerSec
					If (arrMetrics(2)=1 and Status = 1) Then _
					CALL Output("163:4",arrTargetsIDs(Counter),FormatNumber(objItem.BytesSentPerSec/1024),strSite)
					'kBytesTotalPerSec
					If (arrMetrics(3)=1 and Status = 1) Then _
					CALL Output("130:4",arrTargetsIDs(Counter),FormatNumber(objItem.BytesTotalPerSec/1024),strSite)
					'ConnectionAttemptsPerSec
					If (arrMetrics(4)=1 and Status = 1) Then _
					CALL Output("207:4",arrTargetsIDs(Counter),objItem.ConnectionAttemptsPerSec,strSite)
					'CurrentAnonymousUsers
					If (arrMetrics(5)=1 and Status = 1) Then _
					Call Output("206:4",arrTargetsIDs(Counter),objItem.CurrentAnonymousUsers,strSite)
					'CurrentConnections
					If (arrMetrics(6)=1 and Status = 1) Then _
					Call Output("83:4",arrTargetsIDs(Counter),objItem.CurrentConnections,strSite)
					'CurrentNonAnonymousUsers
					If (arrMetrics(7)=1 and Status = 1) Then _
					CALL Output("76:4",arrTargetsIDs(Counter),objItem.CurrentNonAnonymousUsers,strSite)
					'LogonAttemptsPersec
					If (arrMetrics(8)=1 and Status = 1) Then _
					CALL Output("153:4",arrTargetsIDs(Counter),objItem.LogonAttemptsPersec,strSite)
					'ServiceUptime
					If (arrMetrics(9)=1 and Status = 1) Then _
					CALL Output("34:4",arrTargetsIDs(Counter),objItem.ServiceUptime,strSite)
					'TotalkbytesReceived
					If (arrMetrics(10)=1 and Status = 1) Then _
					CALL Output("146:4",arrTargetsIDs(Counter),FormatNumber(objItem.TotalbytesReceived/1024),strSite)
					'TotalkBytesSent
					If (arrMetrics(11)=1 and Status = 1) Then _
					CALL Output("69:4",arrTargetsIDs(Counter),FormatNumber(objItem.TotalBytesSent/1024),strSite)
					'TotalkBytesTransferred
					If (arrMetrics(12)=1 and Status = 1) Then _
					CALL Output("173:4",arrTargetsIDs(Counter),FormatNumber(objItem.TotalBytesTransferred/1024),strSite)
					'TotalFilesreceived
					If (arrMetrics(13)=1 and Status = 1) Then _
					CALL Output("169:4",arrTargetsIDs(Counter),objItem.TotalFilesreceived,strSite)
					'TotalFilesSent
					If (arrMetrics(14)=1 and Status = 1) Then _
					CALL Output("128:4",arrTargetsIDs(Counter),objItem.TotalFilesSent,strSite)
					'TotalFilesTransferred
					If (arrMetrics(15)=1 and Status = 1) Then _
					CALL Output("187:4",arrTargetsIDs(Counter),objItem.TotalFilesTransferred,strSite)
					'TotalLogonAttempts
					If (arrMetrics(16)=1 and Status = 1) Then _
					CALL Output("71:4",arrTargetsIDs(Counter),objItem.TotalLogonAttempts,strSite)
					'TotalMethodRequests
					If (arrMetrics(17)=1 and Status = 1) Then _
					CALL Output("168:4",arrTargetsIDs(Counter),objItem.TotalMethodRequests,strSite)
					'TotalMethodRequestsPerSec
					If (arrMetrics(19)=1 and Status = 1) Then _
					CALL Output("28:4",arrTargetsIDs(Counter),objItem.TotalMethodRequestsPerSec,strSite)
					
				Next
			Else
				'If there is no response in WMI query
				CALL ShowError(5, Host)
			End If
		End If
			Counter = Counter + 1
	Next	
			If Err.number <> 0 Then
				CALL ShowError(5, Host)
				Err.Clear
			End If
					End If
				Next
			Else
				'If there is no response in WMI query
				CALL ShowError(5, Host)
			End If
	



	
Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Sub Output(SourceUUID, TargetUUID, SourceValue, SourceObject)
	If SourceObject <> "" Then
		If SourceValue <> "" Then
			WScript.Echo ToUTC() & "|" & SourceUUID & "|" & TargetUUID & "|" & SourceValue & "|" & SourceObject & vbCr
		Else
			CALL ShowError(5, Host) 
		End If			
	Else
		If SourceValue <> "" Then
			WScript.Echo ToUTC() & "|" & SourceUUID & "|" & TargetUUID & "|" & SourceValue & vbCr 
		Else
			CALL ShowError(5, Host) 
		End If
	End If
End Sub

Function ToUTC()
	Dim dtmDateValue, dtmAdjusted
	Dim objShell, lngBiasKey, lngBias, k, UTC
	dtmDateValue = Now()
	'Obtain local Time Zone bias from machine registry.
	Set objShell = CreateObject("Wscript.Shell")
	lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
	If (UCase(TypeName(lngBiasKey)) = "LONG") Then
		lngBias = lngBiasKey
		ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
			lngBias = 0
		For k = 0 To UBound(lngBiasKey)
			lngBias = lngBias + (lngBiasKey(k) * 256^k)
		Next
	End If
	'Convert datetime value to UTC.
	UTC = DateAdd("n", lngBias, dtmDateValue)
	ToUTC =  FormatDateTime(UTC,2) & " " & FormatDateTime(UTC,3)
End Function
