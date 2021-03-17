Option Explicit

public wshShell, args, arg1, arg2,arg3, objFS0, objFolder, objStartFolder, objFile, colFiles, DotNetInstall, DotNetInstallLog, DotNetInstallConf, DotNetInstallDir, DotNetInstallDataDir, FileContents, Line, Format, PromptContinue, ControllerHost, ControllerAccountName, ControllerApplicationName, ControllerAccessKey, ControllerAutoDiscovery, xmlDoc, xmlDocControllerInfo, xmlDocControllerHost, xmlDocControllerPort, xmlDocControllerSSL, xmlDocApplicationName, xmlDocTierName, xmlDocNodeName, xmlDocUseEncryptedCreds, xmlDocCredsStoreFile, xmlDocCredsStorePass, xmlDocAccountName, xmlDocAccountKey, xmlDocKeystoreFile, xmlDocKeystorePass, xmlDocAccountInfo, xmlDocApplicationInfo, xmlDocAutoDiscoveryInfo, iisRestart, DotNetAgentServiceState, cmd, cmd1, WinVer, WinOS, applist, a, Result, registryInstallationDir, registryUninstallSpace, regDotNetAgentFolder, strKeyPath, objectReg, arrSubKeys, searchString, subkey, regUninstallString, regDisplayName, WinArch, WinHfInstalled, JavaAgentType, JavaAgentSourceDir, subFolder, subFolders, JavaAgentRootDir, JavaAgentInstallDir, JavaAgentConfDir, objSubFolder, JavaAgentControllerXML, JavaAgentSCSTool, JavaAgentSecretKeystore, JavaAgentSecretKeystorePass, JavaAgentControllerKeystore, JavaAgentSourceSecretKeystore, JavaAgentSourceControllerKeystore, JavaAgentControllerKeystorePass, CustomerApplication, appStartUpScript, appStartUpScriptPath, strText, LineBuffer, JavaAgentHomeDir, JavaAgentTierName, JavaAgentNodeName, maInstallSource, maHomeDir, objShell, FilesInZip, maConfDir, maControllerXML, maSCSTool, maSourceSecretKeystore, maSourceControllerKeystore, maSecretKeystore, maSecretKeystorePass, maControllerKeystore, maControllerKeystorePass, maServiceState, maJava, maTierName, maNodeName, appTierName, appNodeName, ProcessId, ProcessName, RegExp, arrayStandaloneDotNetApps, Match, Matches, Results, MonitorStandaloneDotNet, xmlDocAppDynamicsAgent, StandaloneDotNetAppName, StandaloneDotNetAppTierName, xmlDocStandaloneApp, xmlDocStandaloneAppTier, xmlDocStandaloneApplications, xmlDocStandaloneAppAttr, xmlDocStandaloneAppTierAttr, DotNetInstallConfTempl, configProperties, proxyHost, proxyPort, proxyEnabled, proxyInfo, xmlDocControllerProxy, Drive, DiskDrive, maRuntimeParamsFile, java, ArrTomcatOptions, MultiValueName, oReg, IsInstrumented, ValueName, backupDirJava, backupDirMachine, backupDirDotNet, TomcatServiceName, objService, regComponentID, regDotNetVersion, WebSphereASPolicyFile, WebSphereASServerXML, xmlDocPMIService, WebSphereASstatisticSet, xmlDocJvmEntries, JvmParams, IsWebSphereASPolicyFileConfigured, DotNetAgentCurrentVersion, tmpFolder, UpgradeMSI, DotNetAgentUpgradeToVersion, DotNetUpgradeLog, xmlDocSimEnabled, simEnabled, maOutputLog, MachineAgentCurrentVersion, UpgradeZIP, MachineAgentUpgradeToVersion, tmpUpgrFolder, maSourceControllerXML, maSourceAnalyticsAgentProperties, maSourceServiceVmoptions, maMonitorsDir, colFolders, arrCustomMonitors, SkipVmoptionsSetup, SkipMonitorsSetup, tmpUpgrMonFolder, colFolder, JavaAgentSourceControllerXML, JavaAgentCurrentVersion, JavaAgentUpgradeToVersion, JavaAgentJAR, UpgrDir, colFile, FileType, autodeployXML, strComputerName, xmlDocScenario, xmlDocScenarioAppName, autoDeploymentType, autoAgentName, autoProdMon, autoAppName, autoDiscoIIS, autoRestIIS, autoStandMon, autoStandOpts, autoTomcatServiceName, autoTierName, autoNodeName, autoAppType, autoAppStartScript, autoAppConf, autoAppJVM, tmpInstrFolder, sourceJavaAgentType, autoAgentPath, autoSimEnabled, sourceJavaAgentVersion, sourceFile, dstFolder, FileName, tmpUninstFolder, xmlDocControllerProxyInfoHost, xmlDocControllerProxyInfoPort, xmlDocControllerProxyInfoEnabled, xmlDocControllerApiInfoApiApp, xmlDocControllerApiInfoApiTokProd, xmlDocControllerApiInfoApiTokNonprod, ControllerUriApiApp, ControllerTokenApiProd, ControllerTokenApiNonprod, restReq, ApiAccessToken, controllerHttpState, url, MachineAgentOnly, autoAppAgentExists, xmlDocControllerHostProd, xmlDocControllerPortProd, xmlDocControllerProtocolProd, ControllerUrlProd, xmlDocControllerPortNonprod, xmlDocControllerProtocolNonprod, xmlDocControllerHostNonprod, ControllerUrlNonprod, xmlDocControllerAccountProd, xmlDocControllerAccountNonprod, ControllerUrlProdAccount, ControllerUrlNonprodAccount, ControllerHostProd, ControllerHostNonprod, xmlDocControllerAccessKeyProd, xmlDocControllerAccessKeyNonprod, ControllerUrlProdAccessKey, ControllerUrlNonprodAccessKey, app, registryDataDir, xmlDocControllerApplicationName, ControllerPort, ControllerSSL, ControllerAccountKey, dataDir, SkipConfigSetup, DotNetCompatibility, xmlDocDotNetCompatibilityEnabled, xmlDocUniqueHostId, ControllerHostPortProd, ControllerHostPortNonprod, ControllerHostProtocolProd, ControllerHostProtocolNonprod, ControllerHostProtocol, xmlDocControllerProxyInfoProtocol, proxyProtocol, proxyUrl, autoStartWebLogic, length, startWebLogic, WebLogicServiceName, domainDir, CmdLine, WebLogicPolicyFile, IsWebLogicPolicyFileConfigured, WebLogicDomainRegistryFile, domainName, 

' Commonly used objects are set here
Set wshShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
Set objFS0 = CreateObject("Scripting.FileSystemObject")
Set RegExp = CreateObject("vbscript.regexp")
Set objService = GetObject("Winmgmts:{impersonationlevel=impersonate}!\Root\Cimv2")
Set restReq = CreateObject("Msxml2.ServerXMLHTTP")
Set args = WScript.Arguments

' Commonly used constants are set here
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const HKEY_LOCAL_MACHINE = &H80000002
Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056

' Commonly used vars are set here
WinOS = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
WinReleaseVer = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
WinArch = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
backupDirJava = objFs0.GetFolder("backup\JavaAgent")
backupDirMachine = objFs0.GetFolder("backup\MachineAgent")
backupDirDotNet = objfs0.GetFolder("backup\DotNetAgent")
Set tmpFolder = objFS0.GetFolder("tmp")
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
RegExp.Global = True
RegExp.Pattern = "\d\d\d\d"
For each match in RegExp.Execute(WinOS)
	WinVer = match
Next

Function getSAASControllerDetails()

	' Build URL
	url = url & ControllerUriApiApp
	
	' Open connection, set proxy if needed and SSL settings
	On Error Resume Next
	Err.Clear
	restReq.open "GET", url, false
	If proxyEnabled = "true" Then
		proxyUrl = proxyProtocol & "://" & proxyHost & ":" & proxyPort
		restReq.setProxy 2, proxyUrl, ""
	ElseIf proxyEnabled = "false" Then
	End If
	restReq.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
	
	' Set authentication header and send request to Destination
	WScript.Echo ""
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - " & url
	restReq.setRequestHeader "Authorization", "Bearer " & ApiAccessToken
	restReq.send
	
	' Get the HTTP state
	controllerHttpState=restReq.Status
	
	' Call applications API
	If controllerHttpState = 200 Or controllerHttpState = 200200 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Controller connectivity check succeeded (" & controllerHttpState & ")."
		objFS0.CreateTextFile("response.log")
		Set objFile = objFS0.OpenTextFile("response.log", ForWriting)
		objFile.Write restReq.responseText
		Set objFile = objFS0.OpenTextFile("response.log", ForReading)
		applist = Array()
		Do until objFile.atEndOfStream
    		Line = objFile.ReadLine
    		If InStr(Line,"<name>") > 0 Then
        		Line = Mid(Replace(Replace(Line, "<name>", ""), "</name>", ""), 3)
        		ReDim Preserve applist(UBound(applist) + 1)
				applist(UBound(applist)) = Line
    		End If
		Loop
		objFile.Close
		objFS0.DeleteFile("response.log")
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not verify connection to Controller (" & controllerHttpState & ")."
		WScript.Quit
	End If
	On Error Goto 0

End Function

Function WaitForWshShellExec()
	Const maxWait = 60000
	Const waitInLoop = 100
	Dim waitTime
	waitTime = 0
	Do While cmd.Status = 0
		WScript.Sleep waitInLoop
		waitTime = waitTime + waitInLoop
		If waitTime > maxWait Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Script aborted due to timeout during waiting for cmd."
			WScript.Quit
		End If
	Loop
End Function

Function getDrive()
	On Error Resume Next
	Set DiskDrive = objFs0.GetDrive("D:")
	If Err.Description = "Device unavailable" Then
		DiskDrive = "C:"
	Else
		If DiskDrive.DriveType <> 2 Then
			DiskDrive = "C:"
		End If
	End If
	On Error Goto 0
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Selected drive " & DiskDrive & "."
End Function

Function getExistingInstall()

	' Get agent install dir
	If objFS0.FolderExists("D:\AppDynamics\JavaAgent") Then
		Set objFolder = objFS0.GetFolder("D:\AppDynamics\JavaAgent")
		If (objFolder.Files.Count > 0) Or (ObjFolder.SubFolders.Count > 0) Then
			JavaAgentHomeDir = "D:\AppDynamics\JavaAgent"
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Existing Java Agent dir is empty (D:\AppDynamics\JavaAgent)."
			WScript.Quit
		End If
	ElseIf objFS0.FolderExists("C:\AppDynamics\JavaAgent") Then
		Set objFolder = objFS0.GetFolder("C:\AppDynamics\JavaAgent")
		If (objFolder.Files.Count > 0) Or (ObjFolder.SubFolders.Count > 0) Then
			JavaAgentHomeDir = "C:\AppDynamics\JavaAgent"
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Existing Java Agent dir is empty (C:\AppDynamics\JavaAgent)."
			WScript.Quit
		End If
	Else
		 WScript.Echo FormatDateTime(Now, vbLongTime) & " - Couldn't detect existing installation."
		 WScript.Quit
	End If
	JavaAgentInstallDir = JavaAgentHomeDir

	' Get ver and conf folders. Also controller-info.xml
	Set objFolder = objFS0.GetFolder(JavaAgentHomeDir)
	Set objSubFolder = objFolder.SubFolders
	For each subFolder in objSubFolder
		If instr(subFolder, "ver") Then
			JavaAgentHomeDir = subFolder
			JavaAgentConfDir = subFolder & "\conf\"
			JavaAgentControllerXML = JavaAgentConfDir & "controller-info.xml"
		End If
	Next
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Using " & JavaAgentHomeDir & "."
	
	' Retrieve existing settings via controller-info.xml
	If objFS0.FileExists(JavaAgentControllerXML) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Found controller-info.xml."
		Set xmlDoc = _
			CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = "False"
			xmlDoc.Load(JavaAgentControllerXML)
		Set xmlDocControllerHost=xmlDoc.selectsinglenode ("//controller-info/controller-host")
		Set xmlDocAccountName=xmlDoc.selectsinglenode ("//controller-info/account-name")
		ControllerHost = xmlDocControllerHost.Text
		ControllerAccountName = xmlDocAccountName.Text
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not find controller-info.xml. Setup will quit."
		WScript.Quit
	End If

	' Retrieve type of agent via MANIFEST.MF
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Looking for type of Java Agent."
	tmpInstrFolder = tmpFolder & "\instrument"
	If objFS0.FolderExists(tmpInstrFolder) Then 
		Set objFolder = objFS0.GetFolder(tmpInstrFolder)
		objFS0.DeleteFolder(tmpInstrFolder)
		objFS0.CreateFolder(tmpInstrFolder)
	Else
		objFS0.CreateFolder(tmpInstrFolder)
	End If
	objFS0.CopyFile JavaAgentHomeDir & "\javaagent.jar", tmpInstrFolder & "\javaagent.zip"
	call getJavaAgentDetails(tmpInstrFolder & "\javaagent.zip", tmpInstrFolder)
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Java Agent (" & sourceJavaAgentType & ")."
	WScript.Echo ""

End Function

Function getConfigProperties()

	' Load properties file
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	Set objFS0 = CreateObject("Scripting.FileSystemObject")
	configProperties = objFs0.GetFile("conf\properties.xml")
	xmlDoc.Async = "false"
	xmlDoc.Load(configProperties)

	' Load prod Controller details
	Set xmlDocControllerHostProd=xmlDoc.selectsinglenode ("//deployAgent-info/controller/prod/host")
	Set xmlDocControllerPortProd=xmlDoc.selectsinglenode ("//deployAgent-info/controller/prod/port")
	Set xmlDocControllerProtocolProd=xmlDoc.selectsinglenode ("//deployAgent-info/controller/prod/protocol")
	Set xmlDocControllerAccountProd=xmlDoc.selectsinglenode ("//deployAgent-info/controller/prod/account")
	Set xmlDocControllerAccessKeyProd=xmlDoc.selectsinglenode ("//deployAgent-info/controller/prod/accesskey")

	' Load nonprod Controller details
	Set xmlDocControllerHostNonprod=xmlDoc.selectsinglenode ("//deployAgent-info/controller/nonprod/host")
	Set xmlDocControllerPortNonprod=xmlDoc.selectsinglenode ("//deployAgent-info/controller/nonprod/port")
	Set xmlDocControllerProtocolNonprod=xmlDoc.selectsinglenode ("//deployAgent-info/controller/nonprod/protocol")
	Set xmlDocControllerAccountNonprod=xmlDoc.selectsinglenode ("//deployAgent-info/controller/nonprod/account")
	Set xmlDocControllerAccessKeyNonprod=xmlDoc.selectsinglenode ("//deployAgent-info/controller/nonprod/accesskey")

	' Load Controller proxy details
	Set xmlDocControllerProxyInfoHost=xmlDoc.selectsinglenode ("//deployAgent-info/controller-proxy-info/host")
	Set xmlDocControllerProxyInfoPort=xmlDoc.selectsinglenode ("//deployAgent-info/controller-proxy-info/port")
	Set xmlDocControllerProxyInfoEnabled=xmlDoc.selectsinglenode ("//deployAgent-info/controller-proxy-info/enabled")
	Set xmlDocControllerProxyInfoProtocol=xmlDoc.selectsinglenode ("//deployAgent-info/controller-proxy-info/protocol")

	' Load Controller APIs and Authentication Token
	Set xmlDocControllerApiInfoApiApp=xmlDoc.selectsinglenode ("//deployAgent-info/controller-api-info/api-app")
	Set xmlDocControllerApiInfoApiTokProd=xmlDoc.selectsinglenode ("//deployAgent-info/controller-api-info/api-token-prod")
	Set xmlDocControllerApiInfoApiTokNonprod=xmlDoc.selectsinglenode ("//deployAgent-info/controller-api-info/api-token-nonprod")

	' Set variables
	ControllerHostProd = xmlDocControllerHostProd.Text
	ControllerHostPortProd = xmlDocControllerPortProd.Text
	ControllerHostProtocolProd = xmlDocControllerProtocolProd.Text
	ControllerUrlProd = xmlDocControllerProtocolProd.Text & "://" & xmlDocControllerHostProd.Text & ":" & xmlDocControllerPortProd.Text & "/controller"
	ControllerUrlProdAccount = xmlDocControllerAccountProd.Text
	ControllerUrlProdAccessKey = xmlDocControllerAccessKeyProd.Text
	ControllerHostNonprod = xmlDocControllerHostNonprod.Text
	ControllerHostPortNonprod = xmlDocControllerPortNonprod.Text
	ControllerHostProtocolNonprod = xmlDocControllerProtocolNonprod.Text
	ControllerUrlNonprod = xmlDocControllerProtocolNonprod.Text & "://" & xmlDocControllerHostNonprod.Text & ":" & xmlDocControllerPortNonprod.Text & "/controller"
	ControllerUrlNonprodAccount = xmlDocControllerAccountNonprod.Text
	ControllerUrlNonprodAccessKey = xmlDocControllerAccessKeyNonprod.Text
	ControllerUriApiApp = xmlDocControllerApiInfoApiApp.Text
	ControllerTokenApiProd = xmlDocControllerApiInfoApiTokProd.Text
	ControllerTokenApiNonprod = xmlDocControllerApiInfoApiTokNonprod.Text
	proxyHost = xmlDocControllerProxyInfoHost.Text
	proxyPort = xmlDocControllerProxyInfoPort.Text
	proxyEnabled = xmlDocControllerProxyInfoEnabled.Text
	proxyProtocol = xmlDocControllerProxyInfoProtocol.Text

	' Address a problem with Winston agent when proxy is not needed and port is empty
	If proxyEnabled = "false" Then
		proxyPort = "80"
	End If

End Function

Function setAppName()

	' Support for autodeploy option
	If IsEmpty(autoAppName) Then
		WScript.Echo ""
		WScript.Echo "Available applications for " & ControllerHost & ":"
		For Each app in applist
			WScript.stdout.write app
			WScript.stdout.write " | "
		Next
		WScript.Echo ""
		WScript.Echo ""
		WScript.StdOut.Write "Application name: "
		PromptContinue = WScript.StdIn.ReadLine
		ControllerApplicationName = PromptContinue
	Else
		ControllerApplicationName = autoAppName
	End If

	' Validate for space or empty
	If (ControllerApplicationName = "") Or (Left(ControllerApplicationName, 1) = " ") Or (Right(ControllerApplicationName, 1) = " ") Then
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: 'space' at first and last char or no value."
			WScript.Quit
	End If
	
	' Validate if input name matches applist
	For Each app in applist
		If InStr(app, ControllerApplicationName) Then
			Result = "true"
		End If
	Next
	If Result = "true" Then
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Application not part of any known list."
		WScript.Quit
	End If

End Function

Function setIISDotNetMon()

	If IsEmpty(autoDiscoIIS) Then
		WScript.StdOut.Write ""
		WScript.StdOut.Write "IIS Automatic Application Discovery (yes/no)? "
		PromptContinue = WScript.StdIn.ReadLine
		If InStr("|yes|y|YES|Y|", "|" & PromptContinue & "|") Then
			ControllerAutoDiscovery = "true"
			WScript.StdOut.Write "Automatically restart IIS after install (yes/no)? "
			PromptContinue = WScript.StdIn.ReadLine
			If InStr("|yes|y|YES|Y|", "|" & PromptContinue & "|") Then
				iisRestart = "yes"
			ElseIf InStr("|no|n|NO|N|", "|" & PromptContinue & "|") Then
				iisRestart = "no"
			Else
				WScript.Echo ""
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'yes' or 'no'."
				WScript.Quit
			End If
		ElseIf InStr("|no|n|NO|N|", "|" & PromptContinue & "|") Then
			ControllerAutoDiscovery = "false"
		Else
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'yes' or 'no'."
			WScript.Quit
		End If
	ElseIf autoDiscoIIS = "true" Then
		ControllerAutoDiscovery = autoDiscoIIS
		If autoRestIIS = "true" Then
			iisRestart = "yes"
		ElseIf autoRestIIS = "false" Then
			iisRestart = "no"
		End If
	ElseIf autoDiscoIIS = "false" Then
		ControllerAutoDiscovery = autoDiscoIIS
	End If

End Function

Function setStandaloneDotNetServiceMon()

	If IsEmpty(autoStandMon) Then
		WScript.StdOut.Write "Monitor .Net Standalone applications (yes/no)?"
		PromptContinue = WScript.StdIn.ReadLine
		WScript.Echo ""
		If InStr("|yes|y|YES|Y|", "|" & PromptContinue & "|") Then
			MonitorStandaloneDotNet = "yes"
			WScript.Echo ""
			WScript.Echo "Available .Net Standalone applications:"
			Set cmd = wshShell.Exec("cmd.exe /C tasklist /m ""mscor*"" | findstr /BVI ""image* ===* appdynamics* sql* servermanager*")
			WaitForWshShellExec()
			Results = cmd.StdOut.ReadAll
			RegExp.Global = True
			RegExp.Pattern = "(\d+)"
			Set Matches = RegExp.Execute(Results)
			For each Match in Matches
   		 		ProcessId = Match
    			Set cmd = wshShell.Exec("cmd.exe /C tasklist /FI ""PID eq " & ProcessId & """ /FO list | find ""Image Name""")
    			WaitForWshShellExec()
    			Result = cmd.StdOut.ReadAll
    			Result = Right(Result, LEN(Result) - 14)
    			a = Split(Result, vbNewLine)
    			ProcessName = a(0)
    			WScript.stdout.write ProcessName
    			WScript.stdout.write " | "
			Next
			WScript.Echo ""
			WScript.Echo ""
			WScript.Echo "Please provide service/tier (service1.exe:TierName,Service2.exe:TierName2)."
			WScript.StdOut.Write "List: "
			PromptContinue = WScript.StdIn.ReadLine
			arrayStandaloneDotNetApps = Split(PromptContinue, ",")
			WScript.Echo ""
		ElseIf InStr("|no|n|NO|N|", "|" & PromptContinue & "|") Then
			MonitorStandaloneDotNet = "no"
		Else
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'yes' or 'no'."
			WScript.Quit
		End If
	ElseIf autoStandMon = "true" Then
		MonitorStandaloneDotNet = "yes"
		arrayStandaloneDotNetApps = Split(autoStandOpts, ",")
	ElseIf autoStandMon = "false" Then
		MonitorStandaloneDotNet = "no"
	End If

End Function

Function setAppTierName()

	' Support for autodeploy option
	If IsEmpty(autoTierName) Then
		WScript.StdOut.Write "Tier name: "
		PromptContinue = WScript.StdIn.ReadLine
		appTierName = PromptContinue
	Else
		appTierName = autoTierName
	End If
	If (appTierName = "") Or InStr(appTierName, " ") Then
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: name must not contain any spaces."
		WScript.Quit
	End If

End Function

Function setAppNodeName()

	' Support for autodeploy option
	If IsEmpty(autoNodeName) Then
		WScript.StdOut.Write "Node name: "
		PromptContinue = WScript.StdIn.ReadLine
		appNodeName = PromptContinue
	Else
		appNodeName = autoNodeName
	End If
	If (appNodeName = "") Or InStr(appNodeName, " ") Then
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: name must not contain any spaces."
	WScript.Quit
	End If

End Function

Function setSAASController()

	' Support for autodeploy option
	If IsEmpty(autoProdMon) Then
		WScript.Echo ""
		WScript.StdOut.Write "Production monitoring (yes/no)? "
		PromptContinue = WScript.StdIn.ReadLine
		If InStr("|yes|y|YES|Y|", "|" & PromptContinue & "|") Then
			ControllerHost = ControllerHostProd
			ControllerPort = ControllerHostPortProd
			ControllerHostProtocol = ControllerHostProtocolProd
			ControllerAccountName = ControllerUrlProdAccount
			ApiAccessToken = ControllerTokenApiProd
			ControllerAccessKey = ControllerUrlProdAccessKey
			url = ControllerUrlProd
		ElseIf InStr("|no|n|NO|N|", "|" & PromptContinue & "|") Then
			ControllerHost = ControllerHostNonprod
			ControllerPort = ControllerHostPortNonprod
			ControllerHostProtocol = ControllerHostProtocolNonprod
			ControllerAccountName = ControllerUrlNonprodAccount
			ApiAccessToken = ControllerTokenApiNonprod
			ControllerAccessKey = ControllerUrlNonprodAccessKey
			url = ControllerUrlNonprod
		Else
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'yes' or 'no'."
			WScript.Quit
		End If
	ElseIf autoProdMon = "true" Then
		ControllerHost = ControllerHostNonprod
		ControllerPort = ControllerHostPortNonprod
		ControllerHostProtocol = ControllerHostProtocolNonprod
		ControllerAccountName = ControllerUrlNonprodAccount
		ApiAccessToken = ControllerTokenApiNonprod
		ControllerAccessKey = ControllerUrlNonprodAccessKey
		url = ControllerUrlNonprod
	ElseIf autoProdMon = "false" Then
		ControllerHost = ControllerHostNonprod
		ControllerPort = ControllerHostPortNonprod
		ControllerHostProtocol = ControllerHostProtocolNonprod
		ControllerAccountName = ControllerUrlNonprodAccount
		ApiAccessToken = ControllerTokenApiNonprod
		ControllerAccessKey = ControllerUrlNonprodAccessKey
		url = ControllerUrlNonprod
	End If

	' Set Controller SSL based on protocol
	If (ControllerHostProtocol = "http") Then
		ControllerSSL = "false"
	ElseIf (ControllerHostProtocol = "https") Then
		ControllerSSL = "true"
	End If

End Function

Function validateWindowsVersion()

	' 4.x versions:		98 (4.10), ME (4.9)
	' 5.x versions:		2000 (5.0), XP (5.1), XP 64-bit (5.2), Server 2003 (5.2), Server 2003 R2 (5.2)
	' 6.x versions:		Vista (6.0), Server 2008 (6.0), Server 2008 R2 (6.1), 7 (6.1), Server 2012 (6.2), 8 (6.2), Server 2012 R2 (6.3), 8.1 (6.3)
	' 10.x versions:	Windows Server 2016 / 10: 10.0 / 10.0
	If WinReleaseVer < 6 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - " & WinOS & " is not supported."
		WScript.Quit
	ElseIf (WinReleaseVer >= 6.0) And (WinReleaseVer <= 10.0) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - " & WinOS & " " & "(" & WinArch & ", v" & WinReleaseVer & ")"
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a supported OS version (" & WinReleaseVer & ")"
		WScript.Quit
	End If

End Function

Function validateKbInstalled2k3()

	If WinReleaseVer < 6 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if KB948963 is deployed."
		Set Result = objService.ExecQuery("select HotFixID from Win32_QuickFixEngineering where HotFixID = ""KB948963""")
		WinHfInstalled = Result.Count
		If WinHfInstalled = 1 Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - KB948963 is deployed."
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - KB948963 is required before install can continue."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please install KB948963 and run script again."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - https://support.microsoft.com/en-us/help/948963/an-update-is-available-to-add-support-for-the-tls-rsa-with-aes-128-cbc"
			WScript.Quit
		End If
	End If

End Function

Function setJavaAgentType()

	' Support for autodeploy option
	If IsEmpty(autoAppJVM) Then
		WScript.StdOut.Write "Select JVM (Sun/IBM): "
		PromptContinue = WScript.StdIn.ReadLine
		JavaAgentType = PromptContinue
	Else
		JavaAgentType = autoAppJVM
	End If
	If JavaAgentType = "Sun" Then
		JavaAgentRootDir = "binaries\JavaInstall\Sun\"
		Set JavaAgentSourceDir = objFS0.GetFolder(JavaAgentRootDir)
		Set subFolders = JavaAgentSourceDir.SubFolders
		For Each subFolder in subFolders
			JavaAgentSourceDir = subFolder.Name
		Next
		JavaAgentSourceDir = objFS0.GetFolder(JavaAgentRootDir & JavaAgentSourceDir) 
	ElseIf JavaAgentType = "IBM" Then
		JavaAgentRootDir = "binaries\JavaInstall\IBM\"
		Set JavaAgentSourceDir = objFS0.GetFolder(JavaAgentRootDir)
		Set subFolders = JavaAgentSourceDir.SubFolders
		For Each subFolder in subFolders
			JavaAgentSourceDir = subFolder.Name
		Next
		JavaAgentSourceDir = objFS0.GetFolder(JavaAgentRootDir & JavaAgentSourceDir) 
	Else
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'Sun' or 'IBM'."
		WScript.Quit
	End If

End Function

Function setJavaAgentAppType()

	' Support for autodeploy option
	If IsEmpty(autoAppType) Then
		WScript.Echo ""
		WScript.Echo "Available applications for monitoring: "
		WScript.Echo "JBoss | Tomcat | WebSphereAS | WebLogic"
		WScript.Echo ""
		WScript.StdOut.Write "Application to monitor: "
		PromptContinue = WScript.StdIn.ReadLine
		CustomerApplication = PromptContinue
		WScript.Echo ""
	Else
		CustomerApplication = autoAppType
	End If
	If InStr("|JBoss|Tomcat|WebSphereAS|WebLogic|", "|" & CustomerApplication & "|") Then
	Else
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: choose from one of the existing options."
		WScript.Quit
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Selected monitoring for " & CustomerApplication & "."
	If CustomerApplication = "JBoss" Then
		' Support for autodeploy option
		If IsEmpty(autoAppStartScript) Then
			WScript.Echo ""
			WScript.StdOut.Write "Full path to application start-up script (ex: D:\jboss-as\bin\run.conf.bat): "
			PromptContinue = WScript.StdIn.ReadLine
			appStartUpScript = PromptContinue
		Else
			appStartUpScript = autoAppStartScript
		End If
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if start up script exists."
		If objFS0.FileExists(appStartUpScript) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script exists. Creating backup of file."
			objFS0.CopyFile appStartUpScript, backupDirJava & "\" & objFs0.GetFileName(appStartUpScript) & "-" & DateDiff("s", "01/01/1970 00:00:00", Now())
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found. Setup will quit."
			WScript.Quit
		End If
		Set objFile = objFS0.OpenTextFile(appStartUpScript, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.close
		For each Line in FileContents
			If (InStr(Line, "-Dappdynamics")) Or (InStr(Line, "-javaagent:")) Then
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script must contain no instrumentation settings."
				WScript.Quit
			Else
			End If
		Next
	ElseIf CustomerApplication = "Tomcat" Then
		
		' Select the right Tomcat instance for monitoring
		setTomcatJavaAgentMon()
		
		' Checking existing Tomcat configuration and backup of registry value
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Looking for Tomcat start up parameters."
		If WinArch = 32 Then
			strKeyPath = "HKLM\SOFTWARE\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java\Options"
		ElseIf WinArch = 64 Then
			strKeyPath = "HKLM\SOFTWARE\Wow6432Node\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java\Options"
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not determine Windows Server architecture (32/64)."
			WScript.Quit
		End If
		on error resume next
		wshShell.RegRead(strKeyPath)
		If err.number = 0 then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Found registry Key/Value for " & TomcatServiceName & "."
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Registry Key/Value not found  for " & TomcatServiceName & " (" & Err.Number & ")."
			WScript.Quit
		End If
		on error goto 0
		If WinArch = 32 Then
			strKeyPath = "SOFTWARE\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java"
		ElseIf WinArch = 64 Then
			strKeyPath = "SOFTWARE\Wow6432Node\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java"
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not determine Windows Server architecture (32/64)."
			WScript.Quit
		End If
		ValueName = "Options"
		Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,ValueName,ArrTomcatOptions
		For each ArrTomcatOptions in ArrTomcatOptions
			If (InStr(ArrTomcatOptions, "-javaagent")) Or (InStr(ArrTomcatOptions, "-Dappdynamics")) Then 
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up parameters must contain no instrumentation settings."
				WScript.Quit
			Else
			End If
		Next
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Creating registry key backup."
		WshShell.Run("cmd.exe /C regedit /E " & backupDirJava & "\tomcatJavaOptions" & "-" & DateDiff("s", "01/01/1970 00:00:00", Now()) & ".reg " & """" & "HKEY_LOCAL_MACHINE\" & strKeyPath & """"), 1, True
	ElseIf CustomerApplication = "WebSphereAS" Then
		
		' Select the right WebSphereAS for monitoring, backup files and verify
		setWebsphereAppServerJavaAgentMon()

	ElseIf CustomerApplication = "WebLogic" Then
		
		' Select the right WebLogic instance for monitoring, backup files and verify
		setWebLogicJavaAgentMon()
		
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a valid application selection."
		WScript.Quit
	End If

End Function

Function setWebLogicJavaAgentMon()

	' Support for autodeploy option

	If IsEmpty(autoStartWebLogic) Then
		
		' Scan registry for WebLogic services
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Scanning for available WebLogic services."
		Set cmd = wshShell.Exec("cmd.exe /C reg query HKLM\SYSTEM\CurrentControlSet\services /D /E /F WebLogicServer /S | find ""services""")
		Results = cmd.StdOut.ReadAll
		length = Len(Results)
		
		' Select instrumentation method
		If length < 52 Then

			' WEBLOGIC START-UP PARAMETERS FROM DOMAIN START-UP SCRIPT

			' Locate start-up script based on domain selection
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Couldn't find existing WebLogic service."
			WScript.Echo ""
			WScript.StdOut.Write "Full path to domain-registry.xml (ex: C:\Oracle\Middleware\Oracle_Home\domain-registry.xml): "
			WebLogicDomainRegistryFile = WScript.StdIn.ReadLine
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if file exists."
			
			If objFS0.FileExists(WebLogicDomainRegistryFile) Then

				SET objFile = objFS0.GetFile(WebLogicDomainRegistryFile)

				If objFile.Size  > 0 Then
			
					WScript.Echo FormatDateTime(Now, vbLongTime) & " - Discovering WebLogic domains."
					WScript.Echo ""
					Set objFile = objFS0.OpenTextFile(WebLogicDomainRegistryFile, ForReading)
					FileContents = objFile.ReadAll
					objFile.close
					RegExp.Global = True
					RegExp.Pattern = "([a-zA-Z]:\\.*)""\/\>"
					domainDirs = Array()

					For each Match in RegExp.Execute(FileContents)

						ReDim Preserve domainDirs(UBound(domainDirs) + 1)
						domainDirs(UBound(domainDirs)) = Match.SubMatches(0)
					
					Next

					WScript.StdOut.Write "Available WebLogic domains: "
					WScript.Echo ""
					
					For Each Match in domainDirs

						domainName = Split(Match, "\")
						WScript.StdOut.Write domainName(UBound(domainName)) & " | "
					
					Next
					
					WScript.Echo ""
					WScript.Echo ""
					WScript.StdOut.Write "WebLogic domain to monitor: "
					domainName = WScript.StdIn.ReadLine
					WScript.Echo ""

					' Find the domain dir based on selected domain name
					For Each Match in domainDirs

						a = Split(Match, "\")
						strText = a(UBound(a))

						If strText = domainName Then

							domainDir = Match
							WScript.Echo FormatDateTime(Now, vbLongTime) & " - Using " & domainName & "."

						End If

					Next

					If objFS0.FolderExists(domainDir) Then
					
						WScript.Echo FormatDateTime(Now, vbLongTime) & " - Dir for domain " & domainName & " exists."
						startWebLogic = domainDir & "\bin\startWebLogic.cmd"

					End If
					
					If objFS0.FileExists(startWebLogic) Then

						SET objFile = objFS0.GetFile(startWebLogic)

						If objFile.Size  > 0 Then

							WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script exists for domain " & domainName & ". Creating backup of file."
							objFS0.CopyFile startWebLogic, backupDirJava & "\weblogic_startWebLogic.cmd_" & DateDiff("s", "01/01/1970 00:00:00", Now())

						Else

							WScript.Echo FormatDateTime(Now, vbLongTime) & " - File is empty. Setup will quit."
							WScript.Quit

						End If

					Else

						WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found. Setup will quit."
						WScript.Quit

					End If
			
				Else
			
					WScript.Echo FormatDateTime(Now, vbLongTime) & " - File is empty. Setup will quit."
					WScript.Quit
			
				End If

			Else

				WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found. Setup will quit."
				WScript.Quit

			End If

			' Check for existing instrumentation
			Set objFile = objFS0.OpenTextFile(startWebLogic, ForReading)
			FileContents = Split(objFile.ReadAll, vbNewLine)
			objFile.close
			
			For each Line in FileContents
			
				If (InStr(Line, "-Dappdynamics")) Or (InStr(Line, "-javaagent:")) Then
					
					WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
					WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script must contain no instrumentation settings."
					WScript.Quit
			
				Else
			
				End If
			
			Next

		Else
			
			' WEBLOGIC START-UP PARAMETERS FROM WINDOWS SERVICE REGISTRY
			
			' Auto discover start-up script for Windows Service based deployment
			WScript.Echo ""
			RegExp.Global = True
			RegExp.Pattern = "\\services\\(.*)"
			For Each Match in RegExp.Execute(Results)
				WScript.stdout.write Replace(Match.SubMatches(0), VbCr, " | ")
			Next
			WScript.Echo ""
			WScript.Echo ""
			WScript.StdOut.Write "WebLogic service to monitor: "
			WebLogicServiceName = WScript.StdIn.ReadLine
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Using """ & WebLogicServiceName & """."
			
			' Locate domain dir
			Set cmd = wshShell.Exec("cmd.exe /C reg query ""HKLM\SYSTEM\CurrentControlSet\services\" & WebLogicServiceName & "\Parameters"" /v ExecDir | find ""domains""")
			domainDir = cmd.StdOut.ReadAll
			RegExp.Global = True
			RegExp.Pattern = "([a-zA-Z]:\\.*)"
			' Match the regex group
			For Each Match in RegExp.Execute(domainDir)
				domainDir = Replace(Match.SubMatches(0), VbCr, "\")
			Next

			If objFS0.FolderExists(domainDir) Then

				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Going to backup """ & WebLogicServiceName & """ service registry."
				Set cmd = wshShell.Exec("cmd.exe /C reg export ""HKLM\SYSTEM\CurrentControlSet\services\" & WebLogicServiceName & """ """ & backupDirJava & "\weblogic_" & WebLogicServiceName & """_" & DateDiff("s", "01/01/1970 00:00:00", Now()) & ".reg")
				
			Else
				
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not find WebLogic domain directory."
				WScript.Quit
			
			End If

			' Retrieve domain start-up parameters
			Set cmd = wshShell.Exec("cmd.exe /C reg query ""HKLM\SYSTEM\CurrentControlSet\services\" & WebLogicServiceName & "\Parameters"" /v CmdLine | find ""CmdLine""")
			CmdLine = cmd.StdOut.ReadAll
			RegExp.Global = True
			RegExp.Pattern = "(-server.*)"
			For Each Match in RegExp.Execute(CmdLine)
				CmdLine = Match.SubMatches(0)
				CmdLine = Replace(CmdLine, VbCr, "")
			Next
			
			   
			If InStr(CmdLine, "weblogic.Server") Then
				
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Validated domain CmdLine reg_sz."
			
			Else
			
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not validate CmdLine reg_sz."
				WScript.Quit
			
			End If

			' Check for existing instrumentation
			If (InStr(CmdLine, "-Dappdynamics")) And (InStr(CmdLine, "-javaagent:")) Then

				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up params must contain no instrumentation settings."
				WScript.Quit

			Else
			End If

		End If

		' RETRIEVE WEBLOGIC POLICY FILE
		
		' Check if weblogic.policy exists and make backup
		RegExp.Global = True
		RegExp.Pattern = "(user_projects.*)"
		For Each Match in RegExp.Execute(domainDir)
			strText = Match.SubMatches(0)
		Next
		WebLogicPolicyFile = Replace(domainDir, strText, "wlserver\server\lib\weblogic.policy")
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Looking for weblogic.policy file."

		If objFS0.FileExists(WebLogicPolicyFile) Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - File exists. Creating backup."
			objFS0.CopyFile WebLogicPolicyFile, backupDirJava & "\" & objFs0.GetFileName(WebLogicPolicyFile) & "_" & DateDiff("s", "01/01/1970 00:00:00", Now())

		Else
			
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Follow instructions in the end to manually update the policy file."
			WebLogicPolicyFile = "false"
		
		End If

	Else

		' Handle autodeploy for WebLogic start-up types
		' finish
	
	End If

End Function

Function instrumentWebLogicJavaAgent()

	If IsEmpty(CmdLine) Then

		' Weblogic startWebLogic instrumentation
		Set objFile = objFS0.OpenTextFile(startWebLogic, ForReading)
		FileContents = objFile.ReadAll
		objFile.close
		
		If InStr(FileContents, "@REM START WEBLOGIC") Then
			
			If proxyEnabled = "true" Then
				
				strText = "@REM Enable the AppDynamics Java Agent" & vbNewLine & "set JAVA_OPTIONS=%JAVA_OPTIONS% -javaagent:" & JavaAgentHomeDir & "\javaagent.jar" & " -Dappdynamics.agent.applicationName=" & ControllerApplicationName & " -Dappdynamics.agent.tierName=" & appTierName & " -Dappdynamics.agent.nodeName=" & appNodeName & " -Dappdynamics.http.proxyhost=" & proxyHost & " -Dappdynamics.http.proxyPort=" & proxyPort & vbNewLine & "@REM AppDynamics Java Agent END" & vbNewLine & "@REM START WEBLOGIC"

			ElseIf proxyEnabled = "false" Then

				strText = "@REM Enable the AppDynamics Java Agent" & vbNewLine & "set JAVA_OPTIONS=%JAVA_OPTIONS% -javaagent:" & JavaAgentHomeDir & "\javaagent.jar" & " -Dappdynamics.agent.applicationName=" & ControllerApplicationName & " -Dappdynamics.agent.tierName=" & appTierName & " -Dappdynamics.agent.nodeName=" & appNodeName & vbNewLine & "@REM AppDynamics Java Agent END" & vbNewLine & "@REM START WEBLOGIC"

			End If
			
			FileContents = Replace(FileContents, "@REM START WEBLOGIC", strText)
			Set objFile = objFs0.OpenTextFile(startWebLogic, ForWriting)
			objFile.Write FileContents
			objFile.close

		Else
			
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not instrument start-up script."
			objFile.close
			WScript.Quit
		
		End If
		
		' Verify instrumentation of startWebLogic.cmd
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		Set objFile = objFS0.OpenTextFile(startWebLogic, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.close
		
		For each Line in FileContents
		
			If (InStr(Line, "-Dappdynamics")) And (InStr(Line, "-javaagent:")) Then
		
				IsInstrumented = "1"
		
			Else
			End If
		
		Next

	ElseIf InStr(CmdLine, "weblogic.Server") Then

		If proxyEnabled = "true" Then

			strText = "-javaagent:""" & JavaAgentHomeDir & "\javaagent.jar"" -Dappdynamics.agent.applicationName=""" & ControllerApplicationName & """ -Dappdynamics.agent.tierName=""" & appTierName & """ -Dappdynamics.agent.nodeName=""" & appNodeName & """ -Dappdynamics.http.proxyhost=" & proxyHost & " -Dappdynamics.http.proxyPort=" & proxyPort & " weblogic.Server"

		ElseIf proxyEnabled = "false" Then

			strText = "-javaagent:""" & JavaAgentHomeDir & "\javaagent.jar"" -Dappdynamics.agent.applicationName=""" & ControllerApplicationName & """ -Dappdynamics.agent.tierName=""" & appTierName & """ -Dappdynamics.agent.nodeName=""" & appNodeName & """ weblogic.Server"

		End If

		CmdLine = Replace(CmdLine, "weblogic.Server", strText)
		Set cmd = wshShell.Exec("cmd.exe /C reg add ""HKLM\SYSTEM\CurrentControlSet\services\" & WebLogicServiceName & "\Parameters"" /v CmdLine /t REG_SZ /f /d """ & CmdLine & """")
		Result = cmd.StdOut.ReadAll
		RegExp.Global = True
		RegExp.Pattern = "(completed successfully)"
		
		' Match the regex group
		For Each Match in RegExp.Execute(Result)
		
			Result = Match.SubMatches(0)
		
		Next

		If Result = "completed successfully" Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - CmdLine reg_sz modified."
		
		Else

			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to update CmdLine reg_sz."
			WScript.Quit

		End If

		' Verify instrumentation of CmdLine reg_sz
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."

		Set cmd = wshShell.Exec("cmd.exe /C reg query ""HKLM\SYSTEM\CurrentControlSet\services\" & WebLogicServiceName & "\Parameters"" /v CmdLine | find ""CmdLine""")
		CmdLine = cmd.StdOut.ReadAll
		RegExp.Global = True
		RegExp.Pattern = "(javaagent.jar)"
		For Each Match in RegExp.Execute(CmdLine)
			
			Result = Match.SubMatches(0)

			If Result = "javaagent.jar" Then
			
				IsInstrumented = "1"
			
			End If		

		Next

	End If

	' UPDATE WEBLOGIC.POLICY FILE

	If IsInstrumented = 1 Then

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."

	Else

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
		WScript.Quit

	End If

	If InStr(WebLogicPolicyFile, "false") Then

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Add the following to your weblogic.policy file."
		WScript.Echo ""
		'WScript.Echo "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentHomeDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
		WScript.Echo "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentInstallDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
		WScript.Echo ""

	Else

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Going to update weblogic.policy file."
		' Check if there is existing policy configuration
		Set objFile = objFS0.OpenTextFile(WebLogicPolicyFile, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.close

		For each Line in FileContents
		
			If InStr(Line, "AppDynamics") Then
		
				IsWebLogicPolicyFileConfigured = 1
		
			End If
		
		Next
		
		' Update weblogic.policy
		If IsWebLogicPolicyFileConfigured = 1 Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - File already contains AppDynamics reference."
		
		Else
	
			Set objFile = objFS0.OpenTextFile(WebLogicPolicyFile, ForReading)
			strText = objFile.ReadAll
			objFile.Close
			'strText = strText & vbNewLine & "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentHomeDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
			strText = strText & vbNewLine & "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentInstallDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
			Set objFile = objFS0.OpenTextFile(WebLogicPolicyFile, ForWriting)
			objFile.Write strText
			objFile.Close
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - weblogic.policy updated."
	
		End If

	End If

	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."

End Function

Function setTomcatJavaAgentMon()

	' Support for autodeploy option
	If IsEmpty(autoTomcatServiceName) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Scanning for available Tomcat instances."
		Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		If WinArch = 32 Then
			strKeyPath = "SOFTWARE\Apache Software Foundation\Procrun 2.0"
		ElseIf WinArch = 64 Then
			strKeyPath = "SOFTWARE\Wow6432Node\Apache Software Foundation\Procrun 2.0"
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not determine Windows Server architecture (32/64)."
			WScript.Quit
		End If
		on error resume next
		wshShell.RegRead("HKEY_LOCAL_MACHINE\" & strKeyPath & "\")
		If err.number = 0 then
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Couldn't find existing Tomcat settings in registry."
			WScript.Quit
		End If
		WScript.Echo ""
		on error goto 0
		oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
		searchString = "Tomcat"
		For each arrSubKeys in arrSubKeys
    		WScript.stdout.write arrSubKeys
   		 	WScript.Stdout.write " | "
		Next
		WScript.Echo ""
		WScript.Echo ""
		WScript.StdOut.Write "Tomcat service to monitor: "
		PromptContinue = WScript.StdIn.ReadLine
		TomcatServiceName = PromptContinue
		WScript.Echo ""
	Else
		TomcatServiceName = autoTomcatServiceName
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Using " & TomcatServiceName & "."

End Function

Function setWebsphereAppServerJavaAgentMon()

	' Support for autodeploy option
	If IsEmpty(autoAppConf) Then
		' Get path to server.policy from the user
		WScript.Echo ""
		WScript.StdOut.Write "Full path to server.policy (ex: C:\IBM\WebSphere\AppServer\profiles\myapp\properties\server.policy): "
		WebSphereASPolicyFile = WScript.StdIn.ReadLine
		WScript.Echo ""
	Else
		WebSphereASPolicyFile = autoAppConf
	End If

	' Check if server.policy exists and make backup
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if file exists."
	If objFS0.FileExists(WebSphereASPolicyFile) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - File exists. Creating backup."
		objFS0.CopyFile WebSphereASPolicyFile, backupDirJava & "\" & objFs0.GetFileName(WebSphereASPolicyFile) & "-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found. Setup will quit."
		WScript.Quit
	End If

	' Support for autodeploy option
	If IsEmpty(autoAppStartScript) Then
		' Get path to server.xml from the user
		WScript.Echo ""
		WScript.StdOut.Write "Full path to server.xml (ex: C:\IBM\WebSphere\AppServer\profiles\myapp\config\cells\mycell\nodes\mynode\servers\myserver\server.xml): "
		WebSphereASServerXML = WScript.StdIn.ReadLine
	Else
		WebSphereASServerXML = autoAppStartScript
	End If
	WScript.Echo ""

	' Check if server.xml exists and make backup
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if file exists."
	If objFS0.FileExists(WebSphereASServerXML) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - File exists. Creating backup."
		objFS0.CopyFile WebSphereASServerXML, backupDirJava & "\" & objFs0.GetFileName(WebSphereASServerXML) & "-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found. Setup will quit."
		WScript.Quit
	End If

	' Check if statisticSet property exists in server.xml
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.Async = "false"
	xmlDoc.Load(WebSphereASServerXML)
	Set xmlDocPMIService = xmlDoc.selectSingleNode(".//node()[@xmi:type = 'pmiservice:PMIService']")
	If xmlDocPMIService is Nothing Then
		WScript.Echo FormatDateTime(Now, vbLongTime) * " - Couldn't find pmiservice:PMIService config in server.xml."
		WScript.Quit
	Else
		WebSphereASstatisticSet = xmlDocPMIService.Attributes.getNamedItem("statisticSet").Text
	End If
	Set xmlDoc = Nothing

	' Check if there is existing instrumentation
	Set objFile = objFS0.OpenTextFile(WebSphereASServerXML, ForReading)
	FileContents = Split(objFile.ReadAll, vbNewLine)
	objFile.close
	For each Line in FileContents
		If (InStr(Line, "-Dappdynamics")) Or (InStr(Line, "-javaagent:")) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script must contain no instrumentation settings."
			WScript.Quit
		Else
		End If
	Next
	
	' Check if there is existing policy configuration
	Set objFile = objFS0.OpenTextFile(WebSphereASPolicyFile, ForReading)
	FileContents = Split(objFile.ReadAll, vbNewLine)
	objFile.close
	For each Line in FileContents
		If InStr(Line, "AppDynamics") Then
			IsWebSphereASPolicyFileConfigured = 1
		End If
	Next

End Function

Function validateDotNetFrameworkVersion()

	If WinArch = 32 Then
		strKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components"
	ElseIf WinArch = 64 Then
		strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components"
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not determine Windows Server architecture (32/64)."
		WScript.Quit
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if .Net Framework 2 or later is deployed."
	Set objectReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	objectReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	searchString = ".NETFramework"
	For Each subkey In arrSubKeys
		On Error Resume Next
		regComponentID = WshShell.Regread("HKLM\" & strKeyPath & "\" & subkey & "\ComponentID")
		if err.number = 0 then 
			On Error Goto 0
			if instr(regComponentID, searchString) > 0 then
				On Error Resume Next
				regDotNetVersion = WshShell.Regread("HKLM\" & strKeyPath & "\" & subkey & "\Version")
            	On Error GoTo 0
			end if
		end if
	Next
	On Error Goto 0
	If Left(regDotNetVersion, 1) >= 2 Then
   	 WScript.Echo FormatDateTime(Now, vbLongTime) & " - .Net Framework check succeeded (" & regDotNetVersion & ")."
	Else
   	 WScript.Echo FormatDateTime(Now, vbLongTime) & " - .Net Framework check failed. Detected version " & regDotNetVersion & "."
   	 WScript.Quit
	End if

End Function

Function setSimEnabled()

	If IsEmpty(autoSimEnabled) Then
		WScript.Echo ""
		WScript.StdOut.Write "Advanced infrastructure metrics (yes/no)?: "
		PromptContinue = WScript.StdIn.ReadLine
	Else
		PromptContinue = autoSimEnabled
	End If
	If InStr("|yes|y|YES|Y|true|", "|" & PromptContinue & "|") Then
		simEnabled = "true"
	ElseIf InStr("|no|n|NO|N|false|", "|" & PromptContinue & "|") Then
		simEnabled = "false"
	Else
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'yes' or 'no'."
		WScript.Quit
	End If

End Function

Function PrintUsage()

	WScript.Echo "Usage: {type} {item}"
	WScript.Echo "       {type}: install | upgrade | uninstall | instrument | autodeploy | healthcheck"
	WScript.Echo "       {item}: -dna | -ja | -ma | -prod | -nonprod"
	WScript.Echo ""
	WScript.Echo "Available options (current release):"
	WScript.Echo "       'cscript deployAgent.vbs autodeploy' auto deploy one/more of the below."
	WScript.Echo "       'cscript deployAgent.vbs install -dna' installs a Dot Net Agent."
	WScript.Echo "       'cscript deployAgent.vbs install -ja' installs a Java Agent."
	WScript.Echo "       'cscript deployAgent.vbs install -ma' installs a Machine Agent."
	WScript.Echo "       'cscript deployAgent.vbs upgrade -dna' upgrades a Dot Net Agent."
	WScript.Echo "       'cscript deployAgent.vbs upgrade -ja' upgrades a Java Agent."
	WScript.Echo "       'cscript deployAgent.vbs upgrades -ma' upgrades a Machine Agent."	
	WScript.Echo "       'cscript deployAgent.vbs uninstall -dna' uninstalls a Dot Net Agent."
	WScript.Echo "       'cscript deployAgent.vbs uninstall -ja' uninstalls a Java Agent."
	WScript.Echo "       'cscript deployAgent.vbs uninstall -ma' uninstalls a Machine Agent."
	WScript.Echo "       'cscript deployAgent.vbs instrument -ja' instruments an additional JVM."
	WScript.Echo "       'cscript deployAgent.vbs healthcheck -prod' check connectivity for prod."
	WScript.Echo "       'cscript deployAgent.vbs healthcheck -nonprod' for nonprod."
	WScript.Echo " For more information please refer to ReadMe.txt"

End Function

Function Install_dna()

	' Echo Windows Server environment information
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - " & WinOS & " " & "(" & WinArch & ", v" & WinReleaseVer & ")"
	
	' Check if .Net Framework exists
	call validateDotNetFrameworkVersion()

	' Check for KB2999226 (Windows 6.x)
	If (WinReleaseVer >= 6.0) And (WinVer <= 2012) Then
		Set cmd = wshShell.Exec("cmd.exe /C systeminfo.exe | find /c ""KB2999226""")
		WaitForWshShellExec()
		WinHfInstalled = (cmd.StdOut.ReadLine())
		If WinHfInstalled = 1 Then
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - KB2999226 is required before install can continue."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please install KB2999226 and run script again."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - https://support.microsoft.com/en-us/help/2999226/"
			WScript.Quit
		End If
	End If

	' Check for KB948963 (Windows 5.x)
	call validateKbInstalled2k3()

	' Check for existing installation
	On Error Resume Next
	Err.Clear
	registryInstallationDir = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AppDynamics\dotNet Agent\InstallationDir")
	If Err.Number <> 0 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No existing agent installation detected."
		Err.Clear
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Detected existing install."
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Uninstall existing setup or run script with option to update."
		WScript.Quit
	End If
	On Error Goto 0
	
	' Check Windows version and select the right MSI package/conf
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Running agent binaries compatibility check."
	If WinReleaseVer < 6 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Going to use an older agent version for OS compatibility"
		DotNetInstallConf = "conf\dotnet\compatibility\config.xml"
		DotNetInstallConf = objFS0.GetFile(DotNetInstallConf)
		DotNetInstallConfTempl = "conf\dotnet\compatibility\config-template.xml"
		DotNetInstallConfTempl = objFS0.GetFile(DotNetInstallConfTempl)
		If WinArch = 32 Then
			objStartFolder = "binaries\DotNetInstall\compatibility\x86"
		ElseIf WinArch = 64 Then
			objStartFolder = "binaries\DotNetInstall\compatibility\x64"
		End If
		Set objFolder = objFS0.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			DotNetInstall = objFolder & "\" & objFile.name
		Next	
	ElseIf WinReleaseVer >= 6 Then
		DotNetInstallConf = "conf\dotnet\new\AD_config.xml"
		DotNetInstallConf = objFS0.GetFile(DotNetInstallConf)
		DotNetInstallConfTempl = "conf\dotnet\new\AD_config-template.xml"
		DotNetInstallConfTempl = objFS0.GetFile(DotNetInstallConfTempl)
		If WinArch = 32 Then
			objStartFolder = "binaries\DotNetInstall\new\x86"
		ElseIf WinArch = 64 Then
			objStartFolder = "binaries\DotNetInstall\new\x64"
		End If
		Set objFolder = objFS0.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			DotNetInstall = objFolder & "\" & objFile.name
		Next
	End If
	DotNetInstallLog = "DotNetAgentInstall.log"
	
	' Select the type of Controller
	Call setSAASController()

	' Check connection to Controller
	Call getSAASControllerDetails()

	' Select application and key based on a text file list
	Call setAppName()

	' Select if IIS .Net application monitoring is required
	Call setIISDotNetMon()

	' Select if .Net standalone application monitoring is required
	Call setStandaloneDotNetServiceMon()

	' Update config.xml used for installation based on user input
	If WinReleaseVer < 6 Then
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = "false"
			xmlDoc.Load(DotNetInstallConf)
		Set xmlDocAppDynamicsAgent = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent")
		Set xmlDocControllerInfo = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent/controller")
		Set xmlDocControllerProxy = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent/controller/proxy")
		Set xmlDocAccountInfo = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent/controller/account")
		Set xmlDocApplicationInfo = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent/controller/application")
		Set xmlDocAutoDiscoveryInfo = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent/app-agents/IIS/automatic")
		If (MonitorStandaloneDotNet = "yes") And (NOT (IsEmpty(arrayStandaloneDotNetApps))) Then
			For Each arrayStandaloneDotNetApps in arrayStandaloneDotNetApps
				a = Split(arrayStandaloneDotNetApps, ":")
				StandaloneDotNetAppName = a(0)
				StandaloneDotNetAppTierName = a(1)
				Set xmlDocStandaloneApplications = xmlDoc.DocumentElement.selectSingleNode("//appdynamics-agent/app-agents/standalone-applications")
				Set xmlDocStandaloneApp= _
					xmlDoc.createElement("standalone-application")
					xmlDocStandaloneApplications.appendChild xmlDocStandaloneApp
				Set xmlDocStandaloneAppAttr= _
					xmlDoc.createAttribute("executable")
					xmlDocStandaloneApp.setAttribute "executable", StandaloneDotNetAppName
				Set xmlDocStandaloneAppTier= _
					xmlDoc.createElement("tier")
					xmlDocStandaloneApp.appendChild xmlDocStandaloneAppTier
				Set xmlDocStandaloneAppTierAttr= _
					xmlDoc.createAttribute("name")
					xmlDocStandaloneAppTier.setAttribute "name", StandaloneDotNetAppTierName
			Next
		End If
	ElseIf WinReleaseVer >= 6 Then
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = "false"
			xmlDoc.Load(DotNetInstallConf)
		Set xmlDocStandaloneApplications = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/app-agents/standalone-applications")
		Set xmlDocControllerInfo = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/controller")
		Set xmlDocControllerProxy = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/controller/proxy")
		Set xmlDocAccountInfo = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/controller/account")
		Set xmlDocApplicationInfo = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/controller/application")
		Set xmlDocAutoDiscoveryInfo = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/app-agents/IIS/automatic")
		If (MonitorStandaloneDotNet = "yes") And (NOT (IsEmpty(arrayStandaloneDotNetApps))) Then
			For Each arrayStandaloneDotNetApps in arrayStandaloneDotNetApps
				a = Split(arrayStandaloneDotNetApps, ":")
				StandaloneDotNetAppName = a(0)
				StandaloneDotNetAppTierName = a(1)
				Set xmlDocStandaloneApplications = xmlDoc.DocumentElement.selectSingleNode("//winston/appdynamics-agent/app-agents/standalone-applications")
				Set xmlDocStandaloneApp= _
					xmlDoc.createElement("standalone-application")
					xmlDocStandaloneApplications.appendChild xmlDocStandaloneApp
				Set xmlDocStandaloneAppAttr= _
					xmlDoc.createAttribute("executable")
					xmlDocStandaloneApp.setAttribute "executable", StandaloneDotNetAppName
				Set xmlDocStandaloneAppTier= _
					xmlDoc.createElement("tier")
					xmlDocStandaloneApp.appendChild xmlDocStandaloneAppTier
				Set xmlDocStandaloneAppTierAttr= _
					xmlDoc.createAttribute("name")
					xmlDocStandaloneAppTier.setAttribute "name", StandaloneDotNetAppTierName
			Next
		End If
	End If
	xmlDocControllerInfo.Attributes.getNamedItem("host").Text = ControllerHost
	xmlDocAccountInfo.Attributes.getNamedItem("name").Text = ControllerAccountName
	xmlDocControllerProxy.Attributes.getNamedItem("host").Text = proxyHost
	xmlDocControllerProxy.Attributes.getNamedItem("port").Text = proxyPort
	xmlDocControllerProxy.Attributes.getNamedItem("enabled").Text = proxyEnabled
	xmlDocAccountInfo.Attributes.getNamedItem("password").Text = ControllerAccessKey
	xmlDocApplicationInfo.Attributes.getNamedItem("name").Text = ControllerApplicationName
	xmlDocAutoDiscoveryInfo.Attributes.getNamedItem("enabled").Text = ControllerAutoDiscovery
	xmlDoc.save(DotNetInstallConf)
	
	' Installation and start up
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - updated config.xml. Ready for install."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Config file: " & DotNetInstallConf
	DotNetInstallDir = DiskDrive & "\AppDynamics\DotNetAgent\AppDynamics .NET Agent\"
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Install dir: " & DotNetInstallDir
	DotNetInstallDataDir = DiskDrive & "\AppDynamics\DotNetAgent\AppDynamics .Net Agent Data\"
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Data dir: " & DotNetInstallDataDir
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Starting install..."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Validating install folder."
	If objFS0.FolderExists(DiskDrive & "\AppDynamics\DotNetAgent\") Then
		Set objFolder = objFS0.GetFolder(DiskDrive & "\AppDynamics\DotNetAgent\")
		If (objFolder.Files.Count > 0) Or (ObjFolder.SubFolders.Count > 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder already exists and is not empty. Setup will quit."
			WScript.Quit
		End If
	Else
	End If
	If WinReleaseVer < 6 Then
		WshShell.Run ("msiexec " & "/i " & DotNetInstall & " /q" & " /norestart" & " /lv " & Chr(34) & DotNetInstallLog & Chr(34) & " INSTALLDIR=" & Chr(34) & DotNetInstallDir & Chr(34) & " DOTNETAGENTFOLDER=" & Chr(34) & DotNetInstallDataDir & Chr(34)), 0, True
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Copying config.xml"
		objFS0.CopyFile DotNetInstallConf, DotNetInstallDataDir & "Config\config.xml", True
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Starting AppDynamics .NET Agent service."
		wshShell.Run("cmd.exe /C net start AppDynamics.Agent.Coordinator_service"), 0, True
	ElseIf WinReleaseVer >= 6 Then
		WshShell.Run("msiexec " & "/i " & DotNetInstall & " /q" & " /norestart" & " /lv " & Chr(34) & DotNetInstallLog & Chr(34) & " AD_SetupFile=" & Chr(34) & DotNetInstallConf & Chr(34) & " INSTALLDIR=" & Chr(34) & DotNetInstallDir & Chr(34) & " DOTNETAGENTFOLDER=" & Chr(34) & DotNetInstallDataDir & Chr(34)), 0, True
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Install finished. Check " & DotNetInstallLog & " for more details."
	If iisRestart = "yes" Then
		Set cmd = wshShell.Exec("cmd.exe /C sc query AppDynamics.Agent.Coordinator_service | find ""STATE""")
		WaitForWshShellExec()
		DotNetAgentServiceState = Right((cmd.StdOut.ReadLine()), 8)
		If DotNetAgentServiceState = "RUNNING " Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Going to restart IIS using iisreset command."
			wshShell.Run("cmd.exe /C iisreset"), 0, True
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Dot Net Agent service is not running."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please ensure agent service is running and restart IIS manually."
		End If
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No IIS restart (based on user selection)."
	End If
	objFs0.CopyFile DotNetInstallConfTempl, DotNetInstallConf, True

End Function

Function Install_ja()

	' Backup dir
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Backup dir is " & backupDirJava & "."

	' Check for KB948963 (Windows 5.x)
	call validateKbInstalled2k3()

	' Select the type of Controller
	Call setSAASController()

	' Check connection to Controller
	Call getSAASControllerDetails()

	' Select application, key, tier and node
	Call setAppName()
	Call setAppTierName()
	Call setAppNodeName()

	' Select type of Java Agent
	' autodeploy not done from here on
	call setJavaAgentType()
	
	' Select type of application and prepare for instrumentation
	call setJavaAgentAppType()

	' Copy selected Java Agent into installation dir
	JavaAgentInstallDir = DiskDrive & "\AppDynamics\JavaAgent"
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if Java Agent installation folder already exists."
	
	If objFS0.FolderExists(JavaAgentInstallDir) Then
	
		Set objFolder = objFS0.GetFolder(JavaAgentInstallDir)
	
		If (objFolder.Files.Count > 0) Or (ObjFolder.SubFolders.Count > 0) Then
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder already exists and is not empty. Setup will quit."
			WScript.Quit
	
		End If
	
	Else
	
		wshShell.Run ("cmd.exe /C mkdir " & JavaAgentInstallDir), 0, True
	
	End If
	
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder doesn't exist. OK to continue."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Copying Java Agent (" & JavaAgentType & ") to " & JavaAgentInstallDir
	objFS0.CopyFolder JavaAgentSourceDir, JavaAgentInstallDir, True

	' Locate conf folder and controller-info.xml we will be using
	Set objFolder = objFS0.GetFolder(JavaAgentInstallDir)
	Set objSubFolder = objFolder.SubFolders
	
	For each subFolder in objSubFolder
	
		If instr(subFolder, "ver") Then
	
			JavaAgentHomeDir = subFolder
			JavaAgentConfDir = subFolder & "\conf\"
			JavaAgentControllerXML = JavaAgentConfDir & "controller-info.xml"
			JavaAgentSCSTool = subFolder & "\utils\scs\scs-tool.jar"
	
		End If
	
	Next
	
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Dir: " & JavaAgentConfDir
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Conf: " & JavaAgentControllerXML

	' Update controller-info.xml
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Copying cacerts.jks."
	JavaAgentSourceControllerKeystore = objFS0.GetFile("conf\cacerts.jks")
	JavaAgentControllerKeystore = JavaAgentConfDir & "cacerts.jks"
	JavaAgentControllerKeystorePass = "changeit"
	objFS0.CopyFile JavaAgentSourceControllerKeystore, JavaAgentControllerKeystore, True
	
	If objFS0.FileExists(JavaAgentControllerXML) Then
	
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Updating controller-info.xml."
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = "false"
			xmlDoc.Load(JavaAgentControllerXML)
		Set xmlDocControllerInfo = xmlDoc.selectSingleNode("//controller-info")
		Set xmlDocControllerHost = xmlDoc.selectSingleNode("//controller-info/controller-host")
		Set xmlDocControllerSSL = xmlDoc.selectSingleNode("//controller-info/controller-ssl-enabled")
		Set xmlDocControllerPort = xmlDoc.selectSingleNode("//controller-info/controller-port")
		Set xmlDocAccountName = xmlDoc.selectSingleNode("//controller-info/account-name")
		Set xmlDocAccountKey = xmlDoc.selectSingleNode("//controller-info/account-access-key")
		Set xmlDocUseEncryptedCreds = xmlDoc.selectSingleNode("//controller-info/use-encrypted-credentials")
		Set xmlDocKeystoreFile = xmlDoc.selectSingleNode("//controller-info/controller-keystore-filename")
		Set xmlDocKeystorePass = xmlDoc.selectSingleNode("//controller-info/controller-keystore-password")
		xmlDocControllerHost.Text = ControllerHost
		xmlDocControllerSSL.Text = ControllerSSL
		xmlDocControllerPort.Text = ControllerPort
		xmlDocAccountName.Text = ControllerAccountName
		xmlDocAccountKey.Text = ControllerAccessKey
	
		If xmlDocUseEncryptedCreds Is Nothing Then
	
			Set xmlDocUseEncryptedCreds= _
				xmlDoc.createElement("use-encrypted-credentials")
				xmlDocUseEncryptedCreds.Text = "false"
				xmlDocControllerInfo.appendChild xmlDocUseEncryptedCreds
	
		Else
	
			xmlDocUseEncryptedCreds.Text = "false"
	
		End If
	
		If xmlDocKeystoreFile Is Nothing Then
	
			Set xmlDocKeystoreFile= _
				xmlDoc.createElement("controller-keystore-filename")
				xmlDocKeystoreFile.Text = "cacerts.jks"
				xmlDocControllerInfo.appendChild xmlDocKeystoreFile
	
		Else
	
			xmlDocKeystoreFile.Text = "cacerts.jks"
	
		End If
	
		If xmlDocKeystorePass Is Nothing Then
	
			Set xmlDocKeystorePass= _
				xmlDoc.createElement("controller-keystore-password")
				xmlDocKeystorePass.Text = JavaAgentControllerKeystorePass
				xmlDocControllerInfo.appendChild xmlDocKeystorePass
	
		Else
	
			xmlDocKeystorePass.Text = JavaAgentControllerKeystorePass
	
		End If
	
		xmlDoc.save(JavaAgentControllerXML)
	
	Else
	
		WScript.Echo FormatDateTime(Now, vbLongTime) * " - Could not find controller-info.xml"
		WScript.Quit
	
	End If

	' Instrument selected application
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumenting " & CustomerApplication & "."
	
	If CustomerApplication = "JBoss" Then 

		Set objFile = objFS0.OpenTextFile(appStartUpScript, ForAppending)
		objFile.WriteLine LineBuffer
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -javaagent:" & JavaAgentHomeDir & "\javaagent.jar"
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.agent.applicationName=" & ControllerApplicationName
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.agent.tierName=" & appTierName
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.agent.nodeName=" & appNodeName

		If proxyEnabled = "true" Then
			objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.http.proxyHost=" & proxyHost
			objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.http.proxyPort=" & proxyPort
		End If
		
		objFile.close
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		Set objFile = objFS0.OpenTextFile(appStartUpScript, ForReading)
		FileContents = objFile.ReadAll
		
		If InStr(FileContents, "-Dappdynamics.") Then
			
			If InStr(FileContents, "-javaagent:") Then
				
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."
			
			Else
				
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
				WScript.Quit
			
			End If
		
		Else
			
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application start up script is instrumented."
			WScript.Quit
		
		End If
		
		objFile.close

	ElseIf CustomerApplication = "Tomcat" Then

		ArrTomcatOptions = Array()
		oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,ValueName,ArrTomcatOptions
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-javaagent:" & JavaAgentHomeDir & "\javaagent.jar" 
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.agent.applicationName=" & ControllerApplicationName
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.agent.tierName=" & appTierName
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.agent.nodeName=" & appNodeName
		
		If proxyEnabled = "true" Then
			
			ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
			ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.http.proxyHost=" & proxyHost
			ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
			ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.http.proxyPort=" & proxyPort
		
		End If
		
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		oReg.SetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,ValueName,ArrTomcatOptions
		
		For each ArrTomcatOptions in ArrTomcatOptions
			
			If (InStr(ArrTomcatOptions, "-javaagent")) Or (InStr(ArrTomcatOptions, "-Dappdynamics")) Then 
				
				IsInstrumented = 1
			
			End If
		
		Next
		
		If IsInstrumented = 1 Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."
		
		Else
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
			WScript.Quit
		
		End If

	ElseIf CustomerApplication = "WebSphereAS" Then

		' Enable metrics collection from JVM via server.xml
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.Async = "false"
		xmlDoc.Load(WebSphereASServerXML)
		
		If WebSphereASstatisticSet = "none" Then
		
			Set xmlDocPMIService = xmlDoc.selectSingleNode(".//node()[@xmi:type = 'pmiservice:PMIService']")
			xmlDocPMIService.Attributes.getNamedItem("statisticSet").Text = "basic"
		
		End If
		
		' Instrument server.xml
		Set xmlDocJvmEntries = xmlDoc.selectSingleNode("//process:Server/processDefinitions/jvmEntries")
		
		If xmlDocJvmEntries is Nothing Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Couldn't find JVM params in server.xml."
			WScript.Quit
		
		Else
		
			JvmParams = xmlDocJvmEntries.Attributes.getNamedItem("genericJvmArguments").Text
		
			If proxyEnabled = "true" Then
		
				xmlDocJvmEntries.Attributes.getNamedItem("genericJvmArguments").Text = JvmParams & " -javaagent:" & JavaAgentHomeDir & "\javaagent.jar" & " -Dappdynamics.agent.applicationName=" & ControllerApplicationName & " -Dappdynamics.agent.tierName=" & appTierName & " -Dappdynamics.agent.nodeName=" & appNodeName & " -Dappdynamics.http.proxyHost=" & proxyHost & " -Dappdynamics.http.proxyPort=" & proxyPort

			ElseIf proxyEnabled = "false" Then
		
				xmlDocJvmEntries.Attributes.getNamedItem("genericJvmArguments").Text = JvmParams & " -javaagent:" & JavaAgentHomeDir & "\javaagent.jar" & " -Dappdynamics.agent.applicationName=" & ControllerApplicationName & " -Dappdynamics.agent.tierName=" & appTierName & " -Dappdynamics.agent.nodeName=" & appNodeName
		
			End If
		
		End If
		
		xmlDoc.Save(WebSphereASServerXML)
		Set xmlDoc = Nothing
		
		' Update server.policy
		If IsWebSphereASPolicyFileConfigured = 1 Then
		
		Else
		
			Set objFile = objFS0.OpenTextFile(WebSphereASPolicyFile, ForReading)
			strText = objFile.ReadAll
			objFile.Close
			'strText = strText & vbNewLine & "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentHomeDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
			strText = strText & vbNewLine & "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentInstallDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
			Set objFile = objFS0.OpenTextFile(WebSphereASPolicyFile, ForWriting)
			objFile.Write strText
			objFile.Close
		
		End If
		
		' Verify instrumentation of server.xml
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		Set objFile = objFS0.OpenTextFile(WebSphereASServerXML, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.close
		
		For each Line in FileContents
		
			If (InStr(Line, "-Dappdynamics")) And (InStr(Line, "-javaagent:")) Then
		
				IsInstrumented = "1"
		
			Else
		
			End If
		
		Next
		
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.Async = "false"
		xmlDoc.Load(WebSphereASServerXML)
		Set xmlDocPMIService = xmlDoc.selectSingleNode(".//node()[@xmi:type = 'pmiservice:PMIService']")
		WebSphereASstatisticSet = UCase(xmlDocPMIService.Attributes.getNamedItem("statisticSet").Text)
		
		If IsInstrumented = 1 And WebSphereASstatisticSet <> UCase("none") Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."
		
		Else
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
			WScript.Quit
		
		End If

	ElseIf CustomerApplication = "WebLogic" Then
		
		call instrumentWebLogicJavaAgent()

	Else
		
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a valid application selection."
		WScript.Quit

	End if

End Function

Function Install_ma()

	' Validate Windows version
	call validateWindowsVersion()
	
	' Select the right installation package
	If WinArch = 32 Then
		objStartFolder = "binaries\MachineInstall\Windows\x86"
	ElseIf WinArch = 64 Then
		objStartFolder = "binaries\MachineInstall\Windows\x64"
	End If
	Set objFolder = objFS0.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		maInstallSource = objFolder & "\" & objFile.name
	Next

	' Reuse configuration if .Net Agent exists
	On Error Resume Next
	Err.Clear
	registryDataDir = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AppDynamics\dotNet Agent\DotNetAgentFolder")

	If Err.Number <> 0 Then

		DotNetCompatibility = "false"
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No .Net Agent found on system."
		Err.Clear

		' Select the type of Controller
		Call setSAASController()
	
		' Check connection to Controller
		Call getSAASControllerDetails()

		' Select application, key, tier and node
		If IsEmpty(autoAppAgentExists) Then
			WScript.Echo ""
			WScript.StdOut.Write "Host contains application (Java/Node.js/PHP/etc) agent (yes/no)?: "
			PromptContinue = WScript.StdIn.ReadLine
		Else
			PromptContinue = autoAppAgentExists
		End If
		If InStr("|yes|y|YES|Y|", "|" & PromptContinue & "|") Then
			MachineAgentOnly = "false"
		ElseIf InStr("|no|n|NO|N|", "|" & PromptContinue & "|") Then
			MachineAgentOnly = "true"
		Else
			WScript.Echo ""
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: accepts only 'yes' or 'no'."
			WScript.Quit
		End If
		Call setAppName()

		' Set tier and node for standalone install
		If MachineAgentOnly = "true" Then
			Call setAppTierName()
			Call setAppNodeName()
		ElseIf MachineAgentOnly = "false" Then
		End If

	Else

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Reusing existing .Net Agent configuration."
		' Load properties file
		dataDir = registryDataDir & "Config\config.xml"
		If objFS0.FileExists(dataDir) Then
			Set xmlDoc = CreateObject("Microsoft.XMLDOM")
				xmlDoc.Async = "false"
				xmlDoc.Load(dataDir)
			Set xmlDocControllerInfo = xmlDoc.selectSingleNode("//appdynamics-agent/controller")
			Set xmlDocControllerApplicationName = xmlDoc.selectSingleNode("//appdynamics-agent/controller/application")
			Set xmlDocAccountInfo = xmlDoc.selectSingleNode("//appdynamics-agent/controller/account")
			ControllerHost = xmlDocControllerInfo.Attributes.getNamedItem("host").Text
			ControllerPort = xmlDocControllerInfo.Attributes.getNamedItem("port").Text
			ControllerSSL = xmlDocControllerInfo.Attributes.getNamedItem("ssl").Text
			ControllerApplicationName = xmlDocControllerApplicationName.Attributes.getNamedItem("name").Text
			ControllerAccountName = xmlDocAccountInfo.Attributes.getNamedItem("name").Text
			ControllerAccessKey = xmlDocAccountInfo.Attributes.getNamedItem("password").Text
			DotNetCompatibility = "true"

		Else

			WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found " & dataDir & "."
			WScript.Quit
		
		End If
		
	End If

	On Error Goto 0

	' Select if SIM monitoring is needed (advanced infrastructure monitoring)
	Call setSimEnabled()

	' Copy installation
	WScript.Echo ""
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Copying installation bits."
	maHomeDir = DiskDrive & "\AppDynamics\MachineAgent"
	If NOT objFS0.FolderExists(maHomeDir) Then
		wshShell.Run ("cmd.exe /C mkdir " & maHomeDir), 0, True
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Machine Agent installation folder already exists."
		WScript.Quit
	End If
	set FilesInZip = objShell.NameSpace(maInstallSource).items
	objShell.NameSpace(maHomeDir).CopyHere(FilesInZip)
	If objFS0.FolderExists(maHomeDir) Then
		Set objFolder = objFS0.GetFolder(maHomeDir)
		If (objFolder.Files.Count = 0) Or (ObjFolder.SubFolders.Count = 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder already exists but is empty and it shouldn't be."
			WScript.Quit
		ElseIf (objFolder.Files.Count > 0) And (ObjFolder.SubFolders.Count > 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder check after unzip operation succeeded."
		End If
	End If
	

	' Update Configuration
	maConfDir = maHomeDir & "\conf\"
	maControllerXML = maConfDir & "controller-info.xml"
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Copying cacerts.jks."
	maSourceControllerKeystore = objFS0.GetFile("conf\cacerts.jks")
	maControllerKeystore = maConfDir & "cacerts.jks"
	maControllerKeystorePass = "changeit"
	objFS0.CopyFile maSourceControllerKeystore, maControllerKeystore
	If objFS0.FileExists(maControllerXML) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Updating controller-info.xml."
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
			xmlDoc.Async = "false"
			xmlDoc.Load(maControllerXML)
		Set xmlDocControllerInfo = xmlDoc.selectSingleNode("//controller-info")
		Set xmlDocControllerHost = xmlDoc.selectSingleNode("//controller-info/controller-host")
		Set xmlDocControllerSSL = xmlDoc.selectSingleNode("//controller-info/controller-ssl-enabled")
		Set xmlDocControllerPort = xmlDoc.selectSingleNode("//controller-info/controller-port")
		Set xmlDocAccountName = xmlDoc.selectSingleNode("//controller-info/account-name")
		Set xmlDocAccountKey = xmlDoc.selectSingleNode("//controller-info/account-access-key")
		Set xmlDocApplicationName = xmlDoc.selectSingleNode("//controller-info/application-name")
		Set xmlDocTierName = xmlDoc.selectSingleNode("//controller-info/tier-name")
		Set xmlDocNodeName = xmlDoc.selectSingleNode("//controller-info/node-name")
		Set xmlDocUseEncryptedCreds = xmlDoc.selectSingleNode("//controller-info/use-encrypted-credentials")
		Set xmlDocKeystoreFile = xmlDoc.selectSingleNode("//controller-info/controller-keystore-filename")
		Set xmlDocKeystorePass = xmlDoc.selectSingleNode("//controller-info/controller-keystore-password")
		Set xmlDocSimEnabled = xmlDoc.selectSingleNode("//controller-info/sim-enabled")
		Set xmlDocDotNetCompatibilityEnabled = xmlDoc.selectSingleNode("//controller-info/dotnet-compatibility-mode")
		Set xmlDocUniqueHostId = xmlDoc.selectSingleNode("//controller-info/unique-host-id")
		xmlDocControllerHost.Text = ControllerHost
		xmlDocControllerSSL.Text = ControllerSSL
		xmlDocControllerPort.Text = ControllerPort
		xmlDocAccountName.Text = ControllerAccountName
		xmlDocAccountKey.Text = ControllerAccessKey
		xmlDocDotNetCompatibilityEnabled.Text = DotNetCompatibility
		xmlDocUniqueHostId.Text = strComputerName
		If xmlDocUseEncryptedCreds Is Nothing Then
			Set xmlDocUseEncryptedCreds= _
				xmlDoc.createElement("use-encrypted-credentials")
				xmlDocUseEncryptedCreds.Text = "false"
				xmlDocControllerInfo.appendChild xmlDocUseEncryptedCreds
		Else
			xmlDocUseEncryptedCreds.Text = "false"
		End If
		If xmlDocKeystoreFile Is Nothing Then
			Set xmlDocKeystoreFile= _
				xmlDoc.createElement("controller-keystore-filename")
				xmlDocKeystoreFile.Text = ""
				xmlDocControllerInfo.appendChild xmlDocKeystoreFile
		Else
			xmlDocKeystoreFile.Text = ""
		End If
		If xmlDocKeystorePass Is Nothing Then
			Set xmlDocKeystorePass= _
				xmlDoc.createElement("controller-keystore-password")
				xmlDocKeystorePass.Text = maControllerKeystorePass
				xmlDocControllerInfo.appendChild xmlDocKeystorePass
		Else
			xmlDocKeystorePass.Text = maControllerKeystorePass
		End If
		
		If MachineAgentOnly = "true" Then
			If xmlDocApplicationName Is Nothing Then
				Set xmlDocApplicationName= _
					xmlDoc.createElement("application-name")
					xmlDocApplicationName.Text = ControllerApplicationName
					xmlDocControllerInfo.appendChild xmlDocApplicationName
			Else
				xmlDocApplicationName.Text = ControllerApplicationName
			End If
			If xmlDocTierName Is Nothing Then
				Set xmlDocTierName= _
					xmlDoc.createElement("tier-name")
					xmlDocTierName.Text = appTierName
					xmlDocControllerInfo.appendChild xmlDocTierName
			Else
				xmlDocTierName.Text = appTierName
			End If
			If xmlDocNodeName Is Nothing Then
				Set xmlDocNodeName= _
					xmlDoc.createElement("node-name")
					xmlDocNodeName.Text = appNodeName
					xmlDocControllerInfo.appendChild xmlDocNodeName
			Else
				xmlDocNodeName.Text = appNodeName
			End If
		ElseIf MachineAgentOnly = "false" Then
		End If

		If xmlDocSimEnabled Is Nothing Then
			Set xmlDocSimEnabled= _
				xmlDoc.createElement("sim-enabled")
				xmlDocSimEnabled.Text = simEnabled
				xmlDocControllerInfo.appendChild xmlDocSimEnabled
		Else
			xmlDocSimEnabled.Text = simEnabled
		End If

		xmlDoc.save(maControllerXML)
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) * " - Could not find controller-info.xml"
		WScript.Quit
	End If
	
	' Deploy and start machine agent service
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Installing AppDynamics Machine Agent service."
	wshShell.Run ("cmd.exe /C cscript " & maHomeDir & "\InstallService.vbs"), 0, False
	WScript.Sleep 10000
	Set cmd = wshShell.Exec("cmd.exe /C sc query ""AppDynamics Machine Agent"" | find ""STATE""")
	WaitForWshShellExec()
	maServiceState = Right((cmd.StdOut.ReadLine()), 8)
	If maServiceState = "RUNNING " Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Machine Agent is up and running."
		
		' Set proxy settings for Machine Agent after installed and up and running
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Updating proxy settings."
		Set objFile = objFS0.OpenTextFile(maHomeDir & "\bin\MachineAgentService.vmoptions", ForReading)
		strText = objFile.ReadAll
		objFile.Close
		strText = "-Dappdynamics.http.proxyHost=" & proxyHost & vbNewLine & "-Dappdynamics.http.proxyPort=" & proxyPort & vbNewLine & strText
		Set objFile = objFS0.OpenTextFile(maHomeDir & "\bin\MachineAgentService.vmoptions", ForWriting)
		objFile.Write strText
		objFile.Close
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Restarting Machine Agent Service."
		wshShell.Run("cmd.exe /C net stop ""Appdynamics Machine Agent"""), 0, True
		wshShell.Run("cmd.exe /C net start ""Appdynamics Machine Agent"""), 0, True
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Machine Agent is not running."
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Proxy could not be set."
	End If

End Function

Function Uninstall_dna()

	' Check for existing installation
	On Error Resume Next
	Err.Clear
	registryInstallationDir = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AppDynamics\dotNet Agent\InstallationDir")
	If Err.Number <> 0 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No existing agent installation detected."
		Err.Clear
		WScript.Quit
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Detected existing install."
		On Error Goto 0
	
		' Stop IIS and agent
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Stopping IIS."
		wshShell.Run("cmd.exe /C iisreset /stop"), 0, True
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Stopping AppDynamics .NET Agent service."
		wshShell.Run("cmd.exe /C net stop AppDynamics.Agent.Coordinator_service"), 0, True
		DotNetInstallLog = "DotNetAgentInstall.log"
			
		' Look through registry for existing configuration
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Looking for existing configuration."
		regDotNetAgentFolder = WshShell.Regread("HKLM\SOFTWARE\AppDynamics\dotNet Agent\DotNetAgentFolder")
		Set objectReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
		objectReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
		searchString = "AppDynamics .NET Agent"
		For Each subkey In arrSubKeys
			On Error Resume Next
			regDisplayName = WshShell.Regread("HKLM\" & strKeyPath & "\" & subkey & "\DisplayName")
			if err.number = 0 then 
				On Error Goto 0
				if instr(regDisplayName, searchString) > 0 then
					On Error Resume Next
					WScript.Echo FormatDateTime(Now, vbLongTime) & " - Found registry information."
					regUninstallString = WshShell.Regread("HKLM\" & strKeyPath & "\" & subkey & "\UnInstallString")
					On Error GoTo 0
				end if
			end if
		Next
		On Error Goto 0

		' Start uninstall
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Running uninstall."
		WshShell.Run(regUninstallString & " /qn" & " /norestart" & " /lv " & DotNetInstallLog), 0, True
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Uninstall finished. Check " & DotNetInstallLog & " for more details."
		If Right(regDotNetAgentFolder, 1) = "\" Then
			regDotNetAgentFolder = Left(regDotNetAgentFolder, Len(regDotNetAgentFolder) - 1)
		Else
		End If
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Cleaning remaining configuration."
		objFS0.DeleteFolder regDotNetAgentFolder
		wshShell.RegDelete "HKLM\SOFTWARE\AppDynamics\dotNet Agent\"
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Starting IIS."
		wshShell.Run("cmd.exe /C iisreset /start"), 0, True
	End If

End Function

Function Uninstall_ma()

	' Check for existing installation
	Set cmd = wshShell.Exec("cmd.exe /c sc qc ""AppDynamics Machine Agent"" | find ""BINARY_PATH_NAME""")
	WaitForWshShellExec()
	cmd = Mid((cmd.StdOut.ReadLine()), 30)
	If InStr(cmd, "nssm") Then
		maRuntimeParamsFile = Replace(cmd, "nssm.exe", "runAgentAsService.bat")
		maHomeDir = Replace(cmd, "\nssm.exe", "")
	ElseIF InStr(cmd, "bin\MachineAgentService.exe") Then
		maControllerXML = Replace(cmd, "bin\MachineAgentService.exe", "conf\controller-info.xml")
		maHomeDir = Replace(cmd, "\bin\MachineAgentService.exe", "")
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Unable to retrieve Machine Agent information."
		WScript.Quit
	End if

	' Backup existing installation
	If objFs0.FolderExists(maHomeDir) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Home dir: " & maHomeDir & "."
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Backing up existing installation folder."
		objFs0.CopyFolder maHomeDir, mahomeDir & "-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Unable to locate existing Machine Agent home dir."
		WScript.Quit
	End If

	' Uninstall existing agent
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Stopping Machine Agent service."
	wshShell.Run "cmd.exe /c sc stop ""AppDynamics Machine Agent""", 0, True
	If objFs0.FileExists(maRuntimeParamsFile) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Uninstalling Machine Agent service."
		Set cmd = wshShell.Exec("cmd.exe /c " & cmd & " remove ""AppDynamics Machine Agent"" confirm")
		WaitForWshShellExec()
	ElseIf objFs0.FileExists(maControllerXML) Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Uninstalling Machine Agent service."
		Set cmd = wshShell.Exec("cmd.exe /c cscript " & maHomedir & "\UninstallService.vbs")
		WScript.Sleep 10000
	Else 
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Unable to locate existing Machine Agent settings."
		WScript.Quit
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Deleting Machine Agent home dir."
	objFs0.DeleteFolder maHomeDir, True

End Function

Function Upgrade_dna()

	' Validate upgrade install folder
	call validateUpgrDir("binaries\DotNetUpgrade", ".msi")

	' Get current version
	on error resume next
	DotNetAgentCurrentVersion = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AppDynamics\dotNet Agent\Version")
	If err.number = 0 then
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Error reading registry."
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - " & err.description
		WScript.Quit
	End If
	on error goto 0

	' Get new version
	Set objFolder = objFS0.GetFolder("binaries\DotNetUpgrade")
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		UpgradeMSI = objFolder & "\" & objFile.name
	Next
	WshShell.Run ("msiexec /a " & UpgradeMSI & " TARGETDIR=" & tmpFolder & "\upgrade /qn"), 0, True
	DotNetAgentUpgradeToVersion = objFS0.GetFileVersion("tmp\upgrade\AppDynamics\AppDynamics .NET Agent\GAC\AppDynamics.Agent.dll")
	objFS0.DeleteFolder("tmp\upgrade")

	' Compare current vs new version
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Comparing current (" & DotNetAgentCurrentVersion & ") and new (" & DotNetAgentUpgradeToVersion & ") version."
	If DotNetAgentCurrentVersion >= DotNetAgentUpgradeToVersion Then
    	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Version check failed."
		WScript.Quit
	Else
    	WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Version check succeeded."
	End If

	' Upgrade agent
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Starting upgrade of .Net agent."
	DotNetUpgradeLog = "DotNetUpgrade.log"
	DotNetInstallDir = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AppDynamics\dotNet Agent\InstallationDir")
	WshShell.Run ("msiexec " & "/i " & UpgradeMSI & " /q" & " /norestart" & " /lv " & Chr(34) & DotNetUpgradeLog & Chr(34) & " INSTALLDIR=" & Chr(34) & DotNetInstallDir & Chr(34)), 0, True

	' Verify
	DotNetAgentCurrentVersion = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AppDynamics\dotNet Agent\Version")
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Version after upgrade: " & DotNetAgentCurrentVersion & "."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Validating service state."
	Set cmd = wshShell.Exec("cmd.exe /c sc query ""AppDynamics.Agent.Coordinator_service"" | find ""STATE""")
	WaitForWshShellExec()
	cmd = Mid((cmd.StdOut.ReadLine()), 30)
	If InStr(cmd, "RUNNING") Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Service state validation succeeded."
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Couldn't verify service state. Service is not running."
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Try starting AppDynamics.Agent.Coordinator service manually."
		WScript.Quit
	End If

End Function

Function Upgrade_ma()

	' Validate upgrade install folder
	call validateUpgrDir("binaries\MachineUpgrade", ".zip")

	' Validate Windows version
	call validateWindowsVersion()

	' Check for existing installation
	Set cmd = wshShell.Exec("cmd.exe /c sc qc ""AppDynamics Machine Agent"" | find ""BINARY_PATH_NAME""")
	WaitForWshShellExec()
	cmd = Mid((cmd.StdOut.ReadLine()), 30)
	If InStr(cmd, "nssm") Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Upgrade does not support this installation. NSSM detected."
		WScript.Quit
	ElseIF InStr(cmd, "bin\MachineAgentService.exe") Then
		maSourceControllerXML = Replace(cmd, "bin\MachineAgentService.exe", "conf\controller-info.xml")
		maSourceControllerKeystore = Replace(cmd, "bin\MachineAgentService.exe", "conf\cacerts.jks")
		maHomeDir = Replace(cmd, "\bin\MachineAgentService.exe", "")
		maMonitorsDir = Replace(cmd, "\bin\MachineAgentService.exe", "\monitors")
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Unable to retrieve Machine Agent information."
		WScript.Quit
	End if

	' Get current version
	maOutputLog = maHomeDir & "\bin\output.log"
	If objFS0.FileExists(maOutputLog) Then
		Set objFile = objFS0.OpenTextFile(maOutputLog, ForReading)
		FileContents = objFile.ReadAll
		objFile.close
		RegExp.Global = True
		RegExp.Pattern = "v\d+\.\d+\.\d+-\d+"
		Set Matches = RegExp.Execute(FileContents)
		For each Match in Matches
    		MachineAgentCurrentVersion = Mid(Match, 2)
			MachineAgentCurrentVersion = Replace(MachineAgentCurrentVersion, "-", ".")
		Next
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - output.log not found. Unable to determine existing version."
		WScript.Quit
	End If

	' Get new version
	Set objFolder = objFS0.GetFolder("binaries\MachineUpgrade")
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		UpgradeZIP = objFolder & "\" & objFile.name
	Next
	RegExp.Global = True
	RegExp.Pattern = "\d+\.\d+\.\d+\..*"
	Set Matches = RegExp.Execute(UpgradeZIP)
	For each Match in Matches
		MachineAgentUpgradeToVersion = Replace(Match, ".zip", "")
	Next
	
	' Compare current vs new version
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Comparing current (" & MachineAgentCurrentVersion & ") and new (" & MachineAgentUpgradeToVersion & ") version."
	If MachineAgentUpgradeToVersion > MachineAgentCurrentVersion  Then
    	WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Version check succeeded."
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Version check failed."
		WScript.Quit
	End If
	
	' Temp upgrade folders
	tmpUpgrFolder = tmpFolder & "\upgrade"
	tmpUpgrMonFolder = tmpFolder & "\upgrade\monitors"
	If objFS0.FolderExists(tmpUpgrFolder) Then 
		Set objFolder = objFS0.GetFolder(tmpUpgrFolder)
		objFS0.DeleteFolder(tmpUpgrFolder)
		objFS0.CreateFolder(tmpUpgrFolder)
		objFS0.CreateFolder(tmpUpgrMonFolder)
	Else
		objFS0.CreateFolder(tmpUpgrFolder)
		objFS0.CreateFolder(tmpUpgrMonFolder)
	End If
	
	' Backup existing cacerts.jks if it exists
	WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Backing up existing configuration."
	on error resume next
	If objFS0.FileExists(maSourceControllerKeystore) Then
		objFS0.CopyFile maSourceControllerKeystore, tmpUpgrFolder & "\cacerts.jks"
		maSourceControllerKeystore = tmpUpgrFolder & "\cacerts.jks"
	Else
	End If
	on error goto 0

	' Backup existing controller-info.xml
	If objfs0.FileExists(maSourceControllerXML) Then
		objFS0.CopyFile maSourceControllerXML, tmpUpgrFolder & "\controller-info.xml"
		maSourceControllerXML = tmpUpgrFolder & "\controller-info.xml"
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No existing controller-info.xml found."
		WScript.Quit 
	End If
	
	' Backup analytics-agent.properties
	maSourceAnalyticsAgentProperties = maHomeDir & "\monitors\analytics-agent\conf\analytics-agent.properties"
	If objFs0.FileExists(maSourceAnalyticsAgentProperties) Then
		objFS0.CopyFile maSourceAnalyticsAgentProperties, tmpUpgrFolder & "\analytics-agent.properties"
		maSourceAnalyticsAgentProperties = tmpUpgrFolder & "\analytics-agent.properties"
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No existing analytics-agent.properties found."
		WScript.Quit
	End If

	' Backup MachineAgentService.vmoptions if it exists
	maSourceServiceVmoptions = maHomeDir & "\bin\MachineAgentService.vmoptions"
	If objFs0.FileExists(maSourceServiceVmoptions) Then
		objFS0.CopyFile maSourceServiceVmoptions, tmpUpgrFolder & "\MachineAgentService.vmoptions"
		maSourceServiceVmoptions = tmpUpgrFolder & "\MachineAgentService.vmoptions"
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - No existing MachineAgentService.vmoptions found but will continue."
	End If

	' Backup other monitors from current agent
	WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Scanning for custom monitors to be migrated."
	Set objFolder = objFS0.GetFolder(maMonitorsDir)
	Set colFolders = objFolder.SubFolders
	arrCustomMonitors = Array()
	'For Each objFolder in colFolders
	For Each colFolder in colFolders
		If InStr(colFolder, "\analytics-agent") Or InStr(colFolder, "\JavaHardwareMonitor") Or InStr(colFolder, "\HardwareMonitor") Then
		Else 
			ReDim Preserve arrCustomMonitors(UBound(arrCustomMonitors) + 1)
			arrCustomMonitors(UBound(arrCustomMonitors)) = colFolder
		End If
	Next
	If (UBound(arrCustomMonitors) + 1) = 0 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) &  " - No custom monitors found."
		SkipMonitorsSetup = "true"
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Found " & (UBound(arrCustomMonitors) + 1) & " custom monitors for migration. Creating backup."
		SkipMonitorsSetup = "false"
	End If
	For Each arrCustomMonitors in arrCustomMonitors
		objFS0.CopyFolder arrCustomMonitors, tmpUpgrMonFolder & "\"
	Next

	' Uninstall existing agent
	call Uninstall_ma()
	
	' Install new agent version and upgrade settings
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Copying installation bits."
	If NOT objFS0.FolderExists(maHomeDir) Then
		wshShell.Run ("cmd.exe /C mkdir " & maHomeDir), 0, True
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Machine Agent installation folder already exists."
		WScript.Quit
	End If
	set FilesInZip = objShell.NameSpace(UpgradeZIP).items
	objShell.NameSpace(maHomeDir).CopyHere(FilesInZip)
	If objFS0.FolderExists(maHomeDir) Then
		Set objFolder = objFS0.GetFolder(maHomeDir)
		If (objFolder.Files.Count = 0) Or (ObjFolder.SubFolders.Count = 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder already exists but is empty and it shouldn't be."
			WScript.Quit
		ElseIf (objFolder.Files.Count > 0) And (ObjFolder.SubFolders.Count > 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder check after unzip operation succeeded."
		End If
	End If
	on error resume next
	If objFS0.FileExists(maSourceControllerKeystore) Then
		objFS0.CopyFile maSourceControllerKeystore, maHomeDir & "\conf\cacerts.jks"
	Else
	End If
	on error goto 0
	objFS0.CopyFile maSourceControllerXML, maHomeDir & "\conf\controller-info.xml"
	objFS0.CopyFile maSourceAnalyticsAgentProperties, maHomeDir & "\monitors\analytics-agent\conf\analytics-agent.properties"
	If objFs0.FileExists(maSourceServiceVmoptions) Then
		objFS0.CopyFile maSourceServiceVmoptions, maHomeDir & "\bin\MachineAgentService.vmoptions"
		SkipVmoptionsSetup = "true"
	Else
		SkipVmoptionsSetup = "false"
	End If
	If SkipMonitorsSetup = "false" Then
		Set objFolder = objFS0.GetFolder(tmpUpgrMonFolder)
		Set colFolders = objFolder.SubFolders
		For Each colFolder in colFolders
			objFS0.CopyFolder colFolder, maHomeDir & "\monitors\"
		Next
	Else
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Installing AppDynamics Machine Agent service."
	wshShell.Run ("cmd.exe /C cscript " & maHomeDir & "\InstallService.vbs"), 0, False
	WScript.Sleep 10000
	Set cmd = wshShell.Exec("cmd.exe /C sc query ""AppDynamics Machine Agent"" | find ""STATE""")
	WaitForWshShellExec()
	maServiceState = Right((cmd.StdOut.ReadLine()), 8)
	If maServiceState = "RUNNING " Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Machine Agent is up and running."
		If SkipVmoptionsSetup = "true" Then
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Updating proxy settings."
			Set objFile = objFS0.OpenTextFile(maHomeDir & "\bin\MachineAgentService.vmoptions", ForReading)
			strText = objFile.ReadAll
			objFile.Close
			strText = "-Dappdynamics.http.proxyHost=" & proxyHost & vbNewLine & "-Dappdynamics.http.proxyPort=" & proxyPort & vbNewLine & strText
			Set objFile = objFS0.OpenTextFile(maHomeDir & "\bin\MachineAgentService.vmoptions", ForWriting)
			objFile.Write strText
			objFile.Close
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Restarting Machine Agent Service."
			wshShell.Run("cmd.exe /C net stop ""Appdynamics Machine Agent"""), 0, True
			wshShell.Run("cmd.exe /C net start ""Appdynamics Machine Agent"""), 0, True
		End If
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Machine Agent is not running."
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Proxy could not be set."
	End If
	objFS0.DeleteFolder("tmp\upgrade")

End Function

Function Upgrade_ja()

	' Validate upgrade install folder
	call validateUpgrDir("binaries\JavaUpgrade", ".zip")

	' Temp upgrade folders
	tmpUpgrFolder = tmpFolder & "\upgrade"
	If objFS0.FolderExists(tmpUpgrFolder) Then 
		Set objFolder = objFS0.GetFolder(tmpUpgrFolder)
		objFS0.DeleteFolder(tmpUpgrFolder)
		objFS0.CreateFolder(tmpUpgrFolder)
	Else
		objFS0.CreateFolder(tmpUpgrFolder)
	End If

	' Get existing installation
	If IsEmpty(autoAgentPath) Then
		WScript.Echo ""
		WScript.StdOut.Write "Please provide -javaagent value: "
		PromptContinue = WScript.StdIn.ReadLine
		JavaAgentJAR = Replace(PromptContinue, "-javaagent:", "")
		JavaAgentHomeDir = Replace(Replace(PromptContinue, "-javaagent:", ""), "\javaagent.jar", "")
	Else
		JavaAgentJAR = Replace(autoAgentPath, "-javaagent:", "")
		JavaAgentHomeDir = Replace(Replace(autoAgentPath, "-javaagent:", ""), "\javaagent.jar", "")
	End If
	If objFS0.FolderExists(JavaAgentHomeDir) Then
	Else
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Folder not found."
		WScript.Quit
	End If

	' Backup existing controller-info.xml
	WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Backing up existing configuration."
	objFs0.CopyFile JavaAgentHomeDir & "\conf\controller-info.xml", tmpUpgrFolder & "\controller-info.xml"
	JavaAgentSourceControllerXML = tmpUpgrFolder & "\controller-info.xml"

	' Backup existing cacerts.jks if it exists
	JavaAgentSourceControllerKeystore = JavaAgentHomeDir & "\conf\cacerts.jks"
	on error resume next
	If objFS0.FileExists(JavaAgentSourceControllerKeystore) Then
		objFS0.CopyFile JavaAgentSourceControllerKeystore, tmpUpgrFolder & "\cacerts.jks"
		JavaAgentSourceControllerKeystore = tmpUpgrFolder & "\cacerts.jks"
	End If
	on error goto 0

	' Backup existing agent deployment
	If InStr(JavaAgentHomeDir, "\ver") Then
		JavaAgentHomeDir = Left(JavaAgentHomeDir, InStr(JavaAgentHomeDir, "\ver")-1)
		objFs0.CopyFolder JavaAgentHomeDir, backupDirJava & "\JavaAgent-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	Else
		objFs0.CopyFolder JavaAgentHomeDir, backupDirJava & "\JavaAgent-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	End If
	
	' Get current version and clean up
	objFS0.CopyFile JavaAgentJAR, tmpUpgrFolder & "\javaagent.zip"
	call getJavaAgentDetails(tmpUpgrFolder & "\javaagent.zip", tmpUpgrFolder)
	JavaAgentCurrentVersion = sourceJavaAgentVersion
	objFS0.DeleteFile tmpUpgrFolder & "\javaagent.zip"
	objFS0.DeleteFolder tmpUpgrFolder & "\com"
	objFS0.DeleteFolder tmpUpgrFolder & "\META-INF"

	' Get new version and clean up
	Set objFolder = objFS0.GetFolder("binaries\JavaUpgrade")
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		UpgradeZIP = objFolder & "\" & objFile.name
		FileName = objFile.name
	Next
	set FilesInZip = objShell.NameSpace(UpgradeZIP).items
	objFS0.CreateFolder(tmpUpgrFolder & "\" & FileName)
	objShell.NameSpace(tmpUpgrFolder & "\" & FileName).CopyHere(FilesInZip)
	objFS0.CopyFile tmpUpgrFolder & "\" & FileName & "\javaagent.jar", tmpUpgrFolder & "\javaagent.zip"
	call getJavaAgentDetails(tmpUpgrFolder & "\javaagent.zip", tmpUpgrFolder)
	JavaAgentUpgradeToVersion = sourceJavaAgentVersion
	objFS0.DeleteFile tmpUpgrFolder & "\javaagent.zip"
	objFS0.DeleteFolder tmpUpgrFolder & "\com"
	objFS0.DeleteFolder tmpUpgrFolder & "\META-INF"

	' Compare current vs new version
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Comparing current (" & JavaAgentCurrentVersion & ") and new (" & JavaAgentUpgradeToVersion & ") version."
	If JavaAgentCurrentVersion >= JavaAgentUpgradeToVersion Then
    	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Version check failed."
		WScript.Quit
	Else
    	WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Version check succeeded."
	End If

	' Uninstall existing agent
	call getJavaAgentState(JavaAgentJAR, tmpUpgrFolder)
	objFS0.DeleteFolder(JavaAgentHomeDir)

	' Install new agent version and upgrade settings
	objFS0.CreateFolder(JavaAgentHomeDir)
	set FilesInZip = objShell.NameSpace(UpgradeZIP).items
	objShell.NameSpace(JavaAgentHomeDir).CopyHere(FilesInZip)
	If objFS0.FolderExists(JavaAgentHomeDir) Then
		Set objFolder = objFS0.GetFolder(JavaAgentHomeDir)
		If (objFolder.Files.Count = 0) Or (ObjFolder.SubFolders.Count = 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder already exists but is empty and it shouldn't be."
			WScript.Quit
		ElseIf (objFolder.Files.Count > 0) And (ObjFolder.SubFolders.Count > 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Folder check after unzip operation succeeded."
		End If
	End If
	Set objFolder = objFS0.GetFolder(JavaAgentHomeDir)
	Set colFolders = objFolder.SubFolders
	For Each colFolder in colFolders
		If InStr(colFolder, "\ver") Then
			JavaAgentHomeDir = colFolder
		End If
	Next
	objFS0.CopyFile tmpUpgrFolder & "\controller-info.xml", JavaAgentHomeDir & "\conf\controller-info.xml"
	If objFS0.FileExists(tmpUpgrFolder & "\cacerts.jks") Then
		objFS0.CopyFile tmpUpgrFolder & "\cacerts.jks", JavaAgentHomeDir & "\conf\cacerts.jks"
	End If
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Upgrade complete."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please update app start-up settings with the following path (below)."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Then start your application."
	WScript.Echo ""
	WScript.Echo "-javaagent:" & JavaAgentHomeDir & "\javaagent.jar"
	WScript.Echo ""
	objFS0.DeleteFolder("tmp\upgrade")

End Function

Function validateUpgrDir(UpgrDir, FileType)
	
	Set objFolder = objFS0.GetFolder(UpgrDir)
	Set colFiles = objFolder.Files
	For each colFile in colFiles
		If InStr(colFile, FileType) Then
			Result = "true"
		End If	
	Next
	If (objFolder.Files.Count = 1) And (Result = "true") Then
		WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Upgrade folder validated."
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Could not validate Upgrade folder."
		WScript.Echo FormatDateTime(Now, vbLongTime) &  " - Expected only 1 file of type " & FileType & "."
		WScript.Quit
	End If

End Function

Function Instrument_ja()

	' Get configuration properties
	call getConfigProperties()

	' Detect configuration
	Call getExistingInstall()

	' Select application, key, tier and node
	Call setSAASController()
	Call getSAASControllerDetails()
	Call setAppName()
	Call setAppTierName()
	Call setAppNodeName()

	' Support for autodeploy option
	If IsEmpty(autoAppType) Then
		
		' Select type of application and prepare for instrumentation
		WScript.Echo ""
		WScript.Echo "Available applications for monitoring: "
		WScript.Echo "JBoss | Tomcat | WebSphereAS | WebLogic"
		WScript.Echo ""
		WScript.StdOut.Write "Application to monitor: "
		PromptContinue = WScript.StdIn.ReadLine
		CustomerApplication = PromptContinue
	
	Else
	
		CustomerApplication = autoAppType
	
	End If

	If InStr("|JBoss|Tomcat|WebSphereAS|WebLogic|", "|" & CustomerApplication & "|") Then
	
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Selected monitoring for " & CustomerApplication & "."

	Else
	
		WScript.Echo ""
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not allowed: choose from one of the existing options."
		WScript.Quit
	
	End If

	If CustomerApplication = "JBoss" Then
		
		' Check if the right agent type is used
		If sourceJavaAgentType = "Sun" Then
		
		ElseIf sourceJavaAgentType = "IBM" Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Existing agent (" & sourceJavaAgentType & ") incompatible with this JVM type (Sun)."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Operation works for Java Agent of the same type."
			WScript.Quit
		
		End If

		' Support for autodeploy option
		If IsEmpty(autoAppStartScript) Then
		
			WScript.Echo ""
			WScript.StdOut.Write "Full path to application start-up script (ex: D:\jboss-as\bin\run.conf.bat): "
			PromptContinue = WScript.StdIn.ReadLine
			appStartUpScript = PromptContinue
		
		Else
		
			appStartUpScript = autoAppStartScript
		
		End If
		
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if start up script exists."
		
		If objFS0.FileExists(appStartUpScript) Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script exists. Creating backup of file."
			objFS0.CopyFile appStartUpScript, backupDirJava & "\" & objFs0.GetFileName(appStartUpScript) & "-" & DateDiff("s", "01/01/1970 00:00:00", Now())
		
		Else
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - File not found. Setup will quit."
			WScript.Quit
		
		End If
		
		Set objFile = objFS0.OpenTextFile(appStartUpScript, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.close
		
		For each Line in FileContents
		
			If (InStr(Line, "-Dappdynamics")) Or (InStr(Line, "-javaagent:")) Then
		
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up script must contain no instrumentation settings."
				WScript.Quit
		
			Else
		
			End If
		
		Next
	
	ElseIf CustomerApplication = "Tomcat" Then
		
		' Check if the right agent type is used
		If sourceJavaAgentType = "Sun" Then
		
		ElseIf sourceJavaAgentType = "IBM" Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Existing agent (" & sourceJavaAgentType & ") incompatible with this JVM type (Sun)."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Operation works for Java Agent of the same type."
			WScript.Quit
		
		End If

		' Select the right Tomcat instance for monitoring
		setTomcatJavaAgentMon()
		
		' Checking existing Tomcat configuration and backup of registry value
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Looking for Tomcat start up parameters."
		
		If WinArch = 32 Then
		
			strKeyPath = "HKLM\SOFTWARE\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java\Options"
		
		ElseIf WinArch = 64 Then
		
			strKeyPath = "HKLM\SOFTWARE\Wow6432Node\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java\Options"
		
		Else
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not determine Windows Server architecture (32/64)."
			WScript.Quit
		
		End If
		
		on error resume next
		wshShell.RegRead(strKeyPath)
		
		If err.number = 0 then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Found registry Key/Value for " & TomcatServiceName & "."
		
		Else
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Registry Key/Value not found  for " & TomcatServiceName & " (" & Err.Number & ")."
			WScript.Quit
		
		End If
		
		If WinArch = 32 Then
		
			strKeyPath = "SOFTWARE\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java"
		
		ElseIf WinArch = 64 Then
		
			strKeyPath = "SOFTWARE\Wow6432Node\Apache Software Foundation\Procrun 2.0\" & TomcatServiceName & "\Parameters\Java"
		
		Else
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not determine Windows Server architecture (32/64)."
			WScript.Quit
		
		End If
		
		on error goto 0
		ValueName = "Options"
		Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,ValueName,ArrTomcatOptions
		
		For each ArrTomcatOptions in ArrTomcatOptions
		
			If (InStr(ArrTomcatOptions, "-javaagent")) Or (InStr(ArrTomcatOptions, "-Dappdynamics")) Then 
		
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumentation already exists. Setup will quit."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Start up parameters must contain no instrumentation settings."
				WScript.Quit
		
			Else
		
			End If
		
		Next
		
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Creating registry key backup."
		WshShell.Run("cmd.exe /C regedit /E " & backupDirJava & "\tomcatJavaOptions" & "-" & DateDiff("s", "01/01/1970 00:00:00", Now()) & ".reg " & """" & "HKEY_LOCAL_MACHINE\" & strKeyPath & """"), 1, True
	
	ElseIf CustomerApplication = "WebSphereAS" Then
		
		' Check if the right agent type is used
		If sourceJavaAgentType = "IBM" Then
		
		ElseIf sourceJavaAgentType = "Sun" Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Existing agent (" & sourceJavaAgentType & ") incompatible with this JVM type (IBM)."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Operation works for Java Agent of the same type."
			WScript.Quit
		
		End If

		' Select the right WebSphereAS for monitoring, backup files and verify
		setWebsphereAppServerJavaAgentMon()

	ElseIf CustomerApplication = "WebLogic" Then

		' Check if the right agent type is used
		If sourceJavaAgentType = "Sun" Then
		
		ElseIf sourceJavaAgentType = "IBM" Then
		
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Existing agent (" & sourceJavaAgentType & ") incompatible with this JVM type (IBM)."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Operation works for Java Agent of the same type."
			WScript.Quit
		
		End If

		' Select the right WebSphereAS for monitoring, backup files and verify
		setWebLogicJavaAgentMon()
	
	Else
		
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a valid application selection."
		WScript.Quit
	
	End If	

	' Instrument selected application
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Instrumenting " & CustomerApplication & "."
	
	If CustomerApplication = "JBoss" Then 

		Set objFile = objFS0.OpenTextFile(appStartUpScript, ForAppending)
		objFile.WriteLine LineBuffer
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -javaagent:" & JavaAgentHomeDir & "\javaagent.jar"
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.agent.applicationName=" & ControllerApplicationName
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.agent.tierName=" & appTierName
		objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.agent.nodeName=" & appNodeName

		If proxyEnabled = "true" Then
			objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.http.proxyHost=" & proxyHost
			objFile.WriteLine "set JAVA_OPTS=%JAVA_OPTS% -Dappdynamics.http.proxyPort=" & proxyPort
		End If
		
		objFile.close

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		Set objFile = objFS0.OpenTextFile(appStartUpScript, ForReading)
		FileContents = objFile.ReadAll
	
		If InStr(FileContents, "-Dappdynamics.") Then
	
			If InStr(FileContents, "-javaagent:") Then
	
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."
	
			Else
	
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
				WScript.Quit
	
			End If
	
		Else
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application start up script is instrumented."
			WScript.Quit
	
		End If
	
		objFile.close

	ElseIf CustomerApplication = "Tomcat" Then
	
		ArrTomcatOptions = Array()
		oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,ValueName,ArrTomcatOptions
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-javaagent:" & JavaAgentHomeDir & "\javaagent.jar" 
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.agent.applicationName=" & ControllerApplicationName
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.agent.tierName=" & appTierName
		ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
		ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.agent.nodeName=" & appNodeName
		
		If proxyEnabled = "true" Then
			
			ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
			ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.http.proxyHost=" & proxyHost
			ReDim Preserve ArrTomcatOptions(UBound(ArrTomcatOptions) + 1)
			ArrTomcatOptions(UBound(ArrTomcatOptions)) = "-Dappdynamics.http.proxyPort=" & proxyPort
		
		End If

		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		oReg.SetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,ValueName,ArrTomcatOptions
	
		For each ArrTomcatOptions in ArrTomcatOptions
	
			If (InStr(ArrTomcatOptions, "-javaagent")) Or (InStr(ArrTomcatOptions, "-Dappdynamics")) Then 
	
				IsInstrumented = 1
	
			End If
	
		Next
	
		If IsInstrumented = 1 Then
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."
	
		Else
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
			WScript.Quit
	
		End If
	
	ElseIf CustomerApplication = "WebSphereAS" Then
		
		' Enable metrics collection from JVM via server.xml
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.Async = "false"
		xmlDoc.Load(WebSphereASServerXML)
	
		If WebSphereASstatisticSet = "none" Then
	
			Set xmlDocPMIService = xmlDoc.selectSingleNode(".//node()[@xmi:type = 'pmiservice:PMIService']")
			xmlDocPMIService.Attributes.getNamedItem("statisticSet").Text = "basic"
	
		End If

		' Instrument server.xml
		Set xmlDocJvmEntries = xmlDoc.selectSingleNode("//process:Server/processDefinitions/jvmEntries")
	
		If xmlDocJvmEntries is Nothing Then
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Couldn't find JVM params in server.xml."
			WScript.Quit
	
		Else
	
			JvmParams = xmlDocJvmEntries.Attributes.getNamedItem("genericJvmArguments").Text

			If proxyEnabled = "true" Then
		
				xmlDocJvmEntries.Attributes.getNamedItem("genericJvmArguments").Text = JvmParams & " -javaagent:" & JavaAgentHomeDir & "\javaagent.jar" & " -Dappdynamics.agent.applicationName=" & ControllerApplicationName & " -Dappdynamics.agent.tierName=" & appTierName & " -Dappdynamics.agent.nodeName=" & appNodeName & " -Dappdynamics.http.proxyHost=" & proxyHost & " -Dappdynamics.http.proxyPort=" & proxyPort

			ElseIf proxyEnabled = "false" Then
		
				xmlDocJvmEntries.Attributes.getNamedItem("genericJvmArguments").Text = JvmParams & " -javaagent:" & JavaAgentHomeDir & "\javaagent.jar" & " -Dappdynamics.agent.applicationName=" & ControllerApplicationName & " -Dappdynamics.agent.tierName=" & appTierName & " -Dappdynamics.agent.nodeName=" & appNodeName
		
			End If

		End If
	
		xmlDoc.Save(WebSphereASServerXML)
		Set xmlDoc = Nothing
		
		' Update server.policy
		If IsWebSphereASPolicyFileConfigured = 1 Then
	
		Else
	
			Set objFile = objFS0.OpenTextFile(WebSphereASPolicyFile, ForReading)
			strText = objFile.ReadAll
			objFile.Close
			'strText = strText & vbNewLine & "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentHomeDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
			strText = strText & vbNewLine & "// Allow AppDynamics Java Agent Monitoring" & vbNewLine & "grant codeBase ""file:" & Replace(JavaAgentInstallDir, "\", "/") & "/-"" {" & vbNewLine & "  permission java.security.AllPermission;" & vbNewLine & "};"
			Set objFile = objFS0.OpenTextFile(WebSphereASPolicyFile, ForWriting)
			objFile.Write strText
			objFile.Close
	
		End If
	
		' Verify instrumentation of server.xml
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Verifying."
		Set objFile = objFS0.OpenTextFile(WebSphereASServerXML, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.close
	
		For each Line in FileContents
	
			If (InStr(Line, "-Dappdynamics")) And (InStr(Line, "-javaagent:")) Then
	
				IsInstrumented = "1"
	
			Else
	
			End If
	
		Next
	
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.Async = "false"
		xmlDoc.Load(WebSphereASServerXML)
		Set xmlDocPMIService = xmlDoc.selectSingleNode(".//node()[@xmi:type = 'pmiservice:PMIService']")
		WebSphereASstatisticSet = UCase(xmlDocPMIService.Attributes.getNamedItem("statisticSet").Text)
	
		If IsInstrumented = 1 And WebSphereASstatisticSet <> UCase("none") Then
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Done."
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Please restart your application once feasible and apply some load."
	
		Else
	
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Failed to verify if application is instrumented."
			WScript.Quit
	
		End If
	
	ElseIf CustomerApplication = "WebLogic" Then
	
		call instrumentWebLogicJavaAgent()

	Else
	
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a valid application selection."
		WScript.Quit
	
	End if

End Function

Function Uninstall_ja()

	' Temp uninstall folder
	tmpUninstFolder = tmpFolder & "\uninstall"
	If objFS0.FolderExists(tmpUninstFolder) Then 
		Set objFolder = objFS0.GetFolder(tmpUninstFolder)
		objFS0.DeleteFolder(tmpUninstFolder)
		objFS0.CreateFolder(tmpUninstFolder)
	Else
		objFS0.CreateFolder(tmpUninstFolder)
	End If

	' Support for autodeploy option
	If IsEmpty(autoAgentPath) Then
		' Find existing installation and backup
		WScript.StdOut.Write "Please provide -javaagent value: "
		PromptContinue = WScript.StdIn.ReadLine
	Else
		PromptContinue = autoAgentPath
	End If
	JavaAgentJAR = Replace(PromptContinue, "-javaagent:", "")
	JavaAgentHomeDir = Replace(Replace(PromptContinue, "-javaagent:", ""), "\javaagent.jar", "")
	WScript.Echo JavaAgentHomeDir & " " & JavaAgentJAR
	If objFS0.FolderExists(JavaAgentHomeDir) Then
		Set objFolder = objFS0.GetFolder(JavaAgentHomeDir)
		If (objFolder.Files.Count > 0) And (ObjFolder.SubFolders.Count > 0) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Home folder exists and not empty."
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Home folder exists but empty."
			WScript.Quit
		End If
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Home folder does not exist."
		WScript.Quit
	End If
	If InStr(JavaAgentHomeDir, "\ver") Then
		JavaAgentHomeDir = Left(JavaAgentHomeDir, InStr(JavaAgentHomeDir, "\ver")-1)
		objFs0.CopyFolder JavaAgentHomeDir, backupDirJava & "\JavaAgent-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	Else
		objFs0.CopyFolder JavaAgentHomeDir, backupDirJava & "\JavaAgent-" & DateDiff("s", "01/01/1970 00:00:00", Now())
	End If

	' Uninstall agent
	call getJavaAgentState(JavaAgentJAR, tmpUninstFolder)
	objFS0.DeleteFile tmpUninstFolder & "\javaagent.jar"
	objFS0.DeleteFolder JavaAgentHomeDir
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Remove any AppDynamics start-up settings from your app."
	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Then restart your application."

End Function

Function getJavaAgentDetails(sourceFile, dstFolder)

	' Get Java Agent information such as type and version from a reliable source
	set FilesInZip = objShell.NameSpace(sourceFile).items
	objShell.NameSpace(dstFolder).CopyHere(FilesInZip)
	If objFS0.FileExists(dstFolder & "\META-INF\MANIFEST.MF") Then
		Set objFile = objFS0.OpenTextFile(dstFolder & "\META-INF\MANIFEST.MF", ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		objFile.Close
		For Each Line in FileContents
			If InStr(Line, "Implementation-Version") > 0 And InStr(Line, "Server IBM Agent") > 0 Then
				sourceJavaAgentType = "IBM"
			ElseIf InStr(Line, "Implementation-Version") > 0 And InStr(Line, "Server Agent") > 0 Then
				sourceJavaAgentType = "Sun"
			End If
			If InStr(Line, "Default-appagent-version") Then
				sourceJavaAgentVersion = Replace(Line, "Default-appagent-version: ", "")
			End If
		Next
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not find MANIFEST.MF."
		WScript.Quit
	End If

End Function

Function getJavaAgentState(sourceFile, dstFolder)

	WScript.Echo FormatDateTime(Now, vbLongTime) & " - Checking if Java Agent is not running."
	objFS0.CopyFile sourceFile, dstFolder & "\javaagent.jar"
	on error resume next
	Err.clear
	objFS0.DeleteFile sourceFile
	If err.number = 0 Then
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Java Agent is not running."
		objFS0.CopyFile dstFolder & "\javaagent.jar", sourceFile
	Else
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Check failed (" & err.number & "). Please stop monitored app and restart scenario."
		Err.clear
		WScript.Quit
	End If
	on error goto 0
	Err.clear

End Function

If args.count = 0 Then
	WScript.Echo "Type 'cscript deployAgent.vbs help' for list of commands."
End If

If args.count = 1 Then
	arg1 = WScript.Arguments(0)
	If arg1 = "help" Then
		PrintUsage()
	ElseIf arg1 = "autodeploy" Then

		call getConfigProperties()
		call getDrive()
		
		' Validate configuration
		If objFS0.FileExists("conf\autodeploy.xml") Then: autodeployXML = objFS0.GetFile("conf\autodeploy.xml"): WScript.Echo FormatDateTime(Now, vbLongTime) & " - Found autodeploy.xml.": Else: WScript.Echo FormatDateTime(Now, vbLongTime) & " - Could not find autodeploy.xml. Setup will quit.": WScript.Quit: End If
		
		' Look for available scenarios baesd on hostname and execute them
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Looking for scenario for " & strComputerName & "."
		Set objFile = objFS0.OpenTextFile(autodeployXML, ForReading)
		FileContents = Split(objFile.ReadAll, vbNewLine)
		Result = Array()

		' Handle commented lines
		WScript.Echo FormatDateTime(Now, vbLongTime) & " - Validating autodeploy.xml."
		For Each Line in FileContents: If InStr(Line, "!--") > 0 Then: Match = "true": End If: Next
		If Match = "true" Then: WScript.Echo FormatDateTime(Now, vbLongTime) & " - autodeploy.xml must contain no comments.": WScript.Echo FormatDateTime(Now, vbLongTime) & " - Remove any commented lines and restart the process.": WScript.Quit: Else: WScript.Echo FormatDateTime(Now, vbLongTime) & " - Validated.": End If

		' Main automation logic
		For Each Line in FileContents: If InStr(Line, "autoComputerName=""" & strComputerName & """") Then: ReDim Preserve Result(UBound(Result) + 1): Result(UBound(Result)) = Line: End If: Next
		If IsEmpty(Result) Then
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - No scenario found for " & strComputerName & "."
			WScript.Quit
		Else
			WScript.Echo FormatDateTime(Now, vbLongTime) & " - Found " & (UBound(Result) + 1) & " scenario(s) for " & strComputerName & "."
			For Each Result in Result
				RegExp.Pattern = "autoDeploymentType=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoDeploymentType=Replace(Replace(Match, "autoDeploymentType=", ""), """", ""): Next
				RegExp.Pattern = "autoAgentName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAgentName=Replace(Replace(Match, "autoAgentName=", ""), """", ""): Next
				RegExp.Pattern = "autoProdMon=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoProdMon=Replace(Replace(Match, "autoProdMon=", ""), """", ""): Next
				RegExp.Pattern = "autoAppName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppName=Replace(Replace(Match, "autoAppName=", ""), """", ""): Next
				WScript.Echo "--------------------"
				WScript.Echo FormatDateTime(Now, vbLongTime) & " - SCENARIO: " & autoDeploymentType & " (" & autoAgentName & ")"
				If (autoDeploymentType = "install") And (autoAgentName="dotnet") Then
					RegExp.Pattern = "autoDiscoIIS=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoDiscoIIS=Replace(Replace(Match, "autoDiscoIIS=", ""), """", ""): Next
					RegExp.Pattern = "autoRestIIS=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoRestIIS=Replace(Replace(Match, "autoRestIIS=", ""), """", ""): Next
					RegExp.Pattern = "autoStandMon=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoStandMon=Replace(Replace(Match, "autoStandMon=", ""), """", ""): Next
					RegExp.Pattern = "autoStandOpts=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoStandOpts=Replace(Replace(Match, "autoStandOpts=", ""), """", ""): Next
					call Install_dna()
				ElseIf (autoDeploymentType = "install") And (autoAgentName="java") Then
					RegExp.Pattern = "autoAppType=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppType=Replace(Replace(Match, "autoAppType=", ""), """", ""): Next
					RegExp.Pattern = "autoTierName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoTierName=Replace(Replace(Match, "autoTierName=", ""), """", ""): Next
					RegExp.Pattern = "autoNodeName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoNodeName=Replace(Replace(Match, "autoNodeName=", ""), """", ""): Next
					RegExp.Pattern = "autoAppJVM=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppJVM=Replace(Replace(Match, "autoAppJVM=", ""), """", ""): Next
					If autoAppType = "Tomcat" Then
						RegExp.Pattern = "autoTomcatServiceName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoTomcatServiceName=Replace(Replace(Match, "autoTomcatServiceName=", ""), """", ""): Next
					ElseIf autoAppType = "WebSphereAS" Then
						RegExp.Pattern = "autoAppStartScript=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppStartScript=Replace(Replace(Match, "autoAppStartScript=", ""), """", ""): Next
						RegExp.Pattern = "autoAppConf=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppConf=Replace(Replace(Match, "autoAppConf=", ""), """", ""): Next
					ElseIf autoAppType = "JBoss" Then
						RegExp.Pattern = "autoAppStartScript=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppStartScript=Replace(Replace(Match, "autoAppStartScript=", ""), """", ""): Next
					Else
						WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a valid application type selected (" & autoAppType & ")."
						WScript.Quit
					End If
					call Install_ja()
				ElseIf (autoDeploymentType = "install") And (autoAgentName="machine") Then
					RegExp.Pattern = "autoTierName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoTierName=Replace(Replace(Match, "autoTierName=", ""), """", ""): Next
					RegExp.Pattern = "autoNodeName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoNodeName=Replace(Replace(Match, "autoNodeName=", ""), """", ""): Next
					RegExp.Pattern = "autoSimEnabled=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoSimEnabled=Replace(Replace(Match, "autoSimEnabled=", ""), """", ""): Next
					RegExp.Pattern = "autoAppAgentExists=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppAgentExists=Replace(Replace(Match, "autoAppAgentExists=", ""), """", ""): Next
					call Install_ma()
				ElseIf (autoDeploymentType = "upgrade") And (autoAgentName="dotnet") Then
					call Upgrade_dna()
				ElseIf (autoDeploymentType = "upgrade") And (autoAgentName="machine") Then
					call Upgrade_ma()
				ElseIf (autoDeploymentType = "upgrade") And (autoAgentName="java") Then
					RegExp.Pattern = "autoAgentPath=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAgentPath=Replace(Replace(Match, "autoAgentPath=", ""), """", ""): Next
					call Upgrade_ja()
				ElseIf (autoDeploymentType = "instrument") And (autoAgentName="java") Then
					RegExp.Pattern = "autoAppType=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppType=Replace(Replace(Match, "autoAppType=", ""), """", ""): Next
					RegExp.Pattern = "autoTierName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoTierName=Replace(Replace(Match, "autoTierName=", ""), """", ""): Next
					RegExp.Pattern = "autoNodeName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoNodeName=Replace(Replace(Match, "autoNodeName=", ""), """", ""): Next
					RegExp.Pattern = "autoAppJVM=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppJVM=Replace(Replace(Match, "autoAppJVM=", ""), """", ""): Next
					If autoAppType = "Tomcat" Then
						RegExp.Pattern = "autoTomcatServiceName=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoTomcatServiceName=Replace(Replace(Match, "autoTomcatServiceName=", ""), """", ""): Next
					ElseIf autoAppType = "WebSphereAS" Then
						RegExp.Pattern = "autoAppStartScript=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppStartScript=Replace(Replace(Match, "autoAppStartScript=", ""), """", ""): Next
						RegExp.Pattern = "autoAppConf=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppConf=Replace(Replace(Match, "autoAppConf=", ""), """", ""): Next
					ElseIf autoAppType = "JBoss" Then
						RegExp.Pattern = "autoAppStartScript=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAppStartScript=Replace(Replace(Match, "autoAppStartScript=", ""), """", ""): Next
					Else
						WScript.Echo FormatDateTime(Now, vbLongTime) & " - Not a valid application type selected (" & autoAppType & ")."
						WScript.Quit
					End If
					call Instrument_ja()
				ElseIf (autoDeploymentType = "uninstall") And (autoAgentName="dotnet") Then
					call Uninstall_dna()
				ElseIf (autoDeploymentType = "uninstall") And (autoAgentName="machine") Then
					call Uninstall_ma()
				ElseIf (autoDeploymentType = "uninstall") And (autoAgentName="java") Then
					RegExp.Pattern = "autoAgentPath=""(.*?)""": Set Matches = RegExp.Execute(Result): For each Match in Matches: autoAgentPath=Replace(Replace(Match, "autoAgentPath=", ""), """", ""): Next
					call Uninstall_ja()
				End If
			Next
		End If

	Else
		WScript.Echo "Invalid parameters."
		WScript.Echo ""
		PrintUsage()
	End If
End If

If args.count = 2 Then
	arg1 = WScript.Arguments(0)
	arg2 = WScript.Arguments(1)
	If arg1 = "install" Then
		call getConfigProperties()
		If arg2 = "-dna" Then
			call getDrive()
			Install_dna()
		ElseIf arg2 = "-ja" Then
			call getDrive()
			Install_ja()
		ElseIf arg2 = "-ma" Then
			call getDrive()
			Install_ma()
		End If
	ElseIf arg1 = "uninstall" Then
		If arg2 = "-dna" Then
			Uninstall_dna()
		ElseIf arg2 = "-ma" Then
			Uninstall_ma()
		ElseIf arg2 = "-ja" Then
			Uninstall_ja()
		End If
	ElseIf arg1 = "upgrade" Then
		call getConfigProperties()
		'call getDrive()
		If arg2 = "-dna" Then
			call Upgrade_dna()
		ElseIf arg2 = "-ma" Then
			call Upgrade_ma()
		ElseIf arg2 = "-ja" Then
			call Upgrade_ja()
		End If
	ElseIf arg1 = "instrument" Then
		If arg2 = "-ja" Then
			call Instrument_ja()
		End If
	ElseIf arg1 = "healthcheck" Then
	
		call getConfigProperties()
		
		If arg2 = "-prod" Then
			url = ControllerUrlProd
			ApiAccessToken = ControllerTokenApiProd
			call getSAASControllerDetails()
		
		ElseIf arg2 = "-nonprod" Then
		
			url = ControllerUrlNonprod
			ApiAccessToken = ControllerTokenApiNonprod
			call getSAASControllerDetails()
		
		End If
	Else
		WScript.Echo "Invalid parameters."
		WScript.Echo ""
		PrintUsage()
	End If
	
End If
