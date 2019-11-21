' ******************************************
'  Dev:  marius-joe
' ******************************************
'  VBScript slim CLI for the Chocolatey Package Manager
'  v1.0.3
' ******************************************
'  Arguments:
'  [-install pkg1 pkg2 pkg3 ...]
'  [-uninstall pkg1 pkg2 pkg3 ...]
'  [-update]
'  [-upgrade pkg1 pkg2 pkg3 ...]
'  [-list]
'  [-silent]
' ******************************************


Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window

Dim objFSO
Dim objShell_WS
Dim objShell_A


    ' Installation with cmd.exe is superior to one with PowerShell.exe (compability with Windows 7 without the .NET Framework)
    cmd_choco_setup = "@""%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " & _
                        """iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))""" & _
                        "&& SET ""PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"""
         
    cmd_choco_upgrade = "choco upgrade chocolatey -y -r"
    
    pgks_placerholder = "[packages]"
    cmd_pgks_upgrade   = "choco upgrade [packages] -y -r --except=""chocolatey"""
    cmd_pgks_install   = cmd_pgks_upgrade   ' "upgrade" installs missing pgks + updates already installed pgks
    cmd_pgks_uninstall = "choco uninstall [packages] -y -r"
    cmd_pgks_list      = "choco list --local-only"

    ' choco commands need to be processed in an efficient order
    ' a needed choco self-update is always performed before all other commands
    index_pgks_upgrade = 1
    index_pgks_install = 2
    index_pgks_uninstall = 3
    index_pgks_list = 4
             
    cmd_keyToClose = "echo; & echo Close window: & PAUSE"
    
    ' toDo

    runUnattended = True
	arrComplexArgs = parseArguments()
    isArgError = False
	If Not isEmpty_ArrList(arrComplexArgs) Then
        Call setGlobalsIfNecessary("objShell_WS, objShell_A")

        ' check for working Chocolatey installation
        On Error Resume Next
            objShell_WS.Run "choco", 0, true
            isChoco = (Err.Number = 0)
        On Error GoTo 0
        
        isSilentMode = False
        ensureNewChoco = False
        Dim arrCmdsOrdered(3)
        For Each complexArg In arrComplexArgs
			argument = LCase(complexArg(0))
            value = arr_SafeGet(complexArg, 1, "") 
            
            cmd_choco = ""
            choco_mode = ""
            Select Case argument
             'Positionen: -init -upgrade -install "blub" -uninstall "yolo" -list
           

                ' Position: 3
                Case "-install"
                    ' install new packages or upgrade them if already installed
                    If value <> "" Then
                        packages = value
                        cmd_choco = Replace(cmd_pgks_install, pgks_placerholder, packages)
                        ensureNewChoco = True
                        arrCmdsOrdered(index_pgks_install-1) = cmd_choco
                    Else 
                        isArgError = True
                    End If

                    'Call gleiches wie -upgrade mit packages aber auch config file
                    ' https://gitlab.com/luukgrefte/choco-autoinstalller/blob/master/defaultapps.config
                    ' choco install %currentpath%\defaultapps.config
                    
' kann eigenlich alles kombiniert werden, manchmal muss dann choco upgrade wegfallen, wenn im ersten comand schon enthalten
                    
                ' Position: 4
                Case "-uninstall"
                    If value <> "" Then
                        packages = value
                        cmd_choco = Replace(cmd_pgks_uninstall, pgks_placerholder, packages)
                        arrCmdsOrdered(index_pgks_uninstall-1) = cmd_choco
                    Else
                        isArgError = True
                    End If
                    

                ' Position: 1
                Case "-update"
                    ' upgrades or installs only Chocolatey itself
                    ensureNewChoco = True

                    
                ' Position: 2
                Case "-upgrade"
                    ' includes "-update": uprades or installs Chocolatey itself
                    ' upgrades all installed packages or only those specified after the command (-upgrade "myPkg1 myPkg2")
                    ' installs packages if you don't have them already
                    If value <> "" Then
                        packages = value
                    Else
                        packages = "all"
                    End If
                    cmd_choco = Replace(cmd_pgks_upgrade, pgks_placerholder, packages)
                    ensureNewChoco = True
                    arrCmdsOrdered(index_pgks_upgrade-1) = cmd_choco
                    
                    
                'Position als letztes
                Case "-list"
                    ' list all installed packages
                    cmd_choco = cmd_pgks_list
                    arrCmdsOrdered(index_pgks_list-1) = cmd_choco

                    
                Case "-silent"
                    isSilentMode = True
            End Select
        Next
    Else
        isArgError = True
    End If

    If Not isArgError Then
        If isSilentMode Then
            optKeyToClose = ""
            windowVisibility = 0
        Else
            optKeyToClose = cmd_keyToClose
            windowVisibility = C_RunWindowVisibility
        End If
    
        cmds_collected = ""
        For Each cmd_choco In arrCmdsOrdered
            ' collect the ordered choco commands
            If cmd_choco <> "" Then
                If cmds_collected = "" Then 
                    cmds_collected = cmd_choco
                Else
                    cmds_collected = cmds_collected & " && " & cmd_choco
                End If
            End If  
        Next
        
MsgBox "cmds_collected:  " & cmds_collected 
           
        If ensureNewChoco Then
            If isChoco Then
                cmd_prep = cmd_choco_upgrade
            Else
                cmd_prep = cmd_choco_setup
            End If
            
            ' only if choco upgrade is successful, reload the shell
            ' and run possible further actions on the new choco version (recommended by the Chocolatey team)
            If cmds_collected <> "" Then
                cmd_newShell = "start """" /b cmd /c """ & cmds_collected & " & " & optKeyToClose & """"  
                cmd_main = "/k " & qt(cmd_prep & _
                                     " && echo; && " & _
                                     cmd_newShell & _
                                     " && EXIT")
            Else
                cmd_main = "/c " & qt(cmd_prep & " & " & optKeyToClose)
            End If
        Else
            If isChoco Then
                cmd_main = "/c " & qt(cmds_collected & " & " & optKeyToClose)
            Else
                echo "Chocolatey is not installed"
            End If   
        End If
        
MsgBox "cmd_main:  " & cmd_main 
        objShell_A.ShellExecute "cmd", cmd_main, "", "runas", windowVisibility		' Run as admin
    Else
        MsgBox "args sind kacke"
    End If

	Call cleanGlobals("All")


    
' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function


' Helper Bundle - v1.6
' ----------------------------------------------------
Function parseArguments()
    Dim arrComplexArgs()
	Set objArgs = WScript.Arguments
	countArgs = objArgs.Count
	If countArgs > 0 Then
		strArgs = " """ & objArgs(0)
        ' read all arguments (seperated by " " by default)
        ' mark each beginning of an argument part
        ' " can quite safely be used for that, because the " got removed when args were handed to the script
		For i = 1 To countArgs-1
			strArgs = strArgs & " """ & objArgs(i)
		Next
		If Contains(strArgs, " ""-") Then
			arrCorrectArgs = Split(strArgs, " ""-")
			UBarrCorrectArgs = UBound(arrCorrectArgs)
            ReDim arrComplexArgs(UBarrCorrectArgs-1)
			For i = 1 To UBarrCorrectArgs
				arrArgument = Split(Trim(arrCorrectArgs(i)), " """, 2)
				arrArgument(0) = "-" & arrArgument(0)
                If UBound(arrArgument) > 0 Then arrArgument(1) = arrArgument(1)
                arrComplexArgs(i-1) = arrArgument
			Next
		End If
	End If
	parseArguments = arrComplexArgs
End Function

' v1.3
Function Contains(ByRef str, ByRef strSearch)
	' converting to lower case is better than vbTextCompare because of dealing with foreign languages
	If InStr(LCase(str), LCase(strSearch)) > 0 Then returnValue = True Else returnValue = False
	Contains = returnValue
End Function

' v1.9
Function isEmpty_ArrList(ByRef arrOrList)
    functionName = "isEmpty_ArrList"
	returnValue = True
	If IsArray(arrOrList) Then		' is array
		On Error Resume Next
			UBarr = UBound(arrOrList)
			If (Err.Number = 0) And (UBarr >= 0) Then returnValue = False
		On Error GoTo 0
	ElseIf TypeName(arrOrList) = "ArrayList" Then	 ' is list
        If arrOrList.Count > 0 Then
            returnValue = False
        End If
    Else
        returnValue = "Variable 'arrOrList' is no Array or ArrayList: " & TypeName(arrOrList)
    End If
	isEmpty_ArrList = returnValue
End Function

' v1.1
Function arr_SafeGet(ByRef arr, ByRef index, ByRef defaultValue)
    If UBound(arr) >= index Then arr_SafeGet = arr(index) Else arr_SafeGet = defaultValue
End Function
' ----------------------------------------------------


' Helper Bundle - v1.5
' req: Dim objShell_A, objShell_WS, objFSO
' ----------------------------------------------------
Sub setGlobalsIfNecessary(ByRef strObjectNames)
	arrObjectNames = Split(strObjectNames, ",")
	For Each strName In arrObjectNames
		strObj = LCase(Trim(strName))
		If strObj = LCase("objShell_A") Then
			If IsEmpty(objShell_A) Then Set objShell_A = CreateObject("Shell.Application")
		
		ElseIf strObj = LCase("objShell_WS") Then
			If IsEmpty(objShell_WS) Then Set objShell_WS = CreateObject("WScript.Shell")			
		
		ElseIf strObj = LCase("objFSO") Then
			If IsEmpty(objFSO) Then Set objFSO = CreateObject("Scripting.FileSystemObject")		
		End If
	Next
End Sub

Sub cleanGlobals(ByRef strObjectNames)
    If LCase(strObjectNames) = "all" Then
        arrObjectNames = Array("objShell_A", "objShell_WS", "objFSO")
    Else
        arrObjectNames = Split(strObjectNames, ",")
    End If
    
    For Each strName In arrObjectNames
        strObj = LCase(Trim(strName))
        If strObj = LCase("objShell_A") Then
            If Not IsEmpty(objShell_A) Then Set objShell_A = Nothing
        
        ElseIf strObj = LCase("objShell_WS") Then
            If Not IsEmpty(objShell_A) Then Set objShell_WS = Nothing
        
        ElseIf strObj = LCase("objFSO") Then
            If Not IsEmpty(objShell_A) Then Set objFSO = Nothing
        End If
    Next
End Sub
' ----------------------------------------------------