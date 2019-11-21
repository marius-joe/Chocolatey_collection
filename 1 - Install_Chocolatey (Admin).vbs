' ******************************************
'  Dev:  marius-joe
' ******************************************
'  v1.0.1
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window
Dim objShell_A


    isArgError = False
    isSilentMode = False
    arrComplexArgs = parseArguments()
    If Not isEmpty_ArrList(arrComplexArgs) Then
        For Each complexArg In arrComplexArgs
			argument = LCase(complexArg(0))
            value = arr_SafeGet(complexArg, 1, "")
            
            Select Case argument
				Case "-silent"
					isSilentMode = True 
            End Select
        Next
    Else
        isArgError = True
    End If
    
    If Not isArgError Then
        Set objShell_A = CreateObject("Shell.Application")

        ' Installation with cmd.exe is better than with PowerShell.exe for compability with Windows 7 without the .NET Framework
        strCmd = "@""%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command " & _
                 """iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))""" & _
                 "&& SET ""PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"""
        
        If Not isSilentMode Then
            keyToClose = "echo; & echo Close window: & PAUSE"
            windowVisibility = C_RunWindowVisibility
        Else
            keyToClose = ""
            windowVisibility = 0
        End If
        objShell_A.ShellExecute "cmd", "/c " & qt(strCmd & " & " & keyToClose), "", "runas", windowVisibility		' Run as admin
        Set objShell_A = Nothing
    End If


    
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