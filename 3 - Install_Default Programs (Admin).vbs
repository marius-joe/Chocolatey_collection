' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Chocolatey Arguments
'  [-y]       Confirm all prompts - Chooses affirmative answer instead of prompting
'  [-r]       LimitOutput - Limit the output to essential information
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window

Const C_File_DefaultPrograms = "\Default Programs.ini"
Const C_Encoding_ASCII = 0

Dim objShell_A
Dim objFSO


    Call setGlobalsIfNecessary("objShell_A, objFSO")
	
	strParentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
	filePath = objFSO.BuildPath(strParentFolder, C_File_DefaultPrograms)

	' load "Default Programs.ini"
	strPrograms = LoadProgramString(filePath)
    
    strCmd1 = "choco upgrade chocolatey -y -r"
    strCmd2_newWindow = "choco install " & strPrograms & " -y -r"
    
    ' only if choco update is successful, start a new shell and run further actions on the new choco version
    keyToClose = "echo; & echo Close window: & PAUSE"
    strCmd = strCmd1 & " && echo; && start """" /b cmd /c """ & strCmd2_newWindow & " & " & keyToClose & """"
    objShell_A.ShellExecute "cmd", "/k " & qt(strCmd & " && EXIT"), "", "runas", C_RunWindowVisibility		' Run as admin
        
	Call cleanGlobals("All")
 


Function LoadProgramString(ByRef strPath)  ' v1.3
    mark_comment = "#"
    strPrograms = ""
	arrLines = get_TextLines(strPath, C_Encoding_ASCII)
    If Not isEmpty_ArrList(arrLines) Then
		For Each strLine In arrLines
			strLine = Trim(strLine)
			If strLine <> "" Then
				If Not StartsWith(strLine, mark_comment) Then
                    ' solution with ArrayList + Join would be easier, but requires .NET 3.5
					If strPrograms = "" Then strSpace = "" Else strSpace = " "
					strPrograms = strPrograms & strSpace & strLine
				End If
			End If
		Next
	End If
	LoadProgramString = strPrograms
End Function

				
' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function


' Helper Bundle
' ----------------------------------------------------
Function read_File(ByRef strPath, ByRef encoding)  ' v1.4
    functionName = "read_File"
	Call setGlobalsIfNecessary("objFSO")
    If objFSO.FileExists(strPath) Then
        On Error Resume Next
            Set objFile = objFSO.GetFile(strPath)
            
            If Err.Number = 0 Then
                Set objTSO = objFile.OpenAsTextStream(1, encoding)
                read_File = objTSO.Read(objFile.Size)
                Set objTSO = Nothing
            Else
                Err.Clear
                read_File = ""
            End If
            
            Set objFile = Nothing
        On Error GoTo 0
    Else
        Call show_MsgBox("File not found: " & strPath, vbCritical, "Function: " & functionName)
    End If
End Function

' Helper
Function get_TextLines(ByRef strPath, ByRef encoding)  ' v1.3
    fileText = read_File(strPath, encoding)
	If fileText <> "" Then
		get_TextLines = Split(fileText, vbCrlf)
	Else
		get_TextLines =  Array()
	End If
End Function

' Helper
Sub show_MsgBox(ByRef msg_ArrOrStr, ByRef msgType, ByRef strTitle)  ' v1.2
    strSpace = "   "    ' space for text in msgBoxes
    If IsArray(msg_ArrOrStr) Then
        strMsg = Join(msg_ArrOrStr, vbCrlf & strSpace)
    Else
        strMsg = msg_ArrOrStr
    End If
	MsgBox strSpace & strMsg, msgType + vbMsgBoxSetForeground, strTitle
End Sub
' ----------------------------------------------------


' Helper Bundle  v1.4
' ----------------------------------------------------
Sub setGlobalsIfNecessary(ByRef strObjectNames)
	arrObjectNames = Split(strObjectNames, ",")
	For Each strName In arrObjectNames
		strObj = UCase(Trim(strName))
		If strObj = UCase("objShell_A") Then
			If IsEmpty(objShell_A) Then Set objShell_A = CreateObject("Shell.Application")
		
		ElseIf strObj = UCase("objShell_WS") Then
			If IsEmpty(objShell_WS) Then Set objShell_WS = CreateObject("WScript.Shell")			
		
		ElseIf strObj = UCase("objFSO") Then
			If IsEmpty(objFSO) Then Set objFSO = CreateObject("Scripting.FileSystemObject")		
		End If
	Next
End Sub

Sub cleanGlobals(ByRef strObjectNames)
    If UCase(strObjectNames) = "ALL" Then
        arrObjectNames = Array("objShell_A", "objShell_WS", "objFSO")
    Else
        arrObjectNames = Split(strObjectNames, ",")
    End If
    
    For Each strName In arrObjectNames
        strObj = UCase(Trim(strName))
        If strObj = UCase("objShell_A") Then
            If Not IsEmpty(objShell_A) Then Set objShell_A = Nothing
        
        ElseIf strObj = UCase("objShell_WS") Then
            If Not IsEmpty(objShell_A) Then Set objShell_WS = Nothing
        
        ElseIf strObj = UCase("objFSO") Then
            If Not IsEmpty(objShell_A) Then Set objFSO = Nothing
        End If
    Next
End Sub
' ----------------------------------------------------

' Helper
Function StartsWith(ByRef str, ByRef start)  ' v1.1
	Dim startLength
	startLength = Len(start)
	StartsWith = (Left(Trim(UCase(str)), startLength) = UCase(start))
End Function

' Helper
Function isEmpty_ArrList(ByRef arrOrList)	' v1.6
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
        Call show_MsgBox("Variable 'arrOrList' is no Array or ArrayList: " & TypeName(arrOrList), vbCritical, "Function: " & functionName)
    End If
	
	isEmpty_ArrList = returnValue
End Function
