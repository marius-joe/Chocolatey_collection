' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Chocolatey Arguments
'  [-y]       Confirm all prompts - Chooses affirmative answer instead of prompting
'  [-r]       LimitOutput - Limit the output to essential information
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window
                                    
Const C_strSpace = "   "    ' space for text in msgBoxes

Dim objShell_A


	Set objShell_A = CreateObject("Shell.Application")

	programName = InputBox("Suche nach Programm :", "Chocolatey Search")
	If programName <> "" Then
		detailInfo = ""
        msgType = vbQuestion
		result = MsgBox (C_strSpace & "Detail Infos ?", vbYesNo + vbDefaultButton2 + msgType + vbMsgBoxSetForeground, "Chocolatey Search")
		If result = vbYes Then
			detailInfo = " --detailed"
		End If
		
		strCmd = "choco search " & qt_Cmd_PS(programName) & detailInfo
        
        keyToClose = "echo; & echo Close window: & PAUSE"
		objShell_A.ShellExecute "cmd", "/c " & qt(strCmd & " & " & keyToClose), "", "runas", C_RunWindowVisibility		' Run as admin
	End If
	
	Set objShell_A = Nothing

    

' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function

' Helper
Function qt_Cmd_PS(ByVal strValue)
    qt_Cmd_PS = Chr(34) & Chr(39) & strValue & Chr(39) & Chr(34)
End Function
