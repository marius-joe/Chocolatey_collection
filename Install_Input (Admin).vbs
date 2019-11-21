' ******************************************
'  Dev:  marius-joe
' ******************************************
'  v1.0.1
' ******************************************
'  Chocolatey Arguments
'  [-y]       Confirm all prompts - Chooses affirmative answer instead of prompting
'  [-r]       LimitOutput - Limit the output to essential information
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window

Dim objShell_A
Dim objShell_WS


	Set objShell_A = CreateObject("Shell.Application")
    Set objShell_WS = CreateObject("WScript.Shell")	

	programName = InputBox("Programme zur Installation angeben :" & vbCrlf & "(getrennt durch Leerzeichen):", "Chocolatey Installer")
	If programName <> "" Then
    
        ' check for working Chocolatey installation
        On Error Resume Next
            objShell_WS.Run "choco", 0, true
            isChoco = (Err.Number = 0)
        On Error GoTo 0    
    
        strCmd1 = "choco upgrade chocolatey -y -r"
        strCmd2_newWindow = "choco install " & programName & " -y -r"
        
        ' only if choco update is successful, start a new shell and run further actions on the new choco version
        keyToClose = "echo; & echo Close window: & PAUSE"
        strCmd = strCmd1 & " && echo; && start """" /b cmd /c """ & strCmd2_newWindow & " & " & keyToClose & """"
        objShell_A.ShellExecute "cmd", "/k " & qt(strCmd & " && EXIT"), "", "runas", C_RunWindowVisibility		' Run as admin
	End If
	
	Set objShell_A = Nothing



' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function
