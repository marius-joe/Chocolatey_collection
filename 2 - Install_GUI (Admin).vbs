' ******************************************
'  Dev:  marius-joe
' ******************************************
'  Chocolatey Arguments
'  [-y]       Confirm all prompts - Chooses affirmative answer instead of prompting
'  [-r]       LimitOutput - Limit the output to essential information
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window

Dim objShell_A


	Set objShell_A = CreateObject("Shell.Application")
    
    strCmd1 = "choco upgrade chocolatey -y -r"
    strCmd2_newWindow = "choco install chocolateygui -y -r"   
    
    ' only if choco update is successful, start a new shell and run further actions on the new choco version
    keyToClose = "echo; & echo Close window: & PAUSE"
    strCmd = strCmd1 & " && echo; && start """" /b cmd /c """ & strCmd2_newWindow & " & " & keyToClose & """"
    objShell_A.ShellExecute "cmd", "/k " & qt(strCmd & " && EXIT"), "", "runas", C_RunWindowVisibility		' Run as admin
    
	Set objShell_A = Nothing


    
' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function
