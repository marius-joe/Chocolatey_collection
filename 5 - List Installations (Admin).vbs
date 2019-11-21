' ******************************************
'  Dev:  marius-joe
' ******************************************

Const C_RunWindowVisibility = 5		' 5 - Open the application with its window at its current size and position
									' 0 - Open the application with a hidden window

Dim objShell_A


	Set objShell_A = CreateObject("Shell.Application")

	strCmd = "choco list " & "--local-only"
    
    keyToClose = "echo; & echo Close window: & PAUSE"
	objShell_A.ShellExecute "cmd", "/c " & qt(strCmd & " & " & keyToClose), "", "runas", C_RunWindowVisibility		' Run as admin
	
	Set objShell_A = Nothing


    
' Helper
Function qt(ByRef strValue)  ' v1.2
    qt = Chr(34) & strValue & Chr(34)
End Function
