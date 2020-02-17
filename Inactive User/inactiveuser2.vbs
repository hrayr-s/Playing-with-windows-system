' detect_user_idle.vbs
' purpose: GFI RMM user script to detect if user is active/idle (e.g. before Take Control)
' author: Rob Eberhardt, www.slingfive.com
' versions: 2013-12-26  first


OPTION EXPLICIT
'wscript.echo "sleep 3 seconds..."
'WScript.Sleep(3000)
WScript.echo "As of " & Now() & " :"


DIM bLoginShowing, bScrShowing

DIM strCmd, objShell, objExecObject, strText
Set objShell = WScript.CreateObject("WScript.Shell")


'check welcome screen
strCmd = "tasklist /FI ""IMAGENAME eq LogonUI.exe"""
Set objExecObject = objShell.Exec(strCmd)
Do While Not objExecObject.StdOut.AtEndOfStream
    strText = strText & vbCRLF & objExecObject.StdOut.ReadLine()
Loop
If Instr(strText, "LogonUI.exe") >0 Then
    call WScript.Echo("* Login Screen currently showing (...could be using Remote Desktop)")
    bLoginShowing = true
End If
strText=""


'check screensaver
strCmd = "tasklist /v /fo table /nh"
Set objExecObject = objShell.Exec(strCmd)
Do While Not objExecObject.StdOut.AtEndOfStream
    strText = strText & vbCRLF & objExecObject.StdOut.ReadLine()
Loop
If Instr(strText, ".scr ") >0 Then
    call WScript.Echo("* Screensaver currently running")
    bScrShowing = true
End If
strText=""


'check idle time
strCmd = "quser"
Set objExecObject = objShell.Exec(strCmd)
objExecObject.StdOut.ReadLine() ' skip output headers
'TODO: handle multiple users    <-------
Do While Not objExecObject.StdOut.AtEndOfStream
    strText = strText & vbCRLF & objExecObject.StdOut.ReadLine()
Loop
DIM regex   ' turn spaces into pipes, split into array, and pull into vars
    SET regex  = New RegExp
    regex.Pattern = "[ ]{2,}"
    regex.Global = true
    strText = regex.replace(strText, "|")
    DIM strUserName, strSessionName, strID, strState, strIdleTime, strLogonTime
    strUserName = split(strText, "|")(0)
        strUserName = replace(strUserName, ">", "") ' > indicates current user
        strUserName = replace(strUserName, vbCR, "")
        strUserName = replace(strUserName, vbLF, "")
    strSessionName = split(strText, "|")(1)
    strID = split(strText, "|")(2)
    strState = split(strText, "|")(3)
    strIdleTime = split(strText, "|")(4)
    strLogonTime = split(strText, "|")(5)
wscript.echo "* " & strUserName & " " & strState & ", time: " & strIdleTime
strText=""


IF NOT (bLoginShowing or bScrShowing) THEN
    call WScript.Echo("Summary ==> No Login screen or Screensaver running.  Presume that the user is active.")
    'WScript.Quit(-1)   ' causes red X in dashboard?
ELSE
    WScript.Quit(0)
END IF
