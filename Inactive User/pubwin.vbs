dim objShell
set objShell=wscript.createObject("WScript.Shell")
set WSshell = createobject("wscript.shell")

iReturnCode=objShell.Run("""st.bat""  ",0,false)

WScript.Echo "hello"
