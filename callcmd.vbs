Call main()

Sub main()
    'Renamed variables to cmd is your object and cmdline is your file path.
    Dim cmd, cmdline
    'Instantiate WshShell object
    Set cmd = Server.Createobject("WScript.Shell")
    'Set cmdline variable to file path
    cmdline = "c:\windows\system32\cscript.exe //nologo c:\s.vbs"
    'Execute Run and return immediately
    Call cmd.Run(cmdline, 0, False)
End Sub