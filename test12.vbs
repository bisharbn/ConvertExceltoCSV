Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)

WScript.Echo "Hello"
WScript.StdOut.Write "Hello"
WScript.StdOut.WriteLine "Hello"
Stdout.WriteLine "Hello"
Stdout.Write "Hello"