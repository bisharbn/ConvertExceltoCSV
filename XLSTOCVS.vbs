Dim vExcel
Dim vCSV
Set vExcel = CreateObject("Excel.Application")
Set vCSV = vExcel.Workbooks.Open(Wscript.Arguments.Item(0))
WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial04.xlsx" & vbcrlf)
vCSV.SaveAs WScript.Arguments.Item(0) & ".csv", 6
vCSV.Close False
vExcel.Quit
WScript.Echo "Done"