'Created by Bishar for converting excel to CSV from a folder

'Program Start here
Set oWSH = CreateObject("WScript.Shell")
 vbsInterpreter = "cscript.exe"
 
 dim sSourceFolder
 dim sDestinationFolder
 
 ' Configure the source and destination folder below
 sSourceFolder = "C:\Data\VBA\Test VBS\excel\"  'Set the Source folder to get the excel files
 sDestinationFolder= "C:\Data\VBA\Test VBS\excel\"  'Set the destination folder to convert to .csv
 
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Read files from the source folder check for only xls file
For Each oFile In oFSO.GetFolder(sSourceFolder).Files
  If UCase(oFSO.GetExtensionName(oFile.Name)) = "XLS" Then
     ConvertExcel oFile
  End if
Next

Set oFSO = Nothing

'Program End here


'Function to convert excel files to csv
Function ConvertExcel(ExcelFilename)

		Set obj = createobject("Excel.Application")  'Creating an Excel Object
		'obj.visible=True                                   'Making an Excel Object visible
		Set oFSO1 = CreateObject("Scripting.FileSystemObject")  
     '	WScript.Echo oFSO1.GetBaseName(ExcelFilename.path)
	Dim  filename
	 filename=  oFSO1.GetBaseName(ExcelFilename.path)
	'WScript.Echo filename
	Dim ext
	ext =oFSO1.GetExtensionName(ExcelFilename.path)

	'WScript.Echo sSourceFolder
	
	Dim path
	Dim pathfilesaveas	
	path= sSourceFolder + filename + "." + ext
	pathfilesaveas = sDestinationFolder + oFSO1.GetBaseName(ExcelFilename.path) +".csv"

	'WScript.Echo path

	Set obj1 = obj.Workbooks.open(path)    'Opening an Excel file
	Set obj2=obj1.Worksheets("Sheet1")    'Referring Sheet1 of excel file
	'Msgbox obj2.Cells(2,2).Value  'Value from the specified cell will be read and shown
	obj1.SaveAs pathfilesaveas , 6
	obj1.Close                                             'Closing a Workbook
	obj.Quit                                                  'Exit from Excel Application
	Set obj1=Nothing                                 'Releasing Workbook object
	Set obj2 = Nothing                               'Releasing Worksheet object
	Set obj=Nothing                                   'Releasing Excel object
 
 End Function

