'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Printer List and Check for a computer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'On Error Resume Next

Dim strComputer, strExcelPath, objExcel, k, objGroup
Dim objWMIService, colItems, ErrState, Sheet, strPorts, strPortNum

strComputer = Ucase(InputBox ("Please enter the computer name to check.", "Computer Name:","MWxxxSSI"))
strDate = Replace(Replace(Date(),"/","-"),":",".")
strTime = Replace(Replace(Time(),"/","-"),":",".")

if strComputer = "" then
  WScript.quit
end if 

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

' Create a new workbook.
objExcel.Workbooks.Add

Select Case UCase(strComputer)
  Case Else
    PrintServer(strComputer)
End Select

Function PrintServer(strComputer)

k = 3

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer",,48)
Set colPorts =  objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")
Set strPorts = CreateObject("Scripting.Dictionary")
Set strPortNum = CreateObject("Scripting.Dictionary")

k = 4

For Each objPort in colPorts
strPorts.add objPort.Name, objPort.HostAddress
strPortNum.add objPort.Name, objPort.PortNumber
Next

' Populate spreadsheet cells with printer attributes.
objExcel.Cells(1, 1).Value = "Computer: " & strComputer & ",  Date/Time: " & strdate & " @ " & strtime 
objExcel.Cells(2, 1).Value = "Name"
objExcel.Cells(2, 2).Value = "ShareName"
objExcel.Cells(2, 3).Value = "Comment"
objExcel.Cells(2, 4).Value = "Error"
objExcel.Cells(2, 5).Value = "DriverName"
objExcel.Cells(2, 6).Value = "EnableBIDI"
objExcel.Cells(2, 7).Value = "JobCount"
objExcel.Cells(2, 8).Value = "Location"
objExcel.Cells(2, 9).Value = "PortName"
objExcel.Cells(2, 10).Value = "PortAddress"
objExcel.Cells(2, 11).Value = "PortNumber"
objExcel.Cells(2, 12).Value = "Published"
objExcel.Cells(2, 13).Value = "Queued"
objExcel.Cells(2, 14).Value = "Shared"
objExcel.Cells(2, 15).Value = "Status"

For Each objItem in colItems
'put error code into human readable form
Select Case objItem.DetectedErrorState
  Case 4
    ErrState = "Out of Paper"
  Case 5
    ErrState = "Toner low"
  Case 6
    ErrState = "Printing"
  Case 9
    ErrState = "Offline"
  Case Else
    ErrState = objItem.DetectedErrorState
End Select

'populate the row with this printer's data
objExcel.Cells(k, 1).Value = objItem.Name
objExcel.Cells(k, 2).Value = objItem.ShareName
objExcel.Cells(k, 3).Value = objItem.Comment
objExcel.Cells(k, 4).Value = ErrState
objExcel.Cells(k, 5).Value = objItem.DriverName
objExcel.Cells(k, 6).Value = objItem.EnableBIDI
objExcel.Cells(k, 7).Value = objItem.JobCountSinceLastReset
objExcel.Cells(k, 8).Value = objItem.Location
objExcel.Cells(k, 9).Value = objItem.PortName
objExcel.Cells(k, 10).Value = strPorts.Item(objItem.PortName)
objExcel.Cells(k, 11).Value = strPortNum.Item(objItem.PortName)
objExcel.Cells(k, 12).Value = objItem.Published
objExcel.Cells(k, 13).Value = objItem.Queued
objExcel.Cells(k, 14).Value = objItem.Shared
objExcel.Cells(k, 15).Value = objItem.Status

k = k + 1
Next

' Format the spreadsheet.
objExcel.Range("A1:O1").Font.Bold = True
'objExcel.Select
objExcel.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

objExcel.Range("A2:O2").Select
objExcel.Selection.Interior.ColorIndex = 19
objExcel.Selection.Font.ColorIndex = 11
objExcel.Selection.Font.Bold = True
objExcel.Cells.EntireColumn.AutoFit
End Function
 

Set objExcel = objExcel.ActiveWorkbook.Worksheets(1)
Set objRange = objExcel.Range("O1")

'objRange.Sort objRange,2,,,,,,1

 

MsgBox "Done"
