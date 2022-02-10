'On Error Resume Next

Dim strComputer, strExcelPath, objExcel, objSheet, k, objGroup
Dim objWMIService, colItems, ErrState, Sheet, strPorts, strPortNum

'Sheet = spreadsheet page, k = row in sheet
Sheet = 1
k = 2

strComputer = Ucase(InputBox ("Please enter the computer name to check.", "Computer Name:","MWxxxSSI"))
StrDate = Replace(Replace(Now(),"/","-"),":",".")

if strComputer = "" then
  WScript.quit
end if

'strExcelPath = InputBox ("Please enter the path to save the file to: " & VbCrLf & VbCrLF _
'& "(NOTE:  The filename should be followed by a \)", "File path", "C:\")
strExcelPath = "PrinterCheckLog\" & StrDate & "-Printers for..." & strComputer & ".xlsx"

' Bind to Excel object.
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
  On Error GoTo 0
  Wscript.Echo "Excel application not found."
  Wscript.Quit
End If

On Error GoTo 0

' Create a new workbook.
objExcel.Workbooks.Add

Select Case UCase(strComputer)
  Case Else
    PrintServer(strComputer)
End Select

Function PrintServer(strComputer)

k=2
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer",,48)
Set colPorts =  objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")
Set strPorts = CreateObject("Scripting.Dictionary")
Set strPortNum = CreateObject("Scripting.Dictionary")

For Each objPort in colPorts
strPorts.add objPort.Name, objPort.HostAddress
strPortNum.add objPort.Name, objPort.PortNumber
Next

' Bind to worksheet.
Set objSheet = objExcel.ActiveWorkbook.Worksheets(Sheet)
objSheet.Name = strComputer

' Populate spreadsheet cells with printer attributes.
objSheet.Cells(1, 1).Value = "Name"
objSheet.Cells(1, 2).Value = "ShareName"
objSheet.Cells(1, 3).Value = "Comment"
objSheet.Cells(1, 4).Value = "Error"
objSheet.Cells(1, 5).Value = "DriverName"
objSheet.Cells(1, 6).Value = "EnableBIDI"
objSheet.Cells(1, 7).Value = "JobCount"
objSheet.Cells(1, 8).Value = "Location"
objSheet.Cells(1, 9).Value = "PortName"
objSheet.Cells(1, 10).Value = "PortAddress"
objSheet.Cells(1, 11).Value = "PortNumber"
objSheet.Cells(1, 12).Value = "Published"
objSheet.Cells(1, 13).Value = "Queued"
objSheet.Cells(1, 14).Value = "Shared"
objSheet.Cells(1, 15).Value = "Status"

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
objSheet.Cells(k, 1).Value = objItem.Name
objSheet.Cells(k, 2).Value = objItem.ShareName
objSheet.Cells(k, 3).Value = objItem.Comment
objSheet.Cells(k, 4).Value = ErrState
objSheet.Cells(k, 5).Value = objItem.DriverName
objSheet.Cells(k, 6).Value = objItem.EnableBIDI
objSheet.Cells(k, 7).Value = objItem.JobCountSinceLastReset
objSheet.Cells(k, 8).Value = objItem.Location
objSheet.Cells(k, 9).Value = objItem.PortName
objSheet.Cells(k, 10).Value = strPorts.Item(objItem.PortName)
objSheet.Cells(k, 11).Value = strPortNum.Item(objItem.PortName)
objSheet.Cells(k, 12).Value = objItem.Published
objSheet.Cells(k, 13).Value = objItem.Queued
objSheet.Cells(k, 14).Value = objItem.Shared
objSheet.Cells(k, 15).Value = objItem.Status

k = k + 1
Next

' Format the spreadsheet.
objSheet.Range("A1:O1").Font.Bold = True
objSheet.Select
objSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True
objExcel.Columns(1).ColumnWidth = 20
objExcel.Columns(2).ColumnWidth = 15
objExcel.Columns(3).ColumnWidth = 25
objExcel.Columns(4).ColumnWidth = 10
objExcel.Columns(5).ColumnWidth = 25
objExcel.Columns(6).ColumnWidth = 10
objExcel.Columns(8).ColumnWidth = 30
objExcel.Columns(9).ColumnWidth = 20
objExcel.Columns(10).ColumnWidth = 15
objExcel.Columns(11).ColumnWidth = 20
End Function

' Save the spreadsheet and close the workbook.
' objExcel.ActiveWorkbook.SaveAs strExcelPath
' objExcel.ActiveWorkbook.Close

' Quit Excel.
' objExcel.Application.Quit

' Clean Up
Set objUser = Nothing
Set objGroup = Nothing
Set objSheet = Nothing
Set objExcel = Nothing

WScript.Echo "Printer log is complete for " & Ucase(StrComputer) & " at" & VbCrLf _
& Ucase(strExcelPath)