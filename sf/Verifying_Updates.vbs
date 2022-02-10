''''''''''''''''''''''''''''''''''''''''''''''
'
'  Verify Windows Updates for a computer
'
''''''''''''''''''''''''''''''''''''''''''''''
'On Error Resume Next

strComputer = Ucase(InputBox ("Please enter the computer name to check.", "Computer Name:","MWxxxSSI"))
strDate = Replace(Replace(Date(),"/","-"),":",".")
strTime = Replace(Replace(Time(),"/","-"),":",".")
 

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

objExcel.Workbooks.Add

intRow = 3


objExcel.Cells(1, 1).Value = strdate & " @ " & strtime 

objExcel.Cells(2, 1).Value = "Machine Name"

objExcel.Cells(2, 2).Value = "Update"

objExcel.Cells(2, 3).Value = "Status"

objExcel.Cells(2, 4).Value = "Date"

objExcel.Cells(2, 5).Value = "Source"

 

Set objSession = CreateObject("Microsoft.Update.Session", strComputer)

Set objSearcher = objSession.CreateUpdateSearcher

intHistoryCount = objSearcher.GetTotalHistoryCount

Set colHistory = objSearcher.QueryHistory(1, intHistoryCount)

 

For Each objEntry in colHistory

objExcel.Cells(intRow, 1).Value = UCase(strComputer)

 

objExcel.Cells(intRow, 2).Value = objEntry.Title

Select Case objEntry.ResultCode

Case 0 ResultCode = "Not Started"

Case 1 ResultCode = "In Progress"

Case 2 ResultCode = "Success"

Case 3 ResultCode = "Error"

Case 4 ResultCode = "Failed"

Case 5 ResultCode = "Cancelled"

End Select


objExcel.Cells(intRow, 3).Value = ResultCode

objExcel.Cells(intRow, 4).Value = objEntry.Date

objExcel.Cells(intRow, 5).Value = objEntry.ClientApplicationID

intRow = intRow + 1

Next

 

objExcel.Range("A2:E2").Select

objExcel.Selection.Interior.ColorIndex = 19

objExcel.Selection.Font.ColorIndex = 11

objExcel.Selection.Font.Bold = True

objExcel.Cells.EntireColumn.AutoFit

 

Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)

Set objRange = objExcel.Range("D1")

objRange.Sort objRange,2,,,,,,1

 

MsgBox "Done"
