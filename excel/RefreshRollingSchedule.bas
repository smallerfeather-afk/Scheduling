Attribute VB_Name = "RefreshRollingSchedule"
Option Explicit

' Optional macro:
' Expands 14 rolling dates into dated staffing rows by joining RollingCalendar dates
' to MasterTemplate rows using TemplateDay (1..14).
Public Sub RefreshRollingScheduleOutput()
    Dim wsDates As Worksheet
    Dim wsMaster As Worksheet
    Dim wsOut As Worksheet

    Dim lastMasterRow As Long
    Dim outRow As Long
    Dim i As Long, j As Long

    Dim rollingDate As Date
    Dim templateDay As Long

    Set wsDates = ThisWorkbook.Worksheets("RollingCalendar")
    Set wsMaster = ThisWorkbook.Worksheets("MasterTemplate")
    Set wsOut = GetOrCreateSheet("RollingOutput")

    wsOut.Cells.Clear
    wsOut.Range("A1:H1").Value = Array("Date", "TemplateDay", "Shift", "Role", "Employee", "StartTime", "EndTime", "Notes")

    lastMasterRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    outRow = 2

    ' RollingCalendar layout assumption:
    ' A4:A17 = Date
    ' B4:B17 = TemplateDay
    For i = 4 To 17
        If IsDate(wsDates.Cells(i, "A").Value) Then
            rollingDate = wsDates.Cells(i, "A").Value
            templateDay = CLng(wsDates.Cells(i, "B").Value)

            ' For each rolling date, append every matching master row
            For j = 2 To lastMasterRow
                If CLng(wsMaster.Cells(j, "A").Value) = templateDay Then
                    wsOut.Cells(outRow, "A").Value = rollingDate
                    wsOut.Cells(outRow, "B").Value = templateDay
                    wsOut.Cells(outRow, "C").Value = wsMaster.Cells(j, "B").Value ' Shift
                    wsOut.Cells(outRow, "D").Value = wsMaster.Cells(j, "C").Value ' Role
                    wsOut.Cells(outRow, "E").Value = wsMaster.Cells(j, "D").Value ' Employee
                    wsOut.Cells(outRow, "F").Value = wsMaster.Cells(j, "E").Value ' StartTime
                    wsOut.Cells(outRow, "G").Value = wsMaster.Cells(j, "F").Value ' EndTime
                    wsOut.Cells(outRow, "H").Value = wsMaster.Cells(j, "G").Value ' Notes
                    outRow = outRow + 1
                End If
            Next j
        End If
    Next i

    wsOut.Columns("A:H").AutoFit
    MsgBox "RollingOutput refreshed successfully.", vbInformation
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function
