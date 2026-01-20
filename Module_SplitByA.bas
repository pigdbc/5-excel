Attribute VB_Name = "Module_SplitByA"
Option Explicit

Public Sub SplitSheet1ByColumnA()
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyA As String
    Dim wsDest As Worksheet
    Dim nextRow As Long

    Set wsSrc = ThisWorkbook.Worksheets("sheet1")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For r = 2 To lastRow
        keyA = Trim(CStr(wsSrc.Cells(r, "A").Value))
        If Len(keyA) > 0 Then
            Set wsDest = GetOrCreateSheet(keyA)
            nextRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
            If nextRow < 11 Then
                nextRow = 11
            Else
                nextRow = nextRow + 1
            End If
            wsSrc.Range("A" & r & ":J" & r).Copy Destination:=wsDest.Range("A" & nextRow)
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function
