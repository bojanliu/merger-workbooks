Attribute VB_Name = "工作表合并"
Option Explicit

Sub hebing()
    Dim ws_count As Long
    Dim sh1_lastrow As Long
    Dim last_row As Long
    Dim i As Long
    Dim j As Long
    Dim m As Long
    
    With ThisWorkbook
        ws_count = .Worksheets.Count
        
        For i = 2 To ws_count
            sh1_lastrow = .Sheets(1).[a65536].End(xlUp).Row
            last_row = .Sheets(i).[a65536].End(xlUp).Row
            Application.ScreenUpdating = False
            m = 1
            For j = sh1_lastrow + 1 To sh1_lastrow + last_row
                .Sheets(1).Cells(j, 1) = .Sheets(i).Cells(m, 1)
                .Sheets(1).Cells(j, 2) = .Sheets(i).Cells(m, 2)
                .Sheets(1).Cells(j, 3) = .Sheets(i).Cells(m, 3)
                m = m + 1
            Next j
            Application.ScreenUpdating = False
        Next i
        
        .Sheets(1).Columns("c:c").NumberFormatLocal = "yyyy-m-d"
    End With
End Sub
