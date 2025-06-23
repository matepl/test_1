Option Explicit

' NOTE: For automatic updating to work, you need to place these event handlers
' in the Sheet1 (first worksheet) code module, not in a separate module:
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     UpdateMaterialnummerCount
' End Sub
'
' Private Sub Worksheet_Calculate()
'     UpdateMaterialnummerCount
' End Sub

Sub UpdateMaterialnummerCount()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long
    Dim cellValue As Variant
    
    ' Set references to the worksheets
    Set ws1 = ThisWorkbook.Sheets(1)  ' First sheet (data sheet)
    Set ws2 = ThisWorkbook.Sheets(2)  ' Second sheet (results sheet)
    
    ' Find last row in first column of first sheet
    lastRow = ws1.Cells(ws1.Rows.count, 1).End(xlUp).Row
    
    count = 0
    
    ' Count rows with material numbers
    For i = 1 To lastRow
        cellValue = ws1.Cells(i, 1).Value
        
        If Not IsEmpty(cellValue) And IsNumeric(cellValue) Then
            count = count + 1
        End If
    Next i
    
    ' Update cell B1 in the second sheet
    ws2.Cells(1, 2).Value = count
End Sub

Sub CountMaterialnummer()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long
    Dim cellValue As Variant
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    count = 0
    
    For i = 1 To lastRow
        cellValue = ws.Cells(i, 1).Value
        
        If Not IsEmpty(cellValue) And IsNumeric(cellValue) Then
            count = count + 1
        End If
    Next i
    
    MsgBox "Number of rows with Materialnummer: " & count, vbInformation, "Count Result"
End Sub

Function CountMaterialnummerRows() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim count As Long
    Dim cellValue As Variant
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    count = 0
    
    For i = 1 To lastRow
        cellValue = ws.Cells(i, 1).Value
        
        If Not IsEmpty(cellValue) And IsNumeric(cellValue) Then
            count = count + 1
        End If
    Next i
    
    CountMaterialnummerRows = count
End Function