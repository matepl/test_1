Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' This event fires when any cell in Sheet1 changes
    ' Update the count in Sheet2 cell B1
    UpdateMaterialnummerCount
End Sub

Private Sub Worksheet_Calculate()
    ' This event fires when the worksheet recalculates
    ' Update the count in Sheet2 cell B1
    UpdateMaterialnummerCount
End Sub