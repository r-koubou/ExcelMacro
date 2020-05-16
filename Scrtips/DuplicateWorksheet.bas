' ==============================================================================
'
' DuplicateWorksheet.bas
' (c) R-Koubou
'
' ==============================================================================

' Duplicate a current worksheet macro
Sub DuplicateWorksheet()

    Dim sheetName As String
    Dim sheetIndex As Long

    sheetName = ActiveSheet.Name
    sheetIndex = ActiveSheet.Index
    Worksheets(sheetName).Copy After:=Worksheets(sheetIndex)
    
End Sub
