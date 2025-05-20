VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    With Worksheets("ChessExcel")
        .Unprotect Password:="ChessExcel"
        .EnableSelection = xlNoRestrictions
        .Cells.Locked = False
        .Range("A1:Z100").ClearFormats
        .Range("B2:I9").ClearContents
        ActiveWindow.DisplayGridlines = True
    End With
End Sub

