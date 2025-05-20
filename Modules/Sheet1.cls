VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnNextTurn_Click()
    Call NextTurn
    btnNextTurn.BackColor = RGB(200, 200, 200)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Restrict selection to board area
    If Intersect(Target, Range("B2:I9")) Is Nothing Then
        Application.EnableEvents = False
        Range("A1").Select
        Application.EnableEvents = True
        Exit Sub
    End If

    If selectingOrigin Then
        Set originCell = Target
        ' Validate that a piece is selected
        Dim rO As Integer, cO As Integer
        rO = originCell.row - 1: cO = originCell.Column - 1
        Dim pieceID As String
        pieceID = board(rO, cO)(1)
        If pieceID = "" Then
            Range("C10").Value = "No piece in that cell."
            Exit Sub
        End If
        If (currentTurn = "white" And Right(pieceID, 1) <> "b") Or _
           (currentTurn = "black" And Right(pieceID, 1) <> "n") Then
            MsgBox "It's not your turn to move that piece.", vbExclamation
            Exit Sub
        End If
        selectingOrigin = False
        Range("C10").Value = "Select destination cell"
    Else
        Set destinationCell = Target
        Call MovePiece
        selectingOrigin = True
    End If
End Sub

