Attribute VB_Name = "Module1"
Sub DrawBoard()
    Dim row As Integer, col As Integer

    ' Unprotect sheet before modifying it
    Worksheets("ChessExcel").Unprotect Password:="ChessExcel"

    ' Clear board cells
    Range("B2:I9").ClearContents

    ' Clear sheet formatting and content outside board
    ActiveWindow.DisplayGridlines = False
    Dim rng As Range
    Set rng = Union(Range("A1:A100"), Range("J1:Z100"), Range("A1:Z1"), Range("A10:Z100"))
    rng.ClearFormats
    rng.ClearContents

    ' Draw background for board
    For row = 2 To 9
        For col = 2 To 9
            With Cells(row, col)
                .Font.Size = 70
                .Font.Name = "Segoe UI Symbol"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Locked = True
                If (row + col) Mod 2 = 0 Then
                    .Interior.color = RGB(255, 255, 255)
                Else
                    .Interior.color = RGB(100, 100, 100)
                End If
            End With
        Next col
    Next row

    ' Add thick borders around board
    With Range("B2:I9").Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
        .color = RGB(0, 0, 0)
    End With

    ' Initialize board array
    ReDim board(1 To 8, 1 To 8)

    ' Unicode piece codes for setup
    Dim blackPieces As Variant
    Dim whitePieces As Variant
    blackPieces = Array(9820, 9822, 9821, 9819, 9818, 9821, 9822, 9820)
    whitePieces = Array(9814, 9816, 9815, 9813, 9812, 9815, 9816, 9814)

    ' Place black pieces on row 1 and pawns on row 2
    For col = 0 To 7
        With Cells(2, col + 2)
            .Value = ChrW(blackPieces(col))
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = 70
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        board(1, col + 1) = Array(ChrW(blackPieces(col)), PieceName("n", col))

        With Cells(3, col + 2)
            .Value = ChrW(9823) ' Black pawn
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = 70
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        board(2, col + 1) = Array(ChrW(9823), "pn")

        ' Place white pawns on row 7
        With Cells(8, col + 2)
            .Value = ChrW(9817) ' White pawn
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = 70
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        board(7, col + 1) = Array(ChrW(9817), "pb")

        ' Place white back row on row 8
        With Cells(9, col + 2)
            .Value = ChrW(whitePieces(col))
            .Font.Name = "Segoe UI Symbol"
            .Font.Size = 70
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        board(8, col + 1) = Array(ChrW(whitePieces(col)), PieceName("b", col))
    Next col

    ' Clear the rest of board array
    For row = 4 To 7
        For col = 1 To 8
            board(row - 1, col) = Array("", "")
        Next col
    Next row

    ' Unlock message cells before protecting sheet
    With Worksheets("ChessExcel").Range("B10:C10")
        .Locked = False
    End With

    ' Protect sheet for interface only
    Worksheets("ChessExcel").Protect Password:="ChessExcel", UserInterfaceOnly:=True
    Worksheets("ChessExcel").EnableSelection = xlNoRestrictions

    ' Move focus away from board
    Range("A1").Select

    currentTurn = "white"
    selectingOrigin = True
    Range("B10").Value = "Turn: White"
    Range("C10").Value = "Select a piece"
End Sub

' Function to return piece identifier based on color and index
Function PieceName(color As String, index As Integer) As String
    Dim whiteNames As Variant
    whiteNames = Array("rb", "nb", "bb", "qb", "kb", "bb", "nb", "rb")
    Dim blackNames As Variant
    blackNames = Array("rn", "nn", "bn", "qn", "kn", "bn", "nn", "rn")

    If color = "b" Then
        PieceName = whiteNames(index)
    Else
        PieceName = blackNames(index)
    End If
End Function
