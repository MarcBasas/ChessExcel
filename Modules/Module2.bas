Attribute VB_Name = "Module2"
Public currentTurn As String
Public selectingOrigin As Boolean
Public originCell As Range
Public destinationCell As Range
Public board() As Variant

Sub MovePiece()
    ' Check if origin cell has a piece
    If originCell.Value = "" Then
        MsgBox "No piece in origin cell.", vbExclamation
        Exit Sub
    End If

    ' Ensure destination is different
    If originCell.Address = destinationCell.Address Then
        Range("C10").Value = "Select a different cell"
        Exit Sub
    End If

    ' Calculate board indices
    Dim rO As Integer, cO As Integer
    Dim rD As Integer, cD As Integer
    rO = originCell.row - 1
    cO = originCell.Column - 1
    rD = destinationCell.row - 1
    cD = destinationCell.Column - 1

    ' Validate turn ownership
    Dim pieceID As String
    pieceID = board(rO, cO)(1)
    If (currentTurn = "white" And Right(pieceID, 1) <> "b") Or _
       (currentTurn = "black" And Right(pieceID, 1) <> "n") Then
        MsgBox "It's not your turn to move that piece.", vbExclamation
        Exit Sub
    End If

    ' Check destination occupancy
    Dim destID As String
    destID = board(rD, cD)(1)
    If destID <> "" Then
        Dim colorO As String, colorD As String
        colorO = Right(pieceID, 1)
        colorD = Right(destID, 1)
        If colorO = colorD Then
            MsgBox "You cannot move onto a square occupied by your own piece.", vbExclamation
            Exit Sub
        End If
    End If

    ' Validate piece-specific moves
    Select Case pieceID
        Case "pb", "pn"
            If Not IsValidPawnMove(pieceID, rO, cO, rD, cD, destID) Then
                MsgBox "Invalid pawn move.", vbExclamation
                Exit Sub
            End If
        Case "rb", "rn"
            If Not IsValidRookMove(rO, cO, rD, cD) Then
                MsgBox "Invalid rook move.", vbExclamation
                Exit Sub
            End If
        Case "nb", "nn"
            If Not IsValidKnightMove(rO, cO, rD, cD) Then
                MsgBox "Invalid knight move.", vbExclamation
                Exit Sub
            End If
        Case "bb", "bn"
            If Not IsValidBishopMove(rO, cO, rD, cD) Then
                MsgBox "Invalid bishop move.", vbExclamation
                Exit Sub
            End If
        Case "qb", "qn"
            If Not IsValidQueenMove(rO, cO, rD, cD) Then
                MsgBox "Invalid queen move.", vbExclamation
                Exit Sub
            End If
        Case "kb", "kn"
            If Not IsValidKingMove(rO, cO, rD, cD) Then
                MsgBox "Invalid king move.", vbExclamation
                Exit Sub
            End If
    End Select

    ' Execute move visually
    destinationCell.Value = originCell.Value
    originCell.Value = ""

    ' Check for capture of king
    If destID = "kb" Or destID = "kn" Then
        MsgBox "Checkmate! You have captured the enemy king.", vbExclamation
    End If

    ' Update board array
    board(rD, cD) = board(rO, cO)
    board(rO, cO) = Array("", "")

    ' Prompt next turn
    Range("C10").Value = "Click Next Turn."
    Worksheets("ChessExcel").OLEObjects("btnNextTurn").Object.BackColor = RGB(240, 210, 15)
End Sub

Sub NextTurn()
    If currentTurn = "white" Then
        currentTurn = "black"
        Range("B10").Value = "Turn: Black"
    Else
        currentTurn = "white"
        Range("B10").Value = "Turn: White"
    End If
    Range("C10").Value = "Select a piece"
    Worksheets("ChessExcel").OLEObjects("btnNextTurn").Object.BackColor = RGB(200, 200, 200)
    selectingOrigin = True
End Sub

' Module: Move Validation Functions
Function IsValidPawnMove(pieceID As String, rO As Integer, cO As Integer, rD As Integer, cD As Integer, destID As String) As Boolean
    IsValidPawnMove = False
    Dim direction As Integer, startRow As Integer
    If pieceID = "pb" Then
        direction = -1
        startRow = 7
    Else
        direction = 1
        startRow = 2
    End If
    ' Single forward
    If cO = cD And destID = "" Then
        If rD = rO + direction Then
            IsValidPawnMove = True: Exit Function
        ElseIf rO = startRow And rD = rO + 2 * direction Then
            ' Double from start
            If board(rO + direction, cO)(1) = "" Then
                IsValidPawnMove = True: Exit Function
            End If
        End If
    End If
    ' Diagonal capture
    If Abs(cD - cO) = 1 And rD = rO + direction Then
        If destID <> "" And Right(destID, 1) <> Right(pieceID, 1) Then
            IsValidPawnMove = True: Exit Function
        End If
    End If
End Function

Function IsValidRookMove(rO As Integer, cO As Integer, rD As Integer, cD As Integer) As Boolean
    IsValidRookMove = False
    If rO <> rD And cO <> cD Then Exit Function
    Dim dr As Integer, dc As Integer
    dr = Sgn(rD - rO): dc = Sgn(cD - cO)
    Dim rr As Integer: rr = rO + dr
    Dim cc As Integer: cc = cO + dc
    Do While rr <> rD Or cc <> cD
        If board(rr, cc)(1) <> "" Then Exit Function
        rr = rr + dr: cc = cc + dc
    Loop
    IsValidRookMove = True
End Function

Function IsValidKnightMove(rO As Integer, cO As Integer, rD As Integer, cD As Integer) As Boolean
    Dim dr As Integer, dc As Integer
    dr = Abs(rD - rO): dc = Abs(cD - cO)
    IsValidKnightMove = (dr = 2 And dc = 1) Or (dr = 1 And dc = 2)
End Function

Function IsValidBishopMove(rO As Integer, cO As Integer, rD As Integer, cD As Integer) As Boolean
    IsValidBishopMove = False
    If Abs(rD - rO) <> Abs(cD - cO) Then Exit Function
    Dim dr As Integer, dc As Integer
    dr = Sgn(rD - rO): dc = Sgn(cD - cO)
    Dim rr As Integer: rr = rO + dr
    Dim cc As Integer: cc = cO + dc
    Do While rr <> rD And cc <> cD
        If board(rr, cc)(1) <> "" Then Exit Function
        rr = rr + dr: cc = cc + dc
    Loop
    IsValidBishopMove = True
End Function

Function IsValidQueenMove(rO As Integer, cO As Integer, rD As Integer, cD As Integer) As Boolean
    IsValidQueenMove = IsValidRookMove(rO, cO, rD, cD) Or IsValidBishopMove(rO, cO, rD, cD)
End Function

Function IsValidKingMove(rO As Integer, cO As Integer, rD As Integer, cD As Integer) As Boolean
    Dim dr As Integer, dc As Integer
    dr = Abs(rD - rO): dc = Abs(cD - cO)
    IsValidKingMove = (dr <= 1 And dc <= 1)
End Function


