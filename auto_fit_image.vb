Public Sub FitPic()
    On Error GoTo NOT_SHAPE
    Dim PicWtoHRatio As Single
    Dim CellWtoHRatio As Single
    With Selection
        PicWtoHRatio = .Width / .Height
    End With
    With Selection.TopLeftCell
        CellWtoHRatio = .Width / .RowHeight
    End With
    Select Case PicWtoHRatio / CellWtoHRatio
        Case Is > 1
            With Selection
                .Width = .TopLeftCell.Width
                .Height = .Width / PicWtoHRatio
            End With
        Case Else
            With Selection
                .Height = .TopLeftCell.RowHeight
                .Width = .Height * PicWtoHRatio
            End With
    End Select
    With Selection
        .Top = .TopLeftCell.Top
        .Left = .TopLeftCell.Left
    End With
    Exit Sub
    NOT_SHAPE:
    MsgBox "Select a picture before running this macro."
End Sub