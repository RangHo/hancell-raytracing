''' Initialize the spreadsheet for rendering.
Const ImageWidth = 20

Const ImageHeight = 20

Sub Init()
    ' Select all cells
    Cells.Select

    ' Set the column and row size to 14px x 14px
    ' This size is less than ideal, but somehow Hancell shits itself when column
    ' width is less than 1 and mysteriously adds 0.13 to the width so... yeah
    Selection.ColumnWidth = 1.00 ' (+ mysterious 0.13) = 14px
    Selection.RowHeight = 10.50 ' = 14px

    ' Remove the annoying grid pattern
    With Selection.Interior
        .ColorIndex = 0
        .Pattern = xlPatternSolid
    End With

    ' Move the selection out of the way
    Range("A1").Offset(ImageHeight, ImageWidth).Select
End Sub

''' Render the scene proper.
Sub Render()
    ' For each row...
    For Y = 0 To ImageHeight - 1
        For X = 0 To ImageWidth - 1
            ' Find the colors in [0, 1] range
            Dim R, G, B
            R = X / (ImageWidth - 1)
            G = Y / (ImageHeight - 1)
            B = 0.25

            ' Convert it into [0, 255] integer range
            Dim IR, IG, IB
            IR = Int(255.999 * R)
            IG = Int(255.999 * G)
            IB = Int(255.999 * B)

            ' Color the cell
            With Range("A1").Offset(Y, X).Interior
                .Color = RGB(IR, IG, IB)
                .Pattern = xlPatternSolid
            End With
        Next X
    Next Y
End Sub