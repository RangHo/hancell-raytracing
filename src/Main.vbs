''' Initialize the spreadsheet for rendering.
Sub Init()
    ' Select all cells
    Cells.Select

    ' Set the column and row size to 14px x 14px
    ' This size is less than ideal, but somehow Hancell shits itself when column
    ' width is less than 1 and mysteriously adds 0.13 to the width so... yeah
    Selection.ColumnWidth = 1.00 ' (+ mysterious 0.13) = 14px
    Selection.RowHeight = 10.50 ' = 14px

    ' Select the top left cell
    Range("A1")
End Sub

''' Render the scene proper.
Sub Render()
    MsgBox "Hello, world!"
End Sub
