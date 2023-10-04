''' Aspect ratio of the rendered image.
Dim AspectRatio

''' Width and height of the viewport.
Dim ViewportWidth, ViewportHeight

''' Initialize the workbook for rendering.
Sub Init()
    ' Infer aspect ratio from the given image width and height
    AspectRatio = ImageWidth / ImageHeight

    ' From the aspect ratio, calculate the viewport ratio
    ' (Note that the height is always 2.0)
    ViewportHeight = 2.0
    ViewportWidth = ViewportHeight * AspectRatio

    ' There should be two sheets: one for rendering and one for log output
    Sheets.Add
    Sheets(1).Name = "Output"
    Sheets(2).Name = "Log"

    ' Select the log output sheet
    Sheets("Log").Select

    ' Merge all cells
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .MergeCells = True
    End With

    ' Select the output sheet
    Sheets("Output").Select

    ' Set the column and row size to 14px x 14px
    ' This size is less than ideal, but somehow Hancell shits itself when column
    ' width is less than 1 and mysteriously adds 0.13 to the width so... yeah
    Cells.Select
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
    ' Cache A1 cell
    Dim A1Cell
    Set A1Cell = Range("A1")

    ' World coordinate helpers
    Dim Origin, Horizontal, Vertical
    Origin = Vector_New(0, 0, 0)
    Horizontal = Vector_New(ViewportWidth, 0, 0)
    Vertical = Vector_New(0, -ViewportHeight, 0)

    ' Calculate the vector to the lower left corner of viewport
    ' LowerLeftCorner = Origin - (Horizontal / 2) - (Vertical / 2) - Vector(0, 0, FocalLength)
    LowerLeftCorner = Vector_Subtract(Vector_Subtract(Vector_Subtract(Origin, Vector_Scale(Horizontal, 0.5)), Vector_Scale(Vertical, 0.5)), Vector_New(0, 0, FocalLength))

    ' For each row...
    For Y = 0 To ImageHeight - 1
        ' For each column...
        For X = 0 To ImageWidth - 1
            Debug_Log("Processing (" & X & ", " & Y & ")")

            ' U and V are ratio from the start of the coordinate
            Dim U, V
            U = X / (ImageWidth - 1)
            V = Y / (ImageWidth - 1)

            ' Vector pointing at each pixel on the viewport
            ' Destination = LowerLeftCorner + (Horizontal * U) + (Vertical * V) - Origin
            Dim Destination
            Destination = Vector_Subtract(Vector_Add(Vector_Add(LowerLeftCorner, Vector_Scale(Horizontal, U)), Vector_Scale(Vertical, V)), Origin)

            ' Construct a ray from camera origin to destination
            Dim R
            R = Ray_New(Origin, Destination)

            ' Create a color vector
            Dim PixelColor
            PixelColor = Ray_Color(R)

            ' Color the cell
            With A1Cell.Offset(Y, X).Interior
                .Color = Color_Vector2RGB(PixelColor)
                .Pattern = xlPatternSolid
            End With
        Next X
    Next Y
End Sub

''' Time how long it takes to render a scene.
Sub Benchmark()
    ' Declare variables and save starting time in milliseconds
    Dim StartTime, EndTime
    StartTime = Int(Timer * 1000)

    ' Start rendering
    Call Render

    ' Save ending time
    EndTime = Int(Timer * 1000)

    ' Calculate some variables
    Dim PixelCount, Duration, PerPixelDuration
    PixelCount = ImageWidth * ImageHeight
    Duration = EndTime - StartTime
    PerPixelDuration = Duration / PixelCount

    ' Show how much it took
    MsgBox("It took " & Duration & " milliseconds to paint " & PixelCount & " pixels. (" & PerPixelDuration & "ms/px)")
End Sub