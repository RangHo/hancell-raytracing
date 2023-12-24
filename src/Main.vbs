Dim Initialized

''' Initialize the workbook for rendering.
Sub Init()
    ' If the workbook is already initialized, do nothing
    If Initialized Then
        Exit Sub
    End If

    ' There should be two sheets: one for rendering and one for log output
    Sheets.Add
    Sheets(1).Name = "Output"
    Sheets(2).Name = "Log"

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

    ' Mark the workbook as initialized
    Initialized = True
End Sub

''' Render the scene proper.
Sub Render()
    Debug_Log("Render: ################ NEW RENDERING STARTED ################")

    ' Infer aspect ratio from the given image width and height
    Dim AspectRatio
    AspectRatio = ImageWidth / ImageHeight

    ' From the aspect ratio, calculate the viewport ratio
    ' (Note that the height is always 2.0)
    Dim ViewportWidth, ViewportHeight
    ViewportHeight = 2.0
    ViewportWidth = ViewportHeight * AspectRatio

    ' Viewport helpers
    Dim ViewportU, ViewportV
    ViewportU = Vector_New(ViewportWidth, 0.0, 0.0)
    ViewportV = Vector_New(0.0, -ViewportHeight, 0.0)
    Debug_Log("Render: ViewportU is " & Vector_ToString(ViewportU))
    Debug_Log("Render: ViewportV is " & Vector_ToString(ViewportV))

    ' Per-pixel delta helpers
    Dim PixelDeltaU, PixelDeltaV
    PixelDeltaU = Vector_Scale(ViewportU, 1.0 / ImageWidth)
    PixelDeltaV = Vector_Scale(ViewportV, 1.0 / ImageHeight)
    Debug_Log("Render: PixelDeltaU is " & Vector_ToString(PixelDeltaU))
    Debug_Log("Render: PixelDeltaV is " & Vector_ToString(PixelDeltaV))

    ' Rendering helpers
    ' ViewportOrigin = CameraOrigin - Vector(0, 0, FocalLength) - (ViewportU / 2) - (ViewportV / 2)
    ' PixelOrigin = ViewportOrigin + (PixelDeltaU + PixelDeltaV) / 2
    Dim CameraOrigin, ViewportOrigin, PixelOrigin
    CameraOrigin = Vector_New(0.0, 0.0, 0.0)
    ViewportOrigin = Vector_Subtract(Vector_Subtract(Vector_Subtract(CameraOrigin, Vector_New(0.0, 0.0, FocalLength)), Vector_Scale(ViewportU, 0.5)), Vector_Scale(ViewportV, 0.5))
    PixelOrigin = Vector_Add(ViewportOrigin, Vector_Scale(Vector_Add(PixelDeltaU, PixelDeltaV), 0.5))
    Debug_Log("Render: CameraOrigin is " & Vector_ToString(CameraOrigin))
    Debug_Log("Render: ViewportOrigin is " & Vector_ToString(ViewportOrigin))
    Debug_Log("Render: PixelOrigin is " & Vector_ToString(PixelOrigin))

    ' Define the world
    Dim WorldObjects(2), World
    WorldObjects(1) = Sphere_New(Vector_New(0.0, 0.0, -1.0), 0.5)
    WorldObjects(2) = Sphere_New(Vector_New(0.0, -100.5, -1.0), 100.0)
    World = HittableList_New(WorldObjects)

    ' Cache A1 cell
    Dim A1Cell
    Set A1Cell = Range("A1")

    ' For each row...
    For Y = 0 To ImageHeight - 1
        ' For each column...
        For X = 0 To ImageWidth - 1
            Debug_Log("Render: Rendering (" & X & ", " & Y & ")...")

            ' U and V are ratio from the start of the coordinate
            ' Vector pointing at each pixel on the viewport
            ' Destination = PixelOrigin + PixelDeltaU * X + PixelDeltaV * V
            Dim Destination
            Destination = Vector_Add(PixelOrigin, Vector_Add(Vector_Scale(PixelDeltaU, X), Vector_Scale(PixelDeltaV, Y)))

            ' Construct a ray from camera origin to destination
            Dim R
            R = Ray_New(CameraOrigin, Destination)

            ' Create a color vector
            Dim PixelColor
            PixelColor = Ray_Color(R, World)
            Debug_Log("Render: Color of (" & X & ", " & Y & "): " & Vector_ToString(PixelColor))

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