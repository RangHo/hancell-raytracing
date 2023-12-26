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

    ' Set the current zoom level to 25%
    ActiveWindow.Zoom = 25

    ' Mark the workbook as initialized
    Initialized = True
End Sub

''' Render the scene proper.
Sub Render()
    Debug_Log("Render: ################ NEW RENDERING STARTED ################")

    ' Infer aspect ratio from the given image width and height
    Dim AspectRatio
    AspectRatio = ImageWidth / ImageHeight

    ' Create a camera with the given aspect ratio
    Dim CameraOrigin, CameraFocalLength, Camera
    CameraOrigin = Vector_New(0.0, 0.0, 0.0)
    CameraFocalLength = 1.0
    Camera = Camera_New(CameraOrigin, CameraFocalLength, AspectRatio)

    ' Define the world
    Dim GroundMaterial, LeftMaterial, CenterMaterial, RightMaterial
    GroundMaterial = Lambertian_New(Vector_New(0.8, 0.8, 0.0))
    LeftMaterial = Metal_New(Vector_New(0.8, 0.8, 0.8))
    CenterMaterial = Lambertian_New(Vector_New(0.7, 0.3, 0.3))
    RightMaterial = Metal_New(Vector_New(0.8, 0.6, 0.2))

    Dim WorldObjects(4), World
    WorldObjects(1) = Sphere_New(Vector_New(0.0, -100.5, -1.0), 100.0, GroundMaterial)
    WorldObjects(2) = Sphere_New(Vector_New(-1.0, 0.0, -1.0), 0.5, LeftMaterial)
    WorldObjects(3) = Sphere_New(Vector_New(0.0, 0.0, -1.0), 0.5, CenterMaterial)
    WorldObjects(4) = Sphere_New(Vector_New(1.0, 0.0, -1.0), 0.5, RightMaterial)
    World = HittableList_New(WorldObjects)

    ' Cache A1 cell
    Dim A1Cell
    Set A1Cell = Range("A1")

    ' For each row...
    For Y = 0 To ImageHeight - 1
        ' For each column...
        For X = 0 To ImageWidth - 1
            Debug_Log("Render: Rendering (" & X & ", " & Y & ")...")

            ' Initialize the pixel color
            Dim PixelColor
            PixelColor = Vector_New(0.0, 0.0, 0.0)

            ' For each sample...
            For Sample = 0 To SamplesPerPixel - 1
                Dim U, V
                U = X
                V = (ImageHeight - 1) - Y

                ' Add a random offset to the pixel if there are more than one sample
                If SamplesPerPixel > 1 Then
                    U = U + Rnd()
                    V = V + Rnd()
                End If

                ' Normalize the pixel coordinate
                U = U / (ImageWidth - 1)
                V = V / (ImageHeight - 1)

                ' Get a ray from the camera to the pixel
                Dim R
                R = Camera_GetRay(Camera, U, V)

                ' Accumulate the color
                PixelColor = Vector_Add(PixelColor, Ray_Color(R, World))
            Next Sample

            Dim PixelColorX, PixelColorY, PixelColorZ
            PixelColorX = Sqr(Vector_X(PixelColor) / SamplesPerPixel)
            PixelColorY = Sqr(Vector_Y(PixelColor) / SamplesPerPixel)
            PixelColorZ = Sqr(Vector_Z(PixelColor) / SamplesPerPixel)
            PixelColor = Vector_New(PixelColorX, PixelColorY, PixelColorZ)
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