''' Create new camera.
Function Camera_New(Origin, FocalLength, AspectRatio)
    Dim Result(6)
    Result(1) = Origin
    Result(2) = FocalLength
    Result(3) = AspectRatio

    ' These are supporting cache variables
    Dim ViewportHeight, ViewportLength, ViewportHorizontal, ViewportVertical, ToLowerLeftCorner
    ViewportHeight = 2.0
    ViewportLength = AspectRatio * ViewportHeight
    ViewportHorizontal = Vector_New(ViewportLength, 0.0, 0.0)
    ViewportVertical = Vector_New(0.0, ViewportHeight, 0.0)
    ToLowerLeftCorner = Vector_Subtract(Vector_Subtract(Vector_Subtract(Origin, Vector_Scale(ViewportHorizontal, 0.5)), Vector_Scale(ViewportVertical, 0.5)), Vector_New(0.0, 0.0, FocalLength))
    Result(4) = ViewportHorizontal
    Result(5) = ViewportVertical
    Result(6) = ToLowerLeftCorner

    Camera_New = Result
End Function

''' Get the origin of the camera.
Function Camera_Origin(Camera)
    Camera_Origin = Camera(1)
End Function

''' Get the focal length of the camera.
Function Camera_FocalLength(Camera)
    Camera_FocalLength = Camera(2)
End Function

''' Get the aspect ratio of the camera.
Function Camera_AspectRatio(Camera)
    Camera_AspectRatio = Camera(3)
End Function

''' Get the viewport horizontal vector of the camera.
Function Camera_ViewportHorizontal(Camera)
    Camera_ViewportHorizontal = Camera(4)
End Function

''' Get the viewport vertical vector of the camera.
Function Camera_ViewportVertical(Camera)
    Camera_ViewportVertical = Camera(5)
End Function

''' Get the vector to the lower left corner of the viewport.
Function Camera_ToLowerLeftCorner(Camera)
    Camera_ToLowerLeftCorner = Camera(6)
End Function

''' Get a ray from the camera to a given position in the viewport.
Function Camera_GetRay(Camera, U, V)
    Dim Origin, ToLowerLeftCorner, ViewportHorizontal, ViewportVertical, Direction
    Origin = Camera_Origin(Camera)
    ToLowerLeftCorner = Camera_ToLowerLeftCorner(Camera)
    ViewportHorizontal = Camera_ViewportHorizontal(Camera)
    ViewportVertical = Camera_ViewportVertical(Camera)
    Direction = Vector_Add(Vector_Add(ToLowerLeftCorner, Vector_Scale(ViewportHorizontal, U)), Vector_Scale(ViewportVertical, V))

    Camera_GetRay = Ray_New(Origin, Direction)
End Function

''' Get the color of the camera.