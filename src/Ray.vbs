''' Maximum recursion depth for Ray_Color_Recursive
Const Ray_MaxDepth = 64

''' Create a new ray.
Function Ray_New(Origin, Direction)
    Dim Result(2)
    Result(1) = Origin
    Result(2) = Direction

    Ray_New = Result
End Function

''' Get the origin component of the ray.
Function Ray_Origin(R)
    Ray_Origin = R(1)
End Function

''' Get the direction component of the ray.
Function Ray_Direction(R)
    Ray_Direction = R(2)
End Function

''' Get the location vector to which the ray points at when advanced by T.
Function Ray_At(R, T)
    Ray_At = Vector_Add(Ray_Origin(R), Vector_Scale(Ray_Direction(R), T))
End Function

''' Get the color of the ray
Function Ray_Color(R, World, Optional Depth = 0)
    Debug_Log("Ray_Color: Enter")

    Dim HitResult
    HitResult = Hit(World, R, 0.01, 1000.0)

    If HitResult_IsHit(HitResult) Then
        Dim Target, NewRay
        Target = Vector_Add(Vector_Add(HitResult_Point(HitResult), HitResult_Normal(HitResult)), Math_RandomVector())
        NewRay = Ray_New(HitResult_Point(HitResult), Vector_Subtract(Target, HitResult_Point(HitResult)))
        Ray_Color = Vector_Scale(Ray_Color_Recursive(NewRay, World, Depth + 1), 0.5)
    Else
        ' Find the normalized version of the direction vector
        Dim NormalizedDirection
        NormalizedDirection = Vector_Normalize(Ray_Direction(R))

        ' Based on the normalized Y position, perform linear interpolation
        Dim T
        T = 0.5 * (Vector_Y(NormalizedDirection) + 1.0)

        ' Ray_Color = Vector(1, 1, 1) * (1.0 - T) + Vector(0.5, 0.7, 1.0) * T
        Ray_Color = Vector_Add(Vector_Scale(Vector_New(1, 1, 1), 1.0 - T), Vector_Scale(Vector_New(0.5, 0.7, 1.0), T))
    End If

    Debug_Log("Ray_Color: Exit")
End Function

''' Stub function to recursively call Ray_Color.
Function Ray_Color_Recursive(R, World, Depth)
    If Depth >= Ray_MaxDepth Then
        Debug_Log("Ray_Color_Recursive: Maximum recursion depth reached")
        Ray_Color_Recursive = Vector_New(0, 0, 0)
        Exit Function
    End If

    Ray_Color_Recursive = Ray_Color(R, World, Depth)
End Function