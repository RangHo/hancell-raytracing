''' Get the origin component of the ray.
Function Ray_Origin(R)
    Ray_Origin = R(1)
End Function

''' Get the direction component of the ray.
Function Ray_Direction(R)
    Ray_Direction = R(2)
End Function

''' Create a new ray.
Function Ray_New(Origin, Direction)
    Dim Result(2)
    Result(1) = Origin
    Result(2) = Direction

    Ray_New = Result
End Function

''' Get the location vector to which the ray points at when advanced by T.
Function Ray_At(R, T)
    Ray_At = Vector_Add(Ray_Origin(R), Vector_Scale(Ray_Direction(R), T))
End Function

''' Get the color of the ray
Function Ray_Color(R)
    If Sphere_Hit(Vector_New(0.0, 0.0, -1.0), 0.5, R) Then
        Ray_Color = Vector_New(1.0, 0.0, 0.0)
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
End Function