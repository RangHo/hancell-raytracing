''' Hit a sphere.
Function Sphere_Hit(Center, Radius, Ray)
    Dim OC, A, BHalf, C, D
    OC = Vector_Subtract(Ray_Origin(Ray), Center)
    A = Vector_Dot(Ray_Direction(Ray), Ray_Direction(Ray))
    BHalf = Vector_Dot(OC, Ray_Direction(Ray))
    C = Vector_Dot(OC, OC) - (Radius * Radius)
    D = (BHalf * BHalf) - (A * C)

    If D > 0.0 Then
        Sphere_Hit = (-BHalf - Sqr(D)) / A
    Else
        Sphere_Hit = -1.0
    End If
End Function