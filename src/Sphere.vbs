''' Hit a sphere.
Function Sphere_Hit(Center, Radius, Ray)
    Dim OC, A, B, C, D
    OC = Vector_Subtract(Ray_Origin(Ray), Center)
    A = Vector_Dot(Ray_Direction(Ray), Ray_Direction(Ray))
    B = 2.0 * Vector_Dot(OC, Ray_Direction(Ray))
    C = Vector_Dot(OC, OC) - Radius * Radius
    D = B * B - 4.0 * A * C

    Sphere_Hit = D >= 0
End Function