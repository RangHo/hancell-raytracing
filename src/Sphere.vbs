''' Create a new sphere.
Function Sphere_New(Center, Radius, Material)
    Dim Result(3)
    Result(0) = "Sphere"
    Result(1) = Center
    Result(2) = Radius
    Result(3) = Material

    Sphere_New = Result
End Function

''' Get the center of the sphere.
Function Sphere_Center(Sphere)
    Sphere_Center = Sphere(1)
End Function

''' Get the radius of the sphere.
Function Sphere_Radius(Sphere)
    Sphere_Radius = Sphere(2)
End Function

''' Get the material of the sphere.
Function Sphere_Material(Sphere)
    Sphere_Material = Sphere(3)
End Function

''' Hit a sphere.
Function Sphere_Hit(Sphere, Ray, TMin, TMax)
    Dim Center, Radius, OC, A, BHalf, C, D
    Center = Sphere_Center(Sphere)
    Radius = Sphere_Radius(Sphere)
    OC = Vector_Subtract(Ray_Origin(Ray), Center)
    A = Vector_Dot(Ray_Direction(Ray), Ray_Direction(Ray))
    BHalf = Vector_Dot(OC, Ray_Direction(Ray))
    C = Vector_Dot(OC, OC) - (Radius * Radius)
    D = (BHalf * BHalf) - (A * C)

    If D > 0.0 Then
        Debug_Log("Sphere_Hit: Intersecting ray with discriminant " & D)
        Dim SqrtD, Root, HitPoint, HitNormal, HitMaterial
        SqrtD = Sqr(D)
        Root = (-BHalf - SqrtD) / A

        If Root < TMin Or TMax < Root Then
            Root = (-BHalf + SqrtD) / A

            If Root < TMin Or TMax < Root Then
                Sphere_Hit = HitResult_Miss()
                Exit Function
            End If
        End If

        HitPoint = Ray_At(Ray, Root)
        HitNormal = Vector_Scale(Vector_Subtract(HitPoint, Center), 1.0 / Radius)
        HitMaterial = Sphere_Material(Sphere)

        Sphere_Hit = HitResult_Hit(HitPoint, HitNormal, Root, HitMaterial)
    Else
        Debug_Log("Sphere_Hit: Missing ray with discriminant " & D)
        Sphere_Hit = HitResult_Miss()
    End If
End Function
