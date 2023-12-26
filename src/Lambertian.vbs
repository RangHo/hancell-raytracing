''' Create a new Lambertian reflection material.
Function Lambertian_New(Albedo)
    Dim Result(1)
    Result(0) = "Lambertian"
    Result(1) = Albedo

    Lambertian_New = Result
End Function

''' Get the albedo of the material.
Function Lambertian_Albedo(Lambertian)
    Lambertian_Albedo = Lambertian(1)
End Function

Function Lambertian_Scatter(Lambertian, IncidentRay, HitResult)
    Dim ScatterDirection, ScatteredRay, Attenuation
    ScatterDirection = Vector_Add(HitResult_Normal(HitResult), Math_RandomUnitVector())

    ' Catch degenerate scatter direction
    If Math_IsNearZeroVector(ScatterDirection) Then
        Debug_Log("Lambertian_Scatter: Degenerate scatter direction detected")
        ScatterDirection = HitResult_Normal(HitResult)
    End If

    ScatteredRay = Ray_New(HitResult_Point(HitResult), ScatterDirection)
    Attenuation = Lambertian_Albedo(Lambertian)

    Lambertian_Scatter = ScatterResult_Scatter(ScatteredRay, Attenuation)
End Function