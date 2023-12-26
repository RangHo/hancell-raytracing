''' Create a new metallic material.
Function Metal_New(Albedo)
    Dim Result(1)
    Result(0) = "Metal"
    Result(1) = Albedo

    Metal_New = Result
End Function

''' Get the albedo of the material.
Function Metal_Albedo(Metal)
    Metal_Albedo = Metal(1)
End Function

Function Metal_Scatter(Metal, IncidentRay, HitResult)
    Dim Reflected, ScatteredRay, Attenuation
    Reflected = Math_ReflectVector(Vector_Normalize(Ray_Direction(IncidentRay)), HitResult_Normal(HitResult))
    ScatteredRay = Ray_New(HitResult_Point(HitResult), Reflected)
    Attenuation = Metal_Albedo(Metal)

    ' Check if the reflected ray is pointing "outwards"
    Dim DotProduct
    DotProduct = Vector_Dot(Ray_Direction(ScatteredRay), HitResult_Normal(HitResult))
    If DotProduct > 0 Then
        Debug_Log("Metal_Scatter: Reflected ray is pointing outwards")
        Metal_Scatter = ScatterResult_Scatter(ScatteredRay, Attenuation)
    Else
        Debug_Log("Metal_Scatter: Reflected ray is pointing inwards")
        Metal_Scatter = ScatterResult_Absorb()
    End If
End Function