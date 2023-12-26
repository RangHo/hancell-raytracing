''' Create a new scatter result.
Function ScatterResult_New(IsScattered, ScatteredRay, Attenuation)
    Dim Result(2)
    Result(0) = IsScattered
    Result(1) = ScatteredRay
    Result(2) = Attenuation

    ScatterResult_New = Result
End Function

''' Create a new "scattered" scatter result.
Function ScatterResult_Scatter(ScatteredRay, Attenuation)
    Dim Result
    Result = ScatterResult_New(True, ScatteredRay, Attenuation)

    ScatterResult_Scatter = Result
End Function

''' Create a new "absorbed" scatter result.
Function ScatterResult_Absorb()
    Dim Result
    Result = ScatterResult_New(False, Vector_New(0, 0, 0), Vector_New(0, 0, 0))

    ScatterResult_Absorb = Result
End Function

''' Check if the scatter result contains a valid scatter.
Function ScatterResult_IsScattered(ScatterResult)
    ScatterResult_IsScattered = ScatterResult(0)
End Function

''' Get the scattered ray.
Function ScatterResult_ScatteredRay(ScatterResult)
    ScatterResult_ScatteredRay = ScatterResult(1)
End Function

''' Get the attenuation vector.
Function ScatterResult_Attenuation(ScatterResult)
    ScatterResult_Attenuation = ScatterResult(2)
End Function

Function Material_Type(Material)
    Material_Type = Material(0)
End Function

Function Scatter(Material, IncidentRay, HitResult)
    Select Case Material_Type(Material)
        Case "Lambertian"
            Scatter = Lambertian_Scatter(Material, IncidentRay, HitResult)
        Case Else
            Debug_Log("Scatter: Unknown material type")
            Scatter = ScatterResult_Absorb()
    End Select
End Function
