''' Create a new hit result.
Function HitResult_New(IsHit, Point, Normal, T)
    Dim Result(4)
    Result(0) = IsHit
    Result(1) = Point
    Result(2) = Normal
    Result(3) = T

    HitResult_New = Result
End Function

''' Create a new "hit" hit result.
Function HitResult_Hit(Point, Normal, T)
    Dim Result
    Result = HitResult_New(True, Point, Normal, T)

    HitResult_Hit = Result
End Function

''' Create a new "missed" hit result.
Function HitResult_Miss()
    Dim Result
    Result = HitResult_New(False, Vector_New(0, 0, 0), Vector_New(0, 0, 0), 0)

    HitResult_Miss = Result
End Function

''' Check if the hit result contains a valid hit.
Function HitResult_IsHit(HitResult)
    HitResult_IsHit = HitResult(0)
End Function

''' Get where the ray hit the object.
Function HitResult_Point(HitResult)
    HitResult_Point = HitResult(1)
End Function

''' Get the normal vector of the hit point.
Function HitResult_Normal(HitResult)
    HitResult_Normal = HitResult(2)
End Function

''' Get the T value of the hit.
Function HitResult_T(HitResult)
    HitResult_T = HitResult(3)
End Function

''' Get the face normal of the hit point.
Function HitResult_FaceNormal(HitResult, Ray)
    Dim IsFrontFacing, Normal
    IsFrontFacing = Vector_Dot(Ray_Direction(Ray), HitResult_Normal(HitResult)) < 0.0
    
    If IsFrontFacing Then
        HitResult_FaceNormal = HitResult_Normal(HitResult)
    Else
        HitResult_FaceNormal = Vector_Scale(HitResult_Normal(HitResult), -1.0)
    End If
End Function

''' Get the type string of the hittable object.
Function Hittable_Type(Hittable)
    Hittable_Type = Hittable(0)
End Function

''' Hit a hittable object.
Function Hit(Object, Ray, TMin, TMax)
    Select Case Hittable_Type(Object)
        Case "Sphere"
            Hit = Sphere_Hit(Object, Ray, TMin, TMax)
        Case "HittableList"
            Hit = HittableList_Hit(Object, Ray, TMin, TMax)
        Case Else
            Debug_Log("Unknown hittable object type: " & Hittable_Type(Object))
            Hit = HitResult_Miss()
    End Select
End Function