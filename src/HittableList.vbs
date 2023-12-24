Function HittableList_Objects(HittableList)
    HittableList_Objects = HittableList(1)
End Function

Function HittableList_Length(HittableList)
    HittableList_Length = HittableList(2)
End Function

''' Create a new list of hittable objects.
Function HittableList_New(Objects)
    Dim Result(2)
    Result(0) = "HittableList"
    Result(1) = Objects

    HittableList_New = Result
End Function

''' Hit a list of hittable objects.
Function HittableList_Hit(HittableList, Ray, TMin, TMax)
    Dim Things
    Things = HittableList_Objects(HittableList)

    Debug_Log("HittableList_Hit: Hitting a list of hittable objects of size " & UBound(Things))

    Dim ClosestHitResult, ClosestT
    ClosestHitResult = HitResult_Miss()
    ClosestT = TMax

    Dim Index
    For Index = 1 To UBound(Things)
        Debug_Log("HittableList_Hit: Trying to hit object #" & Index & "...")


        Dim HitResult
        HitResult = Hit(Things(Index), Ray, TMin, ClosestT)

        If HitResult_IsHit(HitResult) Then
            ClosestHitResult = HitResult
            ClosestT = HitResult_T(HitResult)
        End If
    Next

    HittableList_Hit = ClosestHitResult
End Function