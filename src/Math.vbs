Function Math_Clamp(Number, MinValue, MaxValue)
    If Number < MinValue Then
        Math_Clamp = MinValue
    ElseIf Number > MaxValue Then
        Math_Clamp = MaxValue
    Else
        Math_Clamp = Number
    End If
End Function
