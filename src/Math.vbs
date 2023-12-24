''' Clamp a number between a minimum and maximum value.
Function Math_Clamp(Number, MinValue, MaxValue)
    If Number < MinValue Then
        Math_Clamp = MinValue
    ElseIf Number > MaxValue Then
        Math_Clamp = MaxValue
    Else
        Math_Clamp = Number
    End If
End Function

''' Create a random vector with length 1.
Function Math_RandomVector()
    Dim Theta, Phi
    Theta = 360 * Rnd()
    Phi = 360 * Rnd()
    Math_RandomVector = Vector_New(Sin(Theta) * Cos(Phi), Sin(Theta) * Sin(Phi), Cos(Theta))
End Function