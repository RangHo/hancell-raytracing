''' Epsilon value for floating point comparisons.
Const Math_Epsilon = 0.0000001

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
Function Math_RandomUnitVector()
    Dim Theta, Phi
    Theta = 360 * Rnd()
    Phi = 360 * Rnd()
    Math_RandomUnitVector = Vector_New(Sin(Theta) * Cos(Phi), Sin(Theta) * Sin(Phi), Cos(Theta))
End Function

''' Create a random vector with length between 0 and 1.
Function Math_RandomVector()
    Dim V, R
    V = Math_RandomUnitVector()
    R = Rnd()
    Math_RandomVector = Vector_Scale(V, R)
End Function

''' Check if a vector is near zero in all dimensions.
Function Math_VectorIsNearZero(V)
    Dim XIsZero, YIsZero, ZIsZero
    XIsZero = Abs(Vector_X(V)) < Math_Epsilon
    YIsZero = Abs(Vector_Y(V)) < Math_Epsilon
    ZIsZero = Abs(Vector_Z(V)) < Math_Epsilon
    Math_VectorIsNearZero = XIsZero And YIsZero And ZIsZero
End Function
