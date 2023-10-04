''' Get the X component of the vector.
Function Vector_X(V)
    Vector_X = V(1)
End Function

''' Get the Y component of the vector.
Function Vector_Y(V)
    Vector_Y = V(2)
End Function

''' Get the Z component of the vector.
Function Vector_Z(V)
    Vector_Z = V(3)
End Function

''' Allocate a new 3D vector.
Function Vector_New(X, Y, Z)
    ' Allocate new array
    Dim Result(3)
    Result(1) = X
    Result(2) = Y
    Result(3) = Z

    ' Return the array
    Vector_New = Result
End Function

''' Add two vectors together.
Function Vector_Add(LHS, RHS)
    ' Calculate each component
    Dim X, Y, Z
    X = Vector_X(LHS) + Vector_X(RHS)
    Y = Vector_Y(LHS) + Vector_Y(RHS)
    Z = Vector_Z(LHS) + Vector_Z(RHS)

    ' Return the new vector
    Vector_Add = Vector_New(X, Y, Z)
End Function

''' Subtract the RHS vector from the LHS vector.
Function Vector_Subtract(LHS, RHS)
    ' Calculate each component
    Dim X, Y, Z
    X = Vector_X(LHS) - Vector_X(RHS)
    Y = Vector_Y(LHS) - Vector_Y(RHS)
    Z = Vector_Z(LHS) - Vector_Z(RHS)

    ' Return the new vector
    Vector_Subtract = Vector_New(X, Y, Z)
End Function

''' Scale the vector by a scalar.
Function Vector_Scale(V, S)
    ' Calculate each component
    Dim X, Y, Z
    X = Vector_X(V) * S
    Y = Vector_Y(V) * S
    Z = Vector_Z(V) * S

    ' Return the new vector
    Vector_Scale = Vector_New(X, Y, Z)
End Function

''' Calculate the dot product of two vectors.
Function Vector_Dot(LHS, RHS)
    ' Calculate each component
    Dim X, Y, Z
    X = Vector_X(LHS) * Vector_X(RHS)
    Y = Vector_Y(LHS) * Vector_Y(RHS)
    Z = Vector_Z(LHS) * Vector_Z(RHS)

    ' Return the dot product
    Vector_Dot = X + Y + Z
End Function

''' Calculate the cross product of two vectors.
Function Vector_Cross(LHS, RHS)
    ' Calculate each component
    Dim X, Y, Z
    X = Vector_Y(LHS) * Vector_Z(RHS) - Vector_Z(LHS) * Vector_Y(RHS)
    Y = Vector_Z(LHS) * Vector_X(RHS) - Vector_X(LHS) * Vector_Z(RHS)
    Z = Vector_X(LHS) * Vector_Y(RHS) - Vector_Y(LHS) * Vector_X(RHS)

    ' Return the cross product
    Vector_Cross = Vector_New(X, Y, Z)
End Function

''' Calculate the squared length of the vector.
Function Vector_LengthSquared(V)
    ' Calculate each component
    Dim X, Y, Z
    X = Vector_X(V) * Vector_X(V)
    Y = Vector_Y(V) * Vector_Y(V)
    Z = Vector_Z(V) * Vector_Z(V)

    ' Return the squared length
    Vector_LengthSquared = X + Y + Z
End Function

''' Calculate the length of the vector.
Function Vector_Length(V)
    Vector_Length = Sqr(Vector_LengthSquared(V))
End Function

''' Return a normalized version of the vector.
Function Vector_Normalize(V)
    ' Calculate the length
    Dim Length
    Length = Vector_Length(V)

    ' Return the normalized vector
    Vector_Normalize = Vector_Scale(V, 1 / Length)
End Function

''' Convert a vector to a string representation
Function Vector_ToString(V)
    Vector_ToString = "(" & Vector_X(V) & ", " & Vector_Y(V) & ", " & Vector_Z(V) & ")"
End Function