''' Convert a [0, 1] color vector into valid [0, 255] RGB value.
Function Color_Vector2RGB(V)
    Dim IR, IG, IB
    IR = Int(255.999 * Vector_X(V))
    IG = Int(255.999 * Vector_Y(V))
    IB = Int(255.999 * Vector_Z(V))

    Color_Vector2RGB = RGB(IR, IG, IB)
End Function