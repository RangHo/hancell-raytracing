Function Debug__VarTypeString_Simple(VarTypeIndex)
    Select Case VarTypeIndex
        Case 0
            Debug__VarTypeString_Simple = "Uninitialized"
        Case 1
            Debug__VarTypeString_Simple = "Null"
        Case 2
            Debug__VarTypeString_Simple = "Integer"
        Case 3
            Debug__VarTypeString_Simple = "Long Integer"
        Case 4
            Debug__VarTypeString_Simple = "Single-precision Floating Point"
        Case 5
            Debug__VarTypeString_Simple = "Double-precision Floating Point"
        Case 6
            Debug__VarTypeString_Simple = "Currency"
        Case 7
            Debug__VarTypeString_Simple = "Date"
        Case 8
            Debug__VarTypeString_Simple = "String"
        Case 9
            Debug__VarTypeString_Simple = "Object"
        Case 10
            Debug__VarTypeString_Simple = "Error"
        Case 11
            Debug__VarTypeString_Simple = "Boolean"
        Case 12
            Debug__VarTypeString_Simple = "Variant"
        Case 13
            Debug__VarTypeString_Simple = "Data-access Object"
        Case 14
            Debug__VarTypeString_Simple = "Decimal"
        Case 17
            Debug__VarTypeString_Simple = "Byte"
        Case 20
            Debug__VarTypeString_Simple = "Long Long Integer"
        Case Else
            Debug__VarTypeString_Simple = "Unknown"
    End Select
End Function

Function Debug_VarTypeString(Obj)
    Dim VarTypeIndex
    VarTypeIndex = VarType(Obj)

    ' If index >= 8192, then it is an array
    If VarTypeIndex >= 8192 Then
        Dim ContentTypeIndex, ContentTypeString
        ContentTypeIndex = VarTypeIndex - 8192
        ContentTypeString = Debug__VarTypeString_Simple(ContentTypeIndex)

        Debug_VarTypeString = "Array of " & ContentTypeString
    Else
        Debug_VarTypeString = Debug__VarTypeString_Simple(ContentTypeIndex)
    End If
End Function