''' Cell to dump the log outputs.
Dim LogOutputRow

''' Determine the string representation of simple VarType indices.
'''
''' This function is meant to be private.
Function Debug__VarTypeToString_Simple(VarTypeIndex)
    Select Case VarTypeIndex
        Case 0
            Debug__VarTypeToString_Simple = "Uninitialized"
        Case 1
            Debug__VarTypeToString_Simple = "Null"
        Case 2
            Debug__VarTypeToString_Simple = "Integer"
        Case 3
            Debug__VarTypeToString_Simple = "Long Integer"
        Case 4
            Debug__VarTypeToString_Simple = "Single-precision Floating Point"
        Case 5
            Debug__VarTypeToString_Simple = "Double-precision Floating Point"
        Case 6
            Debug__VarTypeToString_Simple = "Currency"
        Case 7
            Debug__VarTypeToString_Simple = "Date"
        Case 8
            Debug__VarTypeToString_Simple = "String"
        Case 9
            Debug__VarTypeToString_Simple = "Object"
        Case 10
            Debug__VarTypeToString_Simple = "Error"
        Case 11
            Debug__VarTypeToString_Simple = "Boolean"
        Case 12
            Debug__VarTypeToString_Simple = "Variant"
        Case 13
            Debug__VarTypeToString_Simple = "Data-access Object"
        Case 14
            Debug__VarTypeToString_Simple = "Decimal"
        Case 17
            Debug__VarTypeToString_Simple = "Byte"
        Case 20
            Debug__VarTypeToString_Simple = "Long Long Integer"
        Case Else
            Debug__VarTypeToString_Simple = "Unknown"
    End Select
End Function

''' Determine the string representation of a VarType index.
Function Debug_VarTypeToString(VarTypeIndex)
    ' If index >= 8192, then it is an array
    If VarTypeIndex >= 8192 Then
        Dim ContentTypeIndex, ContentTypeString
        ContentTypeIndex = VarTypeIndex - 8192
        ContentTypeString = Debug__VarTypeToString_Simple(ContentTypeIndex)

        Debug_VarTypeToString = "Array of " & ContentTypeString
    Else
        Debug_VarTypeToString = Debug__VarTypeToString_Simple(ContentTypeIndex)
    End If
End Function

''' Write log output to a dedicated cell.
'''
''' This function will simply return if `IsDebugging` switch in `Config` module
''' is set to false.
Function Debug_Log(Content)
    ' Log only when debugging
    If IsDebugging Then
        ' Check if the LogOutputRow is initialized
        If VarType(LogOutputRow) = 0 Then
            ' Initialize the LogOutputRow
            LogOutputRow = 1
        End If

        ' Merge the current row
        Sheets("Log").Range("A" & LogOutputRow & ":Z" & LogOutputRow).MergeCells = True

        ' Write the content to the merged cell
        Sheets("Log").Range("A" & LogOutputRow).Formula = Content

        ' Bump to the next row
        LogOutputRow = LogOutputRow + 1
    End If
End Function

''' Show a MsgBox for debugging.
'''
''' This function will simply return if `IsDebugging` switch in `Config` module
''' is set to false.
Function Debug_MsgBox(Content)
    If IsDebugging Then
        MsgBox(Content)
    End If
End Function
