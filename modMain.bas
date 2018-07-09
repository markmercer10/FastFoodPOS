Attribute VB_Name = "modmain"
Public csSelectedColor As Long

Function stringValidate(ByVal s As String) As Boolean
    stringValidate = False
    If s <> "" And Not containsIllegalCharacters(s) Then
        stringValidate = True
    End If
End Function

Function stringValidateSQL(ByVal s As String) As Boolean
    stringValidateSQL = False
    If s <> "" And Not containsIllegalSQLCharacters(s) Then
        stringValidateSQL = True
    End If
End Function

Function containsIllegalCharacters(ByVal s As String) As Boolean
    Dim i As Long
    For i = 0 To 31
        containsIllegalCharacters = InStr(1, s, Chr(i)) > 0
        If containsIllegalCharacters Then Exit For
    Next i
    If Not containsIllegalCharacters Then containsIllegalCharacters = InStr(1, s, Chr(127)) > 0
End Function

Function containsIllegalFilenameCharacters(ByVal s As String) As Boolean
    'containsIllegalFileCharacters = containsIllegalCharacters(s)
    'If Not containsIllegalFileCharacters Then
    '    Dim i As Long
    '    For i = other illegal characters in ascii table
    '        containsIllegalFileCharacters = InStr(1, s, Chr(i)) > 0
    '        If containsIllegalFileCharacters Then Exit For
    '    Next i
    'End If
    MsgBox "Not Implemented"
End Function

Function containsIllegalSQLCharacters(ByVal s As String) As Boolean
    containsIllegalSQLCharacters = containsIllegalCharacters(s)
    If Not containsIllegalSQLCharacters Then
        Dim i As Long
        Dim illegal
        illegal = Array(34, 39, 44, 96)
        For i = 1 To 2
            containsIllegalSQLCharacters = InStr(1, s, Chr(illegal(i))) > 0
            If containsIllegalSQLCharacters Then Exit For
        Next i
    End If
End Function


Function GetLayout(ByVal number As Long) As Double()
    Dim ret() As Double
    If number = 0 Then
        ReDim ret(0, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 1: ret(0, 4) = 1
    ElseIf number = 1 Then
        ReDim ret(1, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 1
        ret(1, 1) = 0.5: ret(1, 2) = 0: ret(1, 3) = 0.5: ret(1, 4) = 1
    ElseIf number = 2 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.33: ret(0, 4) = 1
        ret(1, 1) = 0.33: ret(1, 2) = 0: ret(1, 3) = 0.34: ret(1, 4) = 1
        ret(2, 1) = 0.67: ret(2, 2) = 0: ret(2, 3) = 0.33: ret(2, 4) = 1
    ElseIf number = 3 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.2: ret(0, 4) = 1
        ret(1, 1) = 0.2: ret(1, 2) = 0: ret(1, 3) = 0.6: ret(1, 4) = 1
        ret(2, 1) = 0.8: ret(2, 2) = 0: ret(2, 3) = 0.2: ret(2, 4) = 1
    ElseIf number = 4 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 1
        ret(1, 1) = 0.5: ret(1, 2) = 0: ret(1, 3) = 0.5: ret(1, 4) = 0.5
        ret(2, 1) = 0.5: ret(2, 2) = 0.5: ret(2, 3) = 0.5: ret(2, 4) = 0.5
    ElseIf number = 5 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 0.5
        ret(1, 1) = 0: ret(1, 2) = 0.5: ret(1, 3) = 0.5: ret(1, 4) = 0.5
        ret(2, 1) = 0.5: ret(2, 2) = 0: ret(2, 3) = 0.5: ret(2, 4) = 1
    ElseIf number = 6 Then
        ReDim ret(3, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 0.5
        ret(1, 1) = 0: ret(1, 2) = 0.5: ret(1, 3) = 0.5: ret(1, 4) = 0.5
        ret(2, 1) = 0.5: ret(2, 2) = 0: ret(2, 3) = 0.5: ret(2, 4) = 0.5
        ret(3, 1) = 0.5: ret(3, 2) = 0.5: ret(3, 3) = 0.5: ret(3, 4) = 0.5
    End If
    GetLayout = ret
End Function

