Attribute VB_Name = "modmain"
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

Function containsIllegalFileCharacters(ByVal s As String) As Boolean
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

