Attribute VB_Name = "modmain"
Public csSelectedColor As Long

Function clamp(ByVal number As Double, ByVal min As Double, ByVal max As Double)
    If number < min Then
        clamp = min
    ElseIf number > max Then
        clamp = max
    Else
        clamp = number
    End If
End Function

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
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 1: ret(0, 4) = 0.2
        ret(1, 1) = 0: ret(1, 2) = 0.2: ret(1, 3) = 1: ret(1, 4) = 0.8
    ElseIf number = 2 Then
        ReDim ret(1, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 1
        ret(1, 1) = 0.5: ret(1, 2) = 0: ret(1, 3) = 0.5: ret(1, 4) = 1
    ElseIf number = 3 Then
        ReDim ret(1, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.3: ret(0, 4) = 1
        ret(1, 1) = 0.3: ret(1, 2) = 0: ret(1, 3) = 0.7: ret(1, 4) = 1
    ElseIf number = 4 Then
        ReDim ret(1, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.7: ret(0, 4) = 1
        ret(1, 1) = 0.7: ret(1, 2) = 0: ret(1, 3) = 0.3: ret(1, 4) = 1
    ElseIf number = 5 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.33: ret(0, 4) = 1
        ret(1, 1) = 0.33: ret(1, 2) = 0: ret(1, 3) = 0.34: ret(1, 4) = 1
        ret(2, 1) = 0.67: ret(2, 2) = 0: ret(2, 3) = 0.33: ret(2, 4) = 1
    ElseIf number = 6 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.2: ret(0, 4) = 1
        ret(1, 1) = 0.2: ret(1, 2) = 0: ret(1, 3) = 0.6: ret(1, 4) = 1
        ret(2, 1) = 0.8: ret(2, 2) = 0: ret(2, 3) = 0.2: ret(2, 4) = 1
    ElseIf number = 7 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 1
        ret(1, 1) = 0.5: ret(1, 2) = 0: ret(1, 3) = 0.5: ret(1, 4) = 0.5
        ret(2, 1) = 0.5: ret(2, 2) = 0.5: ret(2, 3) = 0.5: ret(2, 4) = 0.5
    ElseIf number = 8 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 0.5
        ret(1, 1) = 0: ret(1, 2) = 0.5: ret(1, 3) = 0.5: ret(1, 4) = 0.5
        ret(2, 1) = 0.5: ret(2, 2) = 0: ret(2, 3) = 0.5: ret(2, 4) = 1
    ElseIf number = 9 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.45: ret(0, 4) = 1
        ret(1, 1) = 0.45: ret(1, 2) = 0: ret(1, 3) = 0.55: ret(1, 4) = 0.3
        ret(2, 1) = 0.45: ret(2, 2) = 0.3: ret(2, 3) = 0.55: ret(2, 4) = 0.7
    ElseIf number = 10 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.55: ret(0, 4) = 0.3
        ret(1, 1) = 0: ret(1, 2) = 0.3: ret(1, 3) = 0.55: ret(1, 4) = 0.7
        ret(2, 1) = 0.55: ret(2, 2) = 0: ret(2, 3) = 0.45: ret(2, 4) = 1
    ElseIf number = 11 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.65: ret(0, 4) = 1
        ret(1, 1) = 0.65: ret(1, 2) = 0: ret(1, 3) = 0.35: ret(1, 4) = 0.5
        ret(2, 1) = 0.65: ret(2, 2) = 0.5: ret(2, 3) = 0.35: ret(2, 4) = 0.5
    ElseIf number = 12 Then
        ReDim ret(2, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.35: ret(0, 4) = 0.5
        ret(1, 1) = 0: ret(1, 2) = 0.5: ret(1, 3) = 0.35: ret(1, 4) = 0.5
        ret(2, 1) = 0.35: ret(2, 2) = 0: ret(2, 3) = 0.65: ret(2, 4) = 1
    ElseIf number = 13 Then
        ReDim ret(3, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 0.5
        ret(1, 1) = 0: ret(1, 2) = 0.5: ret(1, 3) = 0.5: ret(1, 4) = 0.5
        ret(2, 1) = 0.5: ret(2, 2) = 0: ret(2, 3) = 0.5: ret(2, 4) = 0.5
        ret(3, 1) = 0.5: ret(3, 2) = 0.5: ret(3, 3) = 0.5: ret(3, 4) = 0.5
    ElseIf number = 14 Then
        ReDim ret(3, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 0.35
        ret(1, 1) = 0: ret(1, 2) = 0.35: ret(1, 3) = 0.5: ret(1, 4) = 0.65
        ret(2, 1) = 0.5: ret(2, 2) = 0: ret(2, 3) = 0.5: ret(2, 4) = 0.65
        ret(3, 1) = 0.5: ret(3, 2) = 0.65: ret(3, 3) = 0.5: ret(3, 4) = 0.35
    ElseIf number = 15 Then
        ReDim ret(3, 4) As Double
        ret(0, 1) = 0: ret(0, 2) = 0: ret(0, 3) = 0.5: ret(0, 4) = 0.65
        ret(1, 1) = 0: ret(1, 2) = 0.65: ret(1, 3) = 0.5: ret(1, 4) = 0.35
        ret(2, 1) = 0.5: ret(2, 2) = 0: ret(2, 3) = 0.5: ret(2, 4) = 0.35
        ret(3, 1) = 0.5: ret(3, 2) = 0.35: ret(3, 3) = 0.5: ret(3, 4) = 0.65
    End If
    GetLayout = ret
End Function


Public Sub PaintPic(ByVal picPath As String, ByRef box As PictureBox, Optional ByVal Stretch As Boolean = True)
    On Error GoTo err
    If picPath <> "" Then
        Dim pic As StdPicture
        If Dir(picPath) <> "" Then
            Set pic = LoadPicture(picPath)
            
            box.Cls
            If Stretch Then
                PaintStretchedPicture pic, box
            Else
                PaintScaledPicture pic, box
            End If
        End If
    End If
    Exit Sub
err:    MsgBox "The picture could not be loaded"
End Sub

Public Sub PaintScaledPicture(ByRef pic As StdPicture, ByRef box As PictureBox)
    Dim ar As Double
    Dim box_ar As Double
    ar = 1# * pic.Width / pic.Height
    box_ar = 1# * box.ScaleWidth / box.ScaleHeight
    
    box.PaintPicture pic, _
    clamp((0.5 - ar / box_ar / 2#), 0, 1) * box.ScaleWidth, _
    clamp((0.5 - box_ar / ar / 2#), 0, 1) * box.ScaleHeight, _
    clamp(ar / box_ar, 0, 1) * box.ScaleWidth, _
    clamp(box_ar / ar, 0, 1) * box.ScaleHeight
End Sub

Public Sub PaintStretchedPicture(ByRef pic As StdPicture, ByRef box As PictureBox)
    Dim ar As Double
    Dim box_ar As Double
    Dim rr As Double
    
    ar = 1# * pic.Width / pic.Height
    box_ar = 1# * box.ScaleWidth / box.ScaleHeight
    rr = ar / box_ar
    
    box.PaintPicture pic, _
    clamp((0.5 - rr / 2#), -rr, 0) * box.ScaleWidth, _
    clamp((0.5 - 1 / rr / 2#), 1 / -rr, 0) * box.ScaleHeight, _
    clamp(rr, 1, rr) * box.ScaleWidth, _
    clamp(1 / rr, 1, 1 / rr) * box.ScaleHeight
End Sub

Public Function EscapeBackslashes(ByVal s As String) As String
    EscapeBackslashes = Replace(s, "\", "\\")
End Function
