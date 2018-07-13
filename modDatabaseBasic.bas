Attribute VB_Name = "modDatabaseBasic"
Public db As ADODB.Connection

Sub ConnectDB()
    Set db = New ADODB.Connection
    Dim prefix As String
    
    If RIDE Then
        prefix = "dev"
    Else
        prefix = "release"
    End If
    
    If Dir(App.Path & "\connection.txt") <> "" Then
        Dim fileline As String
        Open App.Path & "\connection.txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, fileline
                If left$(fileline, Len(prefix)) = prefix Then
                    db.ConnectionString = Trim(Mid$(fileline, Len(prefix) + 2))
                    Exit Do
                End If
            Loop
        Close #1
    End If
    
    db.Open
End Sub

Function sqlDate(val As Variant) As String
    Dim d As Date
    If IsNull(val) Then
        sqlDate = ""
    Else
        d = CDate(val)
        sqlDate = """" & Format(d, "YYYY-MM-DD") & """"
    End If
End Function

Function sqlTime(val As Variant) As String
    Dim d As Date
    d = CDate(val)
    sqlTime = """" & Format(d, "hh:nn:ss") & """"
   
End Function

Sub Delete(ByVal Table As String, ByVal primaryKeyField As String, ByVal id As Long)
    db.Execute "DELETE FROM " & Table & " WHERE " & primaryKeyField & " = " & id
End Sub

Sub Upsert(ByVal Table As String, ByRef fields As Variant, ByRef values As Variant)
    ' fields and values are arrays of strings but vb would NOT allow me to assign those arrays in a single line without making them variants!!!!!!
    Dim sql As String
    Dim i As Long
    If LBound(fields) <> 0 Or LBound(values) <> 0 Then
        MsgBox "the first index in the array must be 0 and this must be the primary key of the table"
        Exit Sub
    End If
    
    sql = ""
    sql = sql & "INSERT INTO " & Table & "("
    For i = 0 To UBound(fields)
        sql = sql & fields(i)
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    sql = sql & ") values ("
    For i = 0 To UBound(fields)
        If VarType(values(i)) = vbString Then
            sql = sql & """" & Replace(values(i), """", "'") & """"
        ElseIf VarType(values(i)) = vbDate Then
            sql = sql & sqlDate(values(i))
        Else
            sql = sql & "'" & Replace(values(i), "'", """") & "'"
        End If
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    sql = sql & ") ON DUPLICATE KEY UPDATE "
    For i = 1 To UBound(fields)
        sql = sql & fields(i) & " = values(" & fields(i) & ")"
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    'MsgBox sql
    Clipboard.SetText sql
    db.Execute sql
End Sub

Sub Upsert_Broken(ByVal Table As String, ByRef fields As Variant, ByRef values As Variant)
    ' fields and values are arrays of strings but vb would NOT allow me to assign those arrays in a single line without making them variants!!!!!!
    Dim sql As String
    Dim i As Long
    If LBound(fields) <> 0 Or LBound(values) <> 0 Then
        MsgBox "the first index in the array must be 0 and this must be the primary key of the table"
        Exit Sub
    End If
    
    sql = "SET "
    For i = 0 To UBound(fields)
        If i > UBound(values) Then
            MsgBox "values array is of incorrect length"
            Exit Sub
        End If
        sql = sql & "@" & fields(i) & " = '" & Replace(values(i), """", "'") & "'"
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    sql = sql & "; "
    sql = sql & "INSERT INTO " & Table & "("
    For i = 0 To UBound(fields)
        sql = sql & fields(i)
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    sql = sql & ") values ("
    For i = 0 To UBound(fields)
        sql = sql & "@" & fields(i)
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    sql = sql & ") ON DUPLICATE KEY UPDATE "
    For i = 1 To UBound(fields)
        sql = sql & fields(i) & " = @" & fields(i)
        If i <> UBound(fields) Then sql = sql & ", "
    Next i
    Clipboard.SetText sql
    'MsgBox sql
    db.Execute sql
End Sub


