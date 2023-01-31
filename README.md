# vba-sql-formatter
A4セルにSQLを貼り付けてFormatボタンをクリックするとB4セルに整形後のSQLが出力されます。
```
Sub ButtonClick()
    Dim sql As String
    sql = Range("A4").Value
    sql = EmptyReplace(sql)
    
    sql = FormatSQL(sql)
    
    Do While InStr(sql, vbCrLf & vbCrLf) > 0
        ' 無駄な改行を削除
         sql = Replace(sql, vbCrLf & vbCrLf, vbCrLf)
    Loop
    
    ' 余分なスペースを削除
    sql = SpaceReplace(sql)
    
    Dim lines() As String
    lines = Split(sql, vbCrLf)
    For i = 0 To UBound(lines)
        ' 1行ずつTrimする
        lines(i) = Trim(lines(i))
    Next
    
    Range("B4").Value = Join(lines, vbCrLf)
End Sub

Function FormatSQL(sql As String) As String
    sql = UCase(sql)
    
    sql = SqlReplace(sql, "SELECT", " SELECT " & vbCrLf, False)
    sql = SqlReplace(sql, ",", "," & vbCrLf, False)
    sql = SqlReplace(sql, "FROM", vbCrLf & " FROM " & vbCrLf, True)
    sql = SqlReplace(sql, "WHERE", vbCrLf & " WHERE " & vbCrLf, True)
    sql = SqlReplace(sql, "GROUP BY", vbCrLf & " GROUP BY " & vbCrLf, True)
    sql = SqlReplace(sql, "HAVING", vbCrLf & " HAVING " & vbCrLf, True)
    sql = SqlReplace(sql, "ORDER BY", vbCrLf & " ORDER BY " & vbCrLf, True)
    sql = SqlReplace(sql, "INNER JOIN", vbCrLf & " INNER JOIN ", True)
    sql = SqlReplace(sql, "LEFT JOIN", vbCrLf & " LEFT JOIN ", True)
    sql = SqlReplace(sql, "FULL JOIN", vbCrLf & " FULL JOIN ", True)
    sql = SqlReplace(sql, "ON", "ON" & vbCrLf, True)
    sql = SqlReplace(sql, "AND", vbCrLf & " AND ", True)
    sql = SqlReplace(sql, "OR", vbCrLf & " OR ", True)
    
    sql = SqlReplace(sql, "INSERT INTO", " INSERT INTO " & vbCrLf, False)
    sql = SqlReplace(sql, "VALUES", vbCrLf & " VALUES " & vbCrLf, False)
    
    sql = SqlReplace(sql, "UPDATE", " UPDATE " & vbCrLf, False)
    sql = SqlReplace(sql, "SET", vbCrLf & " SET " & vbCrLf, True)
    
    sql = SqlReplace(sql, "DELETE", " DELETE " & vbCrLf, False)

    FormatSQL = sql
End Function

Function SqlReplace(sql As String, keyword As String, replacedText As String, hasSpace As Boolean) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    
    If hasSpace Then
        ' 前後にスペースまたは改行がある場合のみ改行する
        RegEx.Pattern = "\s+" & keyword & "\s+"
    Else
        RegEx.Pattern = keyword
    End If
    
    RegEx.Global = True
    
    sql = RegEx.Replace(sql, replacedText)
    
    SqlReplace = sql
    
End Function

' 改行や複数のスペースなどをスペースに置換
Function EmptyReplace(sql As String) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = "\s+"
    RegEx.Global = True
    sql = RegEx.Replace(sql, " ")
    EmptyReplace = sql
End Function

' 複数のスペースを1つのスペースに置換
Function SpaceReplace(sql As String) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = " +"
    RegEx.Global = True
    sql = RegEx.Replace(sql, " ")
    SpaceReplace = sql
End Function
```
