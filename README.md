# vba-sql-formatter
A4セルにSQLを貼り付けてFormatボタンをクリックするとB4セルに整形後のSQLが出力されます。
```
Sub ButtonClick()
    Dim sql As String
    sql = FormatSQL(Range("A4").Value)
    
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
    sql = SpaceReplace(sql)
    
    sql = SqlReplace(sql, "SELECT", False, True, False)
    sql = SqlReplace(sql, ",", False, True, False)
    sql = SqlReplace(sql, "FROM", True, True, True)
    sql = SqlReplace(sql, "WHERE", True, True, True)
    sql = SqlReplace(sql, "GROUP", True, True, True)
    sql = SqlReplace(sql, "HAVING", True, True, True)
    sql = SqlReplace(sql, "ORDER BY", True, True, True)
    sql = SqlReplace(sql, "INNER JOIN ", True, False, True)
    sql = SqlReplace(sql, "LEFT JOIN", True, False, True)
    sql = SqlReplace(sql, "FULL JOIN", True, False, True)
    sql = SqlReplace(sql, "ON", False, True, True)
    
    sql = SqlReplace(sql, "INSERT INTO", False, True, False)
    sql = SqlReplace(sql, "VALUES", True, True, True)
    
    sql = SqlReplace(sql, "UPDATE", False, True, False)
    sql = SqlReplace(sql, "SET", True, True, True)
    
    sql = SqlReplace(sql, "DELETE", False, True, False)
    
    sql = ParenthesesReplace(sql)

    FormatSQL = sql
End Function

Function SqlReplace(sql As String, keyword As String, addBreakBefore As Boolean, addBreakAfter As Boolean, useRegular As Boolean) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    
    If useRegular Then
        ' 前後にスペースまたは改行がある場合のみ改行する
        RegEx.Pattern = "\s+" & keyword & "\s+"
    Else
        RegEx.Pattern = keyword
    End If
    
    RegEx.Global = True
    
    Dim replacedText As String
    replacedText = keyword
  
    
    If addBreakBefore Then
        replacedText = vbCrLf & replacedText
    End If
    
    If addBreakAfter Then
        replacedText = replacedText & vbCrLf
    End If
    
    sql = RegEx.Replace(sql, replacedText)
    
    SqlReplace = sql
    
End Function

' 複数のスペースを1つのスペースに置換
Function SpaceReplace(sql As String) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = "\s+"
    RegEx.Global = True
    sql = RegEx.Replace(sql, " ")
    SpaceReplace = sql
End Function

' ()を改行
Function ParenthesesReplace(sql As String) As String
    sql = Replace(sql, "(", vbCrLf & "(")
    sql = Replace(sql, ")", ")" & vbCrLf)
    ParenthesesReplace = sql
End Function
```
