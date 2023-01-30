# vba-sql-formatter
A4セルにSQLを貼り付けてFormatボタンをクリックするとB4セルに整形後のSQLが出力されます。
```
Sub ButtonClick()
    Range("B4").Value = FormatSQL(Range("A4").Value)
    TrimB4Cell
End Sub

Function FormatSQL(sql As String) As String
    sql = UCase(sql)
    sql = Replace(sql, "SELECT", "SELECT" & vbCrLf)
    sql = Replace(sql, "FROM", vbCrLf & "FROM" & vbCrLf)
    sql = Replace(sql, "WHERE", "WHERE" & vbCrLf)
    sql = Replace(sql, "GROUP BY", vbCrLf & "GROUP BY" & vbCrLf)
    sql = Replace(sql, "HAVING", vbCrLf & "HAVING" & vbCrLf)
    sql = Replace(sql, "ORDER BY", vbCrLf & "ORDER BY" & vbCrLf)
    sql = Replace(sql, ",", "," & vbCrLf)
    sql = Replace(sql, " INNER JOIN ", vbCrLf & " INNER JOIN ")
    sql = Replace(sql, " LEFT JOIN ", vbCrLf & " LEFT JOIN ")
    sql = Replace(sql, " RIGHT JOIN ", vbCrLf & " RIGHT JOIN ")
    sql = Replace(sql, " FULL JOIN ", vbCrLf & " FULL JOIN ")
    sql = Replace(sql, " ON ", " ON " & vbCrLf)
    ' OR 前後にスペースまたは改行がある場合のみ改行する
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\s+OR\s+"
    regEx.Global = True
    sql = regEx.Replace(sql, " OR" & vbCrLf)
    
    ' ON 前後にスペースまたは改行がある場合のみ改行する
    regEx.Pattern = "\s+ON\s+"
    regEx.Global = True
    sql = regEx.Replace(sql, " ON" & vbCrLf)
    
    FormatSQL = sql
End Function

Sub TrimB4Cell()
  Dim cellValue As String
  cellValue = Range("B4").Value
  Dim lines() As String
  lines = Split(cellValue, vbCrLf)
  For i = 0 To UBound(lines)
    lines(i) = Trim(lines(i))
  Next
  Range("B4").Value = Join(lines, vbCrLf)
End Sub
```
