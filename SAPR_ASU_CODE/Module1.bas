Public Sub SQL_Query_To_Smart_Table(Table_SourceAddress As String)
    'ВАЖНО: есть ограничение SQL при указании адреса диапазона (A1:Z100) в 65536 строк, либо надо указывать полностью A:Z
    Dim sTblQuery As String
    Dim SheetName As String
'    Dim Table_SourceAddress As String
    Dim sSQL_text As String
    Dim sConnStr As String
    Dim dtData1 As Date
    Dim dtData2 As Date
    Dim i As Long
     
    SheetName = "Лист2"

'    Table_SourceAddress = Worksheets(SheetName).ListObjects("Таблица1").Range.Address(0, 0)
    sTblQuery = "[" & SheetName & "$" & Table_SourceAddress & "]"
     
    'Вариант 1
'    Dim oRecordSet As Object
'    Dim oConn As Object
'    Set oRecordSet = CreateObject("ADODB.Recordset")
'    Set oConn = CreateObject("ADODB.Connection")
     
    'Вариант 2
    'нужно добавить ссылку на Microsoft ActiveX Data Objects 6.1 Library, то
        Dim oConn As New ADODB.Connection
        Dim oRecordSet As New ADODB.Recordset
      
    oRecordSet.CursorLocation = 3 'adUseClient ' включает возможность order by
    'Вариант 1 с использованием функции Format
    'sSQL_text = "SELECT * FROM " & sTblQuery & " WHERE ([Дата] >= #" & Format(dtData1, "MM\/dd\/yy hh\:mm\:ss") & "#) AND ([Дата] <= #" & Format(dtData2, "MM\/dd\/yy hh\:mm\:ss") & "#)"
    'Вариант 2 с использованием функции DataSql
    sSQL_text = "SELECT * FROM " & sTblQuery '& " WHERE ([Дата] >= " & DataSql(dtData1) & ")" & " AND ([Дата] <= " & DataSql(dtData2) & ")"
      
    sConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Mode=Read;Data Source=" & ActiveDocument.path & "SAPR_ASU_EKF.xls" & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    oConn.Open sConnStr
    oRecordSet.Open sSQL_text, oConn
'    oRecordSet.Sort = "[Дата] ASC,[Смена] DESC" 'order by
     oRecordSet.MoveFirst
        'заголовки таблицы
        For i = 0 To oRecordSet.Fields.Count - 1
            sName = oRecordSet.Fields(0).name
            sValue = oRecordSet.Fields(0).Value
            oRecordSet.MoveNext
        Next i
        oRecordSet.MoveLast
        ii = oRecordSet.RecordCount

    oRecordSet.Close
    oConn.Close
    MsgBox "Запрос выполнен!", vbInformation, ""
End Sub
 
Private Function DataSql(dt_sql)
    DataSql = "#" & Format(dt_sql, "mm\/dd\/yy hh\:mm\:ss") & "#"
End Function
 
