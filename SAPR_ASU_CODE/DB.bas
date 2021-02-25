'------------------------------------------------------------------------------------------------------------
' Module        : DB - База данных прайс листов и избранного
' Author        : gtfox
' Date          : 2021.02.22
' Description   : База данных прайс листов, избранного и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

'Активация формы выбора элементов схемы из БД
Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
    Load frmDB
    frmDB.Run vsoShape 'Передали его в форму
End Sub

'Function returns DBEngine for current Office Engine Type (DAO.DBEngine.60 or DAO.DBEngine.120)
Public Function GetDBEngine() As Object
    Dim engine As Object
        On Error GoTo EX
        Set GetDBEngine = DBEngine
    Exit Function
EX:
    Set GetDBEngine = CreateObject("DAO.DBEngine.120")
End Function

'Получаем Recordset по запросу
Public Function GetRecordSet(DBName As String, SQLQuery As String) As DAO.Recordset
    Dim dbs As DAO.Database
    Set dbs = GetDBEngine.OpenDatabase(ThisDocument.path & DBName)
    Set GetRecordSet = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)
    Set dbs = Nothing
End Function

'Заполняет ComboBox таблицами/запросами из БД
Public Sub Fill_ComboBox(DBName As String, SQLQuery As String, cmbx As ComboBox, Optional ByVal Skip As Boolean = False)
    Dim rst As DAO.Recordset
    Dim i As Integer
    Set rst = GetRecordSet(DBName, SQLQuery)
    cmbx.Clear
    cmbx.ColumnCount = 1
    i = 0
    With rst
    If .EOF Then Exit Sub
        .MoveFirst
        If Skip Then .MoveNext 'Пропускаем первый элемент
        Do Until .EOF
            cmbx.AddItem .Fields(1).Value
            cmbx.List(i, 1) = .Fields(0).Value
            i = i + 1
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub
