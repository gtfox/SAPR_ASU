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
    Load frmDBPrice
    frmDBPrice.Run vsoShape 'Передали его в форму
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
Public Function GetRecordSet(DBName As String, SQLQuery As String, Optional QueryDefName As String = "") As DAO.Recordset
    Dim dbs As DAO.Database
    Set dbs = GetDBEngine.OpenDatabase(ThisDocument.path & DBName)
    If QueryDefName <> "" Then
        On Error Resume Next
        dbs.QueryDefs.Delete QueryDefName
    End If
    Set GetRecordSet = dbs.CreateQueryDef(QueryDefName, SQLQuery).OpenRecordset(dbOpenDynaset)
    Set dbs = Nothing
End Function

'Заполняет ComboBox запросами из БД
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
            cmbx.List(i, 1) = "" & IIf(.Fields(0).Value = "", "", .Fields(0).Value)
            i = i + 1
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub

'Заполняет ComboBox Производители запросами из БД
Public Sub Fill_cmbxProizvoditel(DBName As String, SQLQuery As String, cmbx As ComboBox, Optional ByVal Skip As Boolean = False)
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
            If .Fields(0).Value <> "" Then
                cmbx.AddItem .Fields(1).Value
                cmbx.List(i, 1) = .Fields(0).Value
                i = i + 1
            End If
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub

'Заполняет lstvTable запросами из БД
Public Function Fill_lstvTable(DBName As String, SQLQuery As String, QueryDefName As String, lstvTable As ListView, Optional ByVal Proizvoditel As Boolean = False) As Double
    Dim i As Double
    Dim itmx As ListItem
    Dim rst As DAO.Recordset
    Set rst = GetRecordSet(DBName, SQLQuery, QueryDefName)
    lstvTable.ListItems.Clear
    i = 0
    With rst
        If .EOF Then Exit Function
        .MoveFirst
        Do Until .EOF
            Set itmx = lstvTable.ListItems.Add(, """" & .Fields("КодПозиции").Value & """", .Fields("Артикул").Value)
            itmx.SubItems(1) = .Fields("Название").Value
            itmx.SubItems(2) = .Fields("Цена").Value
            If Proizvoditel Then itmx.SubItems(3) = .Fields("Производитель").Value
            i = i + 1
            .MoveNext
        Loop
    End With
    Fill_lstvTable = i
    Set rst = Nothing
End Function

''Заполняет lstvTable запросами из Избранного
'Public Function Fill_lstvTableIzb(DBName As String, SQLQuery As String, QueryDefName As String, lstvTable As ListView, Optional ByVal Proizvoditel As Boolean = False) As Double
'    Dim i As Double
'    Dim itmx As ListItem
'    Dim rst As DAO.Recordset
'    Set rst = GetRecordSet(DBName, SQLQuery, QueryDefName)
'    lstvTable.ListItems.Clear
'    i = 0
'    With rst
'        If .EOF Then Exit Function
'        .MoveFirst
'        Do Until .EOF
'            Set itmx = lstvTable.ListItems.Add(, """" & .Fields("ПозицииКод").Value & """", .Fields("Артикул").Value)
'            itmx.SubItems(1) = .Fields("Название").Value
'            itmx.SubItems(2) = .Fields("Цена").Value
'            If Proizvoditel Then itmx.SubItems(3) = .Fields("Производитель").Value
'            i = i + 1
'            .MoveNext
'        Loop
'    End With
'    Fill_lstvTableIzb = i
'    Set rst = Nothing
'End Function