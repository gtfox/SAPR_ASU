'------------------------------------------------------------------------------------------------------------
' Module        : DB - База данных прайс листов и избранного
' Author        : gtfox
' Date          : 2021.02.22
' Description   : База данных прайс листов, избранного и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

Public Const NaboryColor   As Long = &HBD0429

'Активация формы выбора элементов схемы из БД
Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
    Load frmDBPrice
    frmDBPrice.run vsoShape 'Передали его в форму
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

'Выполняем SQL запрос
Public Sub ExecuteSQL(DBName As String, SQLQuery As String, Optional QueryDefName As String = "")
    Dim dbs As DAO.Database
    Set dbs = GetDBEngine.OpenDatabase(ThisDocument.path & DBName)
    dbs.Execute SQLQuery
    Set dbs = Nothing
End Sub

'Заполняет ComboBox запросами из БД
Public Sub Fill_ComboBox(DBName As String, SQLQuery As String, cmbx As ComboBox)
    Dim rst As DAO.Recordset
    Dim i As Integer
    Set rst = GetRecordSet(DBName, SQLQuery)
    cmbx.Clear
    cmbx.ColumnCount = 1
    i = 0
    With rst
    If .EOF Then Exit Sub
        .MoveFirst
        Do Until .EOF
            cmbx.AddItem .Fields(1).Value
            cmbx.List(i, 1) = "" & .Fields(0).Value
            i = i + 1
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub

'Заполняет ComboBox Производители запросами из БД
Public Sub Fill_cmbxProizvoditel(DBName As String, SQLQuery As String, cmbx As ComboBox, Optional ByVal Price As Boolean = False)
    Dim rst As DAO.Recordset
    Dim i As Integer
    Set rst = GetRecordSet(DBName, SQLQuery)
    cmbx.Clear
    cmbx.ColumnCount = 1
    i = 0
    With rst
    If .EOF Then Exit Sub
        .MoveFirst
        .MoveNext 'Пропускаем первый элемент
        Do Until .EOF
            If "" & .Fields(0).Value = "" And Price Then

            Else
                cmbx.AddItem .Fields(1).Value
                cmbx.List(i, 1) = "" & .Fields(0).Value
                cmbx.List(i, 2) = .Fields(2).Value
                i = i + 1
            End If
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub

'Заполняет lstvTable запросами из БД
Public Function Fill_lstvTable(DBName As String, SQLQuery As String, QueryDefName As String, lstvTable As ListView, Optional ByVal TableType As Integer = 0) As Double
    'TableType=1 - Избранное
    'TableType=2 - Набор
    Dim i As Double
    Dim j As Double
    Dim itmx As ListItem
    Dim rst As DAO.Recordset
    Set rst = GetRecordSet(DBName, SQLQuery, QueryDefName)
    lstvTable.ListItems.Clear
    i = 0
    With rst
        If .EOF Then Exit Function
        .MoveFirst
        Do Until .EOF
            Set itmx = lstvTable.ListItems.Add(, """" & .Fields("КодПозиции").Value & "/" & .Fields("ПроизводительКод").Value & """", .Fields("Артикул").Value)
            itmx.SubItems(1) = .Fields("Название").Value
            itmx.SubItems(2) = .Fields("Цена").Value
            'itmx.SubItems(3) = .Fields("Единица").Value
            If TableType = 1 Or TableType = 2 Then itmx.SubItems(3) = .Fields("Производитель").Value
            If TableType = 2 Then itmx.SubItems(4) = .Fields("Количество").Value
            
            'красим наборы
            If TableType = 1 Then  'and .Fields("Артикул").Value like "Набор_*" then
                If .Fields("ПодгруппыКод").Value = 2 Then
                    itmx.ForeColor = NaboryColor
    '                    itmx.Bold = True
                    For j = 1 To itmx.ListSubItems.Count
    '                        itmx.ListSubItems(j).Bold = True
                        itmx.ListSubItems(j).ForeColor = NaboryColor
                    Next
                End If
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
    Fill_lstvTable = i
    Set rst = Nothing
End Function

'Заполняет lstvTableNabor запросами из БД
Public Function Fill_lstvTableNabor(DBName As String, IzbPozCod As String, lstvTable As ListView) As String
    Dim SQLQuery As String

    SQLQuery = "SELECT Наборы.КодПозиции, Наборы.ИзбрПозицииКод, Наборы.Артикул, Наборы.Название, Наборы.Цена, Наборы.Количество, Наборы.ПроизводительКод, Производители.Производитель " & _
                "FROM Производители INNER JOIN Наборы ON Производители.КодПроизводителя = Наборы.ПроизводительКод " & _
                "WHERE Наборы.ИзбрПозицииКод=" & IzbPozCod & ";"
    Fill_lstvTableNabor = Fill_lstvTable(DBName, SQLQuery, "", lstvTable, 2)
End Function

'Считаем цену набора
Public Function CalcCenaNabora(lstvTable As ListView) As Double
    Dim i As Integer
    Dim Sum As Double
    For i = 1 To lstvTable.ListItems.Count
        Sum = Sum + CDbl(lstvTable.ListItems(i).SubItems(2)) * CInt(lstvTable.ListItems(i).SubItems(4))
    Next
    CalcCenaNabora = Sum
End Function