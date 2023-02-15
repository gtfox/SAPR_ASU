'------------------------------------------------------------------------------------------------------------
' Module        : DB_Access - База данных прайс листов и избранного на основе Access
' Author        : gtfox
' Date          : 2021.02.22
' Description   : База данных прайс листов, избранного и их обеспечение на основе Access
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit

Public Const frmMinWdth As Integer = 417 'Минимальна ширина формы
Public Const DBNameIzbrannoeAccess As String = "SAPR_ASU_Izbrannoe.accdb" 'Имя файла избронного

Public Const NaboryColor   As Long = &HBD0429 'синий

#If VBA7 Then
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

'Активация формы выбора элементов схемы из БД. Расположено в модуле DB_Excel
'Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
''    Load frmDBPriceAccess
''    frmDBPriceAccess.run vsoShape 'Передали его в форму
'    Load frmDBPriceExcel
'    frmDBPriceExcel.run vsoShape 'Передали его в форму
'End Sub

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
 Function Fill_lstvTable_(DBName As String, SQLQuery As String, QueryDefName As String, lstvTable As ListView, Optional ByVal TableType As Integer = 0) As Double
    'TableType=1 - Избранное
    'TableType=2 - Набор
    Dim i As Double
    Dim iold As Double
    Dim j As Double
    Dim itmx As ListItem
    Dim rst As DAO.Recordset
    Dim RecordCount As Double

    Set rst = GetRecordSet(DBName, SQLQuery, QueryDefName)
    lstvTable.ListItems.Clear
    If rst.RecordCount > 0 Then
        rst.MoveLast
        RecordCount = rst.RecordCount
    '    frmDBPriceAccess.lblResult.Caption = "Найдено записей: " & RecordCount
    '    frmDBPriceAccess.ProgressBar.Visible = True
    '    frmDBPriceAccess.ProgressBar.Max = RecordCount
'        lstvTable.ListItems.Clear
        i = 0
        iold = 1000
        With rst
            If .EOF Then rst.Close: Exit Function
            .MoveFirst
            Do Until .EOF
                Set itmx = lstvTable.ListItems.Add(, """" & .Fields("КодПозиции").Value & "/" & .Fields("ПроизводительКод").Value & "/" & .Fields("ЕдиницыКод").Value & """", .Fields("Артикул").Value)
                itmx.SubItems(1) = .Fields("Название").Value
                itmx.SubItems(2) = .Fields("Цена").Value
                itmx.SubItems(3) = .Fields("Единица").Value
                If TableType = 1 Then
                    itmx.SubItems(4) = .Fields("Производитель").Value
                    itmx.SubItems(5) = "    "
                ElseIf TableType = 2 Then
                    itmx.SubItems(4) = .Fields("Производитель").Value
                    itmx.SubItems(5) = .Fields("Количество").Value
                    itmx.SubItems(6) = "    "
                End If
    
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
    '            i = i + 1
    '            If iold < i Then
    '                iold = iold + 1000
    '                frmDBPriceAccess.ProgressBar.Value = i
    '            End If
                .MoveNext
            Loop
        End With
        Fill_lstvTable_ = RecordCount 'i
    '    frmDBPriceAccess.ProgressBar.Visible = False
    End If
    Set rst = Nothing

End Function

'Заполняет lstvTableNabor запросами из БД
Public Function Fill_lstvTableNabor(DBName As String, IzbPozCod As String, lstvTable As ListView) As String
    Dim SQLQuery As String

    SQLQuery = "SELECT Наборы.КодПозиции, Наборы.ИзбрПозицииКод, Наборы.Артикул, Наборы.Название, Наборы.Цена, Наборы.Количество, Наборы.ПроизводительКод, Производители.Производитель, Наборы.ЕдиницыКод, Единицы.Единица " & _
                "FROM Единицы INNER JOIN (Производители INNER JOIN Наборы ON Производители.КодПроизводителя = Наборы.ПроизводительКод) ON Единицы.КодЕдиницы = Наборы.ЕдиницыКод " & _
                "WHERE Наборы.ИзбрПозицииКод=" & IzbPozCod & ";"
    Fill_lstvTableNabor = Fill_lstvTable_(DBName, SQLQuery, "", lstvTable, 2)
End Function

'Считаем цену набора
Public Function CalcCenaNabora(lstvTable As ListView) As Double
    Dim i As Integer
    Dim Sum As Double
    For i = 1 To lstvTable.ListItems.Count
        Sum = Sum + CDbl(IIf(lstvTable.ListItems(i).SubItems(2) = "", 0, lstvTable.ListItems(i).SubItems(2))) * CInt(IIf(lstvTable.ListItems(i).SubItems(5) = "", 0, lstvTable.ListItems(i).SubItems(5)))
    Next
    CalcCenaNabora = Sum
End Function

