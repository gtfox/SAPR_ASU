'

'Получаем Recordset по запросу
Public Function GetRecordSet_ADODB_Excel(XlsFileName As String, SQLQuery As String) As ADODB.Recordset
    Dim oConn As New ADODB.Connection
    Dim oRecordSet As New ADODB.Recordset
    oConn.Mode = adModeReadWrite
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & IIf(XlsFileName Like "*:*", XlsFileName, sSAPath & XlsFileName) & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    oRecordSet.CursorType = adOpenStatic
    oRecordSet.Open SQLQuery, oConn
    Set GetRecordSet_ADODB_Excel = oRecordSet
    oRecordSet.Close
    oConn.Close
    Set oRecordSet = Nothing
    Set oConn = Nothing
End Function

'Заполняет ComboBox запросами из БД
Public Sub Fill_ComboBox(XlsFileName As String, SQLQuery As String, StolbExcel As Double, cmbx As ComboBox)
    Dim oRecordSet As ADODB.Recordset
    Set oRecordSet = GetRecordSet_ADODB_Excel(XlsFileName, SQLQuery)
    cmbx.Clear
    cmbx.ColumnCount = 1
    With oRecordSet
    If .EOF Then Exit Sub
        Do Until .EOF
            cmbx.AddItem .Fields(StolbExcel - 1).Value
            .MoveNext
        Loop
    End With
    oRecordSet.Close
    Set oRecordSet = Nothing
End Sub

Private Sub Reset_FiltersCmbx(PoizvoditelSettings As classProizvoditelBD)
    Dim SQLQuery As String
    If cmbxProizvoditel.ListIndex = -1 Then Exit Sub
    bBlock = True
    SQLQuery = "SELECT Категории.КодКатегории, Категории.Категория " & _
                "FROM Категории;"
    Fill_ComboBox PoizvoditelSettings.FileName, SQLQuery, PoizvoditelSettings.StolbKategoriya, cmbxKategoriya
    SQLQuery = "SELECT Группы.КодГруппы, Группы.Группа " & _
                "FROM Группы;"
    Fill_ComboBox PoizvoditelSettings.FileName, SQLQuery, PoizvoditelSettings.StolbGruppa, cmbxGruppa
    SQLQuery = "SELECT Подгруппы.КодПодгруппы, Подгруппы.Подгруппа " & _
                "FROM Подгруппы;"
    Fill_ComboBox PoizvoditelSettings.FileName, SQLQuery, PoizvoditelSettings.StolbPodgruppa, cmbxPodgruppa
    bBlock = False
    lstvTablePrice.ListItems.Clear
    lblResult.Caption = "Найдено записей: 0"
End Sub

Private Sub Filter_CmbxChange(Ncmbx As Integer)
    Dim SQLQuery As String
    Dim fltrKategoriya As String
    Dim fltrGruppa As String
    Dim fltrPodgruppa As String
    Dim fltrMode As Integer
    Dim fltrWHERE As String
    Dim DBName As String

    If cmbxKategoriya.ListIndex = -1 Then
        fltrKategoriya = ""
    Else
        fltrKategoriya = "Прайс.КатегорииКод=" & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1)
    End If
    If cmbxGruppa.ListIndex = -1 Then
        fltrGruppa = ""
    Else
        fltrGruppa = "Прайс.ГруппыКод=" & cmbxGruppa.List(cmbxGruppa.ListIndex, 1)
    End If
    If cmbxPodgruppa.ListIndex = -1 Then
        fltrPodgruppa = ""
    Else
        fltrPodgruppa = "Прайс.ПодгруппыКод=" & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1)
    End If
    
    fltrMode = IIf(fltrKategoriya = "", 0, 4) + IIf(fltrGruppa = "", 0, 2) + IIf(fltrPodgruppa = "", 0, 1)
    
'-------------------ФИЛЬТРАЦИЯ БЕЗ ПРИОРИТЕТА (Нет иерархии: Категория || Группа || Подгруппа)------------------------------------------------
    '*    К   Гр  Пг
    '0    0   0   0
    '1    0   0   1
    '2    0   1   0
    '3    0   1   1
    '4    1   0   0
    '5    1   0   1
    '6    1   1   0
    '7    1   1   1
    
    Select Case fltrMode
        Case 0
            fltrWHERE = ""
        Case 1
            fltrWHERE = " WHERE " & fltrPodgruppa
        Case 2
            fltrWHERE = " WHERE " & fltrGruppa
        Case 3
            fltrWHERE = " WHERE " & fltrGruppa & " AND " & fltrPodgruppa
        Case 4
            fltrWHERE = " WHERE " & fltrKategoriya
        Case 5
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrPodgruppa
        Case 6
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrGruppa
        Case 7
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & " AND " & fltrPodgruppa
        Case Else
            fltrWHERE = ""
            fltrKategoriya = ""
            fltrGruppa = ""
            fltrPodgruppa = ""
    End Select
'-------------------ФИЛЬТРАЦИЯ БЕЗ ПРИОРИТЕТА (Нет иерархии: Категория || Группа || Подгруппа)------------------------------------------------

'-------------------ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
    Select Case Ncmbx
        Case 1
            fltrWHERE = " WHERE " & fltrKategoriya
            fltrGruppa = ""
            fltrPodgruppa = ""
            bBlock = True
            cmbxGruppa.Clear
            cmbxPodgruppa.Clear
            bBlock = False
        Case 2
            fltrWHERE = IIf(fltrKategoriya = "", " WHERE " & fltrGruppa, " WHERE " & fltrKategoriya & " AND " & fltrGruppa)
            fltrPodgruppa = ""
            bBlock = True
            cmbxPodgruppa.Clear
            bBlock = False
        Case 3
            'Работают варианты 1,3,5,7 из ФИЛЬТРАЦИЯ БЕЗ ПРИОРИТЕТА
        Case Else
            fltrWHERE = ""
            fltrKategoriya = ""
            fltrGruppa = ""
            fltrPodgruppa = ""
    End Select
'-------------------ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------


    SQLQuery = "SELECT Прайс.КодПозиции, Прайс.Артикул, Прайс.Название, Прайс.Цена, Прайс.КатегорииКод, Прайс.ГруппыКод, Прайс.ПодгруппыКод, Прайс.ПроизводительКод, Прайс.ЕдиницыКод, Единицы.Единица " & _
                "FROM Единицы INNER JOIN Прайс ON Единицы.КодЕдиницы = Прайс.ЕдиницыКод " & fltrWHERE & ";"
                
    DBName = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1)
    
    NameQueryDef = "FilterSQLQuery"

'lstvTablePrice.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable_(DBName, SQLQuery, NameQueryDef, lstvTablePrice)
'lstvTablePrice.Visible = True

    Fill_FiltersByResultSQLQuery DBName, fltrKategoriya, fltrGruppa, fltrPodgruppa

    ReSize

    'Find_ItemsByText
    
End Sub

Sub Find_ItemsByText()
    Dim DBName As String
    Dim SQLQuery As String
    Dim findMode As Integer
    Dim findWHERE As String
    Dim findArtikul As String
    Dim findNazvanie As String
    
    If cmbxProizvoditel.ListIndex = -1 Then Exit Sub
    
    DBName = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1)
    
    If txtArtikul.Value = "" Then
        findArtikul = ""
    Else
        findArtikul = "Прайс.Артикул like ""*" & txtArtikul.Value & "*"""
    End If
    
    If txtNazvanie1.Value = "" And txtNazvanie2.Value = "" And txtNazvanie3.Value = "" Then
        findNazvanie = ""
    Else
        findNazvanie = "Прайс.Название like ""*" & txtNazvanie1.Value & "*" & Replace(txtNazvanie2.Value, " ", "*") & "*" & txtNazvanie3.Value & "*"""
    End If
    
    findMode = IIf(findArtikul = "", 0, 2) + IIf(findNazvanie = "", 0, 1)

    '*   Арт Наз
    '0   0   0
    '1   0   1
    '2   1   0
    '3   1   1

    Select Case findMode
        Case 0
            findWHERE = ""
        Case 1
            findWHERE = " WHERE " & findNazvanie
        Case 2
            findWHERE = " WHERE " & findArtikul
        Case 3
            findWHERE = " WHERE " & findArtikul & " AND " & findNazvanie
        Case Else
            findWHERE = ""
    End Select

    If cmbxKategoriya.ListIndex = -1 And cmbxGruppa.ListIndex = -1 And cmbxPodgruppa.ListIndex = -1 Then
        NameQueryDef = "FilterSQLQuery"
        SQLQuery = "SELECT Прайс.КодПозиции, Прайс.Артикул, Прайс.Название, Прайс.Цена, Прайс.КатегорииКод, Прайс.ГруппыКод, Прайс.ПодгруппыКод, Прайс.ПроизводительКод, Прайс.ЕдиницыКод, Единицы.Единица " & _
                   "FROM Единицы INNER JOIN Прайс ON Единицы.КодЕдиницы = Прайс.ЕдиницыКод " & findWHERE & ";"
'lstvTablePrice.Visible = False
        lblResult.Caption = "Найдено записей: " & Fill_lstvTable_(DBName, SQLQuery, NameQueryDef, lstvTablePrice)
'lstvTablePrice.Visible = True
        Fill_FiltersByResultSQLQuery DBName, "", "", ""
    Else
        NameQueryDef = ""
        SQLQuery = "SELECT FilterSQLQuery.КодПозиции, FilterSQLQuery.Артикул, FilterSQLQuery.Название, FilterSQLQuery.Цена, FilterSQLQuery.КатегорииКод, FilterSQLQuery.ГруппыКод, FilterSQLQuery.ПодгруппыКод, FilterSQLQuery.ПроизводительКод, FilterSQLQuery.ЕдиницыКод, FilterSQLQuery.Единица " & _
                   "FROM Единицы INNER JOIN FilterSQLQuery ON Единицы.КодЕдиницы = FilterSQLQuery.ЕдиницыКод " & findWHERE & ";"
'lstvTablePrice.Visible = False
        lblResult.Caption = "Найдено записей: " & Fill_lstvTable_(DBName, SQLQuery, NameQueryDef, lstvTablePrice)
'lstvTablePrice.Visible = True
    End If

    ReSize
 
End Sub


'Заполняет lstvTable данными из БД в виде Excel через ADODB
 Function Fill_lstvTable_ADO(XlsFileName As String, SQLQuery As String, lstvTable As ListView, PoizvoditelSettings As classProizvoditelBD, Optional ByVal TableType As Integer = 0) As String
    'TableType=1 - Избранное
    'TableType=2 - Набор
    Dim oRecordSet As ADODB.Recordset
    Dim i As Double
    Dim j As Double
    Dim itmx As ListItem
    
    Set oRecordSet = GetRecordSet_ADODB_Excel(XlsFileName, SQLQuery)
    lstvTable.ListItems.Clear
    With oRecordSet
        If .RecordCount > 0 Then
            If .EOF Then .Close: Exit Function
            Do Until .EOF
                If i < SA_nRows Then
                    Set itmx = lstvTable.ListItems.Add(, , IIf(IsNull(.Fields(PoizvoditelSettings.StolbArtikul - 1).Value), "", .Fields(PoizvoditelSettings.StolbArtikul - 1).Value)) 'Артикул
                    itmx.SubItems(1) = IIf(IsNull(.Fields(PoizvoditelSettings.StolbNazvanie - 1).Value), "", .Fields(PoizvoditelSettings.StolbNazvanie - 1).Value) 'Название
                    itmx.SubItems(2) = IIf(IsNull(.Fields(PoizvoditelSettings.StolbCena - 1).Value), "", .Fields(PoizvoditelSettings.StolbCena - 1).Value) 'Цена
                    itmx.SubItems(3) = IIf(IsNull(.Fields(PoizvoditelSettings.StolbEd - 1).Value), "", .Fields(PoizvoditelSettings.StolbEd - 1).Value) 'Единица
                    If TableType = 1 Then
                        itmx.SubItems(4) = IIf(IsNull(.Fields(5 - 1).Value), "", .Fields(5 - 1).Value) 'Производитель
                    ElseIf TableType = 2 Then
                        itmx.SubItems(4) = IIf(IsNull(.Fields(5 - 1).Value), "", .Fields(5 - 1).Value) 'Производитель
                        itmx.SubItems(5) = IIf(IsNull(.Fields(6 - 1).Value), "", .Fields(6 - 1).Value) 'Количество
                    End If
            
                    'красим наборы
                    If TableType = 1 Then
                        If IIf(IsNull(.Fields(PoizvoditelSettings.StolbArtikul - 1).Value), "", .Fields(PoizvoditelSettings.StolbArtikul - 1).Value) Like "Набор_*" Then
                            itmx.ForeColor = NaboryColor
                           'itmx.Bold = True
                            For j = 1 To itmx.ListSubItems.Count
                               'itmx.ListSubItems(j).Bold = True
                                itmx.ListSubItems(j).ForeColor = NaboryColor
                            Next
                        End If
                    End If
                End If
                i = i + 1
                .MoveNext
            Loop
        End If
    End With
    Fill_lstvTable = IIf(TableType = 2, i, IIf(i <= SA_nRows, i, i & ".  Показано: " & SA_nRows))
    oRecordSet.Close
    Set oRecordSet = Nothing
End Function