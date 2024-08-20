

Dim NazvanieFSA As String
Dim NazvanieShkafa As String



Private Sub btnExportCx_Click()
    FindElementShemyToExcel
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnExportFSA_Click()
    
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnExportCxKJ_Click()
    FindKabeliShemyToExcel
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnExportPlan_Click()
    
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnExportVid_Click()
    
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub obTekListCx_Click()
    frameOutListCx.Visible = True
End Sub

Private Sub obTekListFSA_Click()
    frameOutListFSA.Visible = True
End Sub

Private Sub obTekListPlan_Click()
    frameOutListPlan.Visible = True
End Sub

Private Sub obTekListVid_Click()
    frameOutListVid.Visible = True
End Sub

Private Sub obVseCx_Click()
    frameOutListCx.Visible = False
End Sub

Private Sub obVseCxKJ_Click()
    frameOutListCxKJ.Visible = False
End Sub

Private Sub obVseFSA_Click()
    frameOutListFSA.Visible = False
End Sub

Private Sub obVsePlan_Click()
    frameOutListPlan.Visible = False
End Sub

Private Sub obVseVid_Click()
frameOutListVid.Visible = False
End Sub

Private Sub obVybCx_Click()
    frameOutListCx.Visible = False
End Sub

Private Sub obVybCxKJ_Click()
    frameOutListCxKJ.Visible = True
End Sub

Private Sub obVybFSA_Click()
    frameOutListFSA.Visible = False
End Sub

Private Sub UserForm_Initialize()
    
    Fill_cmbxNazvanieShkafa
    Fill_cmbxNazvanieFSA
    
'    cmbxNazvanieShkafa.style = fmStyleDropDownList
    cmbxNazvanieFSA.style = fmStyleDropDownList
'    cmbxNazvanieShkafaKJ.style = fmStyleDropDownList
    
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieShkafa", 0) Then
        NazvanieShkafa = ActivePage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
        cmbxNazvanieShkafa.text = NazvanieShkafa
        cmbxNazvanieShkafaKJ.text = NazvanieShkafa
    End If
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        NazvanieFSA = ActivePage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
        cmbxNazvanieFSA.text = NazvanieFSA
    End If

    With mpRazdel
        .Left = Me.Left
        .Top = Me.Top
        .Width = Me.Width
        .Height = Me.Height
        .Value = IIf(NazvanieFSA = "", 0, 1)
    End With

    If ActivePage.name Like cListNameCxema & "*" Then
        mpRazdel.Pages("pgCx").Enabled = True
        mpRazdel.Pages("pgFSA").Enabled = False
        mpRazdel.Pages("pgCxKJ").Enabled = False
        mpRazdel.Pages("pgPlan").Enabled = False
        mpRazdel.Pages("pgVid").Enabled = False
        If NazvanieShkafa <> "" Then
            obVybCx.Value = True
        End If
    ElseIf ActivePage.name Like cListNameFSA & "*" Then
        mpRazdel.Pages("pgCx").Enabled = False
        mpRazdel.Pages("pgFSA").Enabled = True
        mpRazdel.Pages("pgCxKJ").Enabled = False
        mpRazdel.Pages("pgPlan").Enabled = False
        mpRazdel.Pages("pgVid").Enabled = False
        If NazvanieFSA <> "" Then
            obVybFSA.Value = True
        End If
    ElseIf ActivePage.name Like cListNameKJ & "*" Then
        mpRazdel.Pages("pgCx").Enabled = False
        mpRazdel.Pages("pgFSA").Enabled = False
        mpRazdel.Pages("pgCxKJ").Enabled = True
        mpRazdel.Pages("pgPlan").Enabled = False
        mpRazdel.Pages("pgVid").Enabled = False
        obVybCxKJ.Value = True
    ElseIf ActivePage.name Like cListNameVID & "*" Then
        mpRazdel.Pages("pgCx").Enabled = False
        mpRazdel.Pages("pgFSA").Enabled = False
        mpRazdel.Pages("pgCxKJ").Enabled = False
        mpRazdel.Pages("pgPlan").Enabled = False
        mpRazdel.Pages("pgVid").Enabled = True
    ElseIf ActivePage.name Like cListNamePlan & "*" Then
        mpRazdel.Pages("pgCx").Enabled = False
        mpRazdel.Pages("pgFSA").Enabled = False
        mpRazdel.Pages("pgCxKJ").Enabled = False
        mpRazdel.Pages("pgPlan").Enabled = True
        mpRazdel.Pages("pgVid").Enabled = False
    Else
        mpRazdel.Pages("pgCx").Enabled = False
        mpRazdel.Pages("pgFSA").Enabled = False
        mpRazdel.Pages("pgCxKJ").Enabled = False
        mpRazdel.Pages("pgPlan").Enabled = False
        mpRazdel.Pages("pgVid").Enabled = False
    End If

    obNaListFSA.Value = True
    obNaListCx.Value = True
    obNaListCxKJ.Value = True
    obNaListPlan.Value = True
    obNaListVid.Value = True
    frameOutListPlan.Visible = True
    frameOutListVid.Visible = True
End Sub

Public Sub FindElementShemyToExcel()
'------------------------------------------------------------------------------------------------------------
' Macros        : FindElementShemyToExcel - Поиск элементов схемы и заполнение полей спецификации
'------------------------------------------------------------------------------------------------------------
    Dim clsStrokaSpecif As classStrokaSpecifikacii
    Dim colStrokaSpecif As Collection
    Dim colCxem As Collection
    Dim nCount As Double
    Dim strColKey As String
    Dim vsoPage As Visio.Page
    Dim vsoShapeOnPage As Visio.Shape
    Dim NazvanieShkafa As String   'Нумерация элементов идет в пределах одного шкафа
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim i As Integer
    Dim j As Integer
    Dim bExcelInit As Boolean
    Dim mNum() As String
    Dim Cxema As classCxema
    Dim xx As Integer
    '-------Вывод EXCEL---------
    Dim apx As Excel.Application
    Dim wb As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    Dim sPath, sFile As String
    Dim NameSheet As String
    Dim str As Integer
    Dim mStr() As String
    '-------Вывод на Лист-------
    Dim shpPerechenElementov As Visio.Shape
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    Dim ncell As Integer
    Dim NRow As Integer
    
    PageName = cListNameCxema  'Имена листов
    
    Set colCxem = New Collection
    Set Cxema = New classCxema
    Set Cxema.colListov = New Collection
    Set clsStrokaSpecif = New classStrokaSpecifikacii
    Set colStrokaSpecif = New Collection

    If obVseCx Then
        For j = 0 To cmbxNazvanieShkafa.ListCount - 1
            NazvanieShkafa = cmbxNazvanieShkafa.List(j)
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.name Like PageName & "*" Then
                    GoSub subShpOnPage
                End If
            Next
            If Not bExcelInit Then
                GoSub OutExcelInit
            Else
                GoSub OutExcelNext
            End If
            wb.Save
        Next
    ElseIf obVybCx Then
        NazvanieShkafa = cmbxNazvanieShkafa.text
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.name Like PageName & "*" Then
                GoSub subShpOnPage
            End If
        Next
        GoSub OutExcelInit
        wb.Save
    ElseIf obTekListCx Then
        Set vsoPage = ActivePage
        NazvanieShkafa = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
        If vsoPage.name Like PageName & "*" Then
            GoSub subShpOnPage
        End If
        If obVExcelCx Then
            GoSub OutExcelInit
            wb.Save
        Else 'obNaListCx
            GoSub OutList
        End If
    End If

Exit Sub

'-----------------------------------------------------------------------------------
subShpOnPage:
    For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
        If ShapeSAType(vsoShapeOnPage) > 1 Then   'Берем только шейпы САПР АСУ
            If vsoShapeOnPage.CellExists("User.Shkaf", 0) Then
                If vsoShapeOnPage.Cells("User.Shkaf").ResultStr(0) = NazvanieShkafa Then
                    UserType = ShapeSAType(vsoShapeOnPage)
                    Set clsStrokaSpecif = New classStrokaSpecifikacii
                    clsStrokaSpecif.SymName = vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
                    clsStrokaSpecif.SAType = vsoShapeOnPage.Cells("User.SAType").Result(0)
                    On Error Resume Next
                    clsStrokaSpecif.NazvanieDB = vsoShapeOnPage.Cells("Prop.NazvanieDB").ResultStr(0)
                    clsStrokaSpecif.ArtikulDB = vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
                    clsStrokaSpecif.ProizvoditelDB = vsoShapeOnPage.Cells("Prop.ProizvoditelDB").ResultStr(0)
                    clsStrokaSpecif.CenaDB = vsoShapeOnPage.Cells("Prop.CenaDB").ResultStr(0)
                    clsStrokaSpecif.EdDB = vsoShapeOnPage.Cells("Prop.EdDB").ResultStr(0)
                    clsStrokaSpecif.KolVo = 1
                    clsStrokaSpecif.PozOboznach = vsoShapeOnPage.Cells("Prop.Number").ResultStr(0)
                    clsStrokaSpecif.KodPoziciiDB = vsoShapeOnPage.Cells("User.KodPoziciiDB").Formula
                    strColKey = vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
                    err.Clear
                    On Error GoTo 0
                    Select Case UserType
                        Case typeCxemaCable 'Кабели на схеме электрической
                            clsStrokaSpecif.SymName = IIf(vsoShapeOnPage.Cells("Prop.BukvOboz").Result(0), vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), "")
                            clsStrokaSpecif.KolVo = vsoShapeOnPage.Cells("Prop.Dlina").Result(0)
                            strColKey = IIf(vsoShapeOnPage.Cells("Prop.BukvOboz").Result(0), vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0), "") & ";" & vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
                            On Error Resume Next
                            colStrokaSpecif.Add clsStrokaSpecif, strColKey
                            If colStrokaSpecif.Count = nCount Then 'Если кол-во не увеличелось, значит уже есть такой элемент - увеличиваем .KolVo в том, который есть
                                colStrokaSpecif(strColKey).KolVo = colStrokaSpecif(strColKey).KolVo + vsoShapeOnPage.Cells("Prop.Dlina").Result(0)
                                colStrokaSpecif(strColKey).PozOboznach = colStrokaSpecif(strColKey).PozOboznach & ";" & vsoShapeOnPage.Cells("Prop.Number").ResultStr(0)
                            Else
                                nCount = colStrokaSpecif.Count
                            End If
                        Case typeCxemaTerm 'Клеммы
                            clsStrokaSpecif.PozOboznach = vsoShapeOnPage.Cells("Prop.NumberKlemmnik").ResultStr(0)
                            On Error Resume Next
                            colStrokaSpecif.Add clsStrokaSpecif, strColKey
                            If colStrokaSpecif.Count = nCount Then 'Если кол-во не увеличелось, значит уже есть такой элемент - увеличиваем .KolVo в том, который есть
                                colStrokaSpecif(strColKey).KolVo = colStrokaSpecif(strColKey).KolVo + 1
                                mNum = Split(colStrokaSpecif(strColKey).PozOboznach, ";")
                                colStrokaSpecif(strColKey).PozOboznach = colStrokaSpecif(strColKey).PozOboznach & IIf(vsoShapeOnPage.Cells("Prop.NumberKlemmnik").ResultStr(0) = mNum(UBound(mNum)), "", ";" & vsoShapeOnPage.Cells("Prop.NumberKlemmnik").ResultStr(0))
                            Else
                                nCount = colStrokaSpecif.Count
                            End If
        
                        Case typeCxemaCoil, typeCxemaParent, typeCxemaElement, typePLCParent, typeCxemaSensor, typeCxemaActuator  ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
                            On Error Resume Next
                            colStrokaSpecif.Add clsStrokaSpecif, strColKey
                            If colStrokaSpecif.Count = nCount Then 'Если кол-во не увеличелось, значит уже есть такой элемент - увеличиваем .KolVo в том, который есть
                                colStrokaSpecif(strColKey).KolVo = colStrokaSpecif(strColKey).KolVo + 1
                                colStrokaSpecif(strColKey).PozOboznach = colStrokaSpecif(strColKey).PozOboznach & ";" & vsoShapeOnPage.Cells("Prop.Number").ResultStr(0)
                            Else
                                nCount = colStrokaSpecif.Count
                            End If
                    End Select
                End If
            End If
        End If
    Next
Return

SortReplace:
    'Сортировка номеров и замена последовательных позиционных обозначений
    For Each clsStrokaSpecif In colStrokaSpecif
        clsStrokaSpecif.PozOboznach = SortNumInString(clsStrokaSpecif.PozOboznach)
        clsStrokaSpecif.PozOboznach = ReplaceSequenceInString(clsStrokaSpecif.PozOboznach)
    Next
Return

OutExcelInit:
    
    Set apx = CreateObject("Excel.Application")
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set wb = apx.Workbooks.Open(sFile)

    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "СП_" & un & ".xls"
    apx.Visible = True
    bExcelInit = True

OutExcelNext:

    GoSub SortReplace
    
    str = colStrokaSpecif.Count
    If obTekListCx Then
        NameSheet = NazvanieShkafa & "_" & vsoPage.name & "_СП"
    Else
        NameSheet = NazvanieShkafa & "_СП"
    End If
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets(NameSheet).Delete
    apx.DisplayAlerts = True
    'Отключаем On Error Resume Next
    err.Clear
    On Error GoTo 0
    'добавляем новый
    apx.Sheets("СП").Visible = True
    apx.Sheets("СП").Copy After:=apx.Sheets(apx.Worksheets.Count)
    apx.Sheets("СП").Visible = False
    
    apx.Sheets("СП (2)").name = NameSheet
    
    lLastRow = apx.Sheets(NameSheet).Cells(apx.Rows.Count, 1).End(xlUp).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets(NameSheet).Activate
    apx.ActiveSheet.Rows("6:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A3:I5").ClearContents

    
    wb.Activate
    apx.ActiveSheet.Range("J1") = Format(Now(), "yyyy.mm.dd hh:mm:ss")
    apx.ActiveSheet.Range("D3:D65536").NumberFormat = "@"
    
    ReplaceNaborToSostav colStrokaSpecif 'Заменяем строки спецификации с наборами на составы наборов из БД
    
    'Access
'    For xx = 1 To str
'        If colStrokaSpecif(xx).ArtikulDB Like "Набор_*" Then
'            mStr = Split(colStrokaSpecif(xx).KodPoziciiDB, "/")
'            NElemNabora = AddSostavNaboraIzBD(colStrokaSpecif, colStrokaSpecif(xx).KolVo, mStr(0), xx)
'            str = str + NElemNabora - 1
'            colStrokaSpecif.Remove xx
'        End If
'    Next

    str = colStrokaSpecif.Count
    
    If str < 5 Then nstr = 5 Else nstr = str
    apx.ActiveSheet.Rows("5:" & nstr + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    For xx = 1 To str
        wb.Sheets(NameSheet).Cells(xx + 2, 1) = "=A" & xx + 1 & "+1" '1 Позиция
        wb.Sheets(NameSheet).Cells(xx + 2, 2) = colStrokaSpecif(xx).NazvanieDB '2 Наименование и техническая характеристика
        wb.Sheets(NameSheet).Cells(xx + 2, 3) = colStrokaSpecif(xx).ArtikulDB '3 Тип, марка, обозначение документа, опросного листа
        wb.Sheets(NameSheet).Cells(xx + 2, 4) = PozNameInString(colStrokaSpecif(xx).PozOboznach, colStrokaSpecif(xx).SymName) '4 Код оборудования, изделия, материала
        wb.Sheets(NameSheet).Cells(xx + 2, 5) = colStrokaSpecif(xx).ProizvoditelDB '5 Завод-изготовитель
        wb.Sheets(NameSheet).Cells(xx + 2, 6) = colStrokaSpecif(xx).EdDB '6 Единица измерения
        wb.Sheets(NameSheet).Cells(xx + 2, 7) = colStrokaSpecif(xx).KolVo '7 Количество
        'WB.Sheets(NameSheet).Cells(xx + 2, 8) = colStrokaSpecif(xx) '8 Масса единицы, кг
        'WB.Sheets(NameSheet).Cells(xx + 2, 9) = colStrokaSpecif(xx) '9 Примечание
        wb.Sheets(NameSheet).Cells(xx + 2, 11) = CSng(IIf(colStrokaSpecif(xx).CenaDB = "", 0, colStrokaSpecif(xx).CenaDB)) 'Цена
        wb.Sheets(NameSheet).Cells(xx + 2, 12) = "=K" & (xx + 2) & "*G" & (xx + 2)
        'wb.Sheets(NameSheet).Range("A" & (xx + 2)).Select 'для наглядности
    Next

    wb.Sheets(NameSheet).Range("A3") = 1
    wb.Sheets(NameSheet).Range("K2") = "Цена"
    wb.Sheets(NameSheet).Range("L2") = "Сумма"
    wb.Sheets(NameSheet).Range("K2:L2").HorizontalAlignment = xlRight
    wb.Sheets(NameSheet).Range("K2:L2").VerticalAlignment = xlCenter
    apx.ActiveSheet.Range("A3:I" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).WrapText = False
    apx.ActiveSheet.Range("A3:I" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид
    apx.ActiveSheet.Range("B3:B" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).HorizontalAlignment = xlLeft
    apx.ActiveSheet.Range("K3:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).NumberFormat = "#,##0"
    apx.ActiveSheet.Range("L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row + 1).FormulaLocal = "=СУММ(L3:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row & ")"
    For i = 7 To 12: apx.ActiveSheet.Range("K2:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row).Borders(i).Weight = 2: Next
    apx.ActiveSheet.Range("K2:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).Columns.AutoFit
    apx.ActiveSheet.Range("J1").Select
    
    
    Set clsStrokaSpecif = New classStrokaSpecifikacii
    Set colStrokaSpecif = New Collection
    
'    WB.Save
'    WB.Close SaveChanges:=True
'    apx.Quit
'    MsgBox "Спецификация экспортирована в файл SP_2_Visio.xls на лист " & NameSheet, vbInformation
Return

OutList:
    GoSub SortReplace
    ActivePage.Drop Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("ПЭ"), 0, ActivePage.PageSheet.Cells("PageHeight").Result(mm) - 10 / 25.4
    Set shpPerechenElementov = Application.ActiveWindow.Selection(1)

    ReplaceNaborToSostav colStrokaSpecif 'Заменяем строки спецификации с наборами на составы наборов из БД

    'Access
'    For NRow = 1 To str
'        If colStrokaSpecif(NRow).ArtikulDB Like "Набор_*" Then
'            mStr = Split(colStrokaSpecif(NRow).KodPoziciiDB, "/")
'            NElemNabora = AddSostavNaboraIzBD(colStrokaSpecif, colStrokaSpecif(NRow).KolVo, mStr(0), NRow)
'            If NRow < 5 Then nstr = 5 Else nstr = NRow
'            str = str + NElemNabora - 1
'            colStrokaSpecif.Remove NRow
'        End If
'    Next
    
    str = colStrokaSpecif.Count
    If str > 30 Then str = 30: MsgBox "Элементов на листе больше, чем строк в таблице(30): " & colStrokaSpecif.Count & vbNewLine & vbNewLine & "Используйте вывод в Excel для разбивки на несколько таблиц", vbExclamation, "САПР-АСУ: Перечень элементов"
    For NRow = 1 To str
        Set shpRow = shpPerechenElementov.Shapes("row" & NRow)
        shpRow.Shapes(NRow & ".1").text = IIf(PozNameInString(colStrokaSpecif(NRow).PozOboznach, colStrokaSpecif(NRow).SymName) = "", " ", PozNameInString(colStrokaSpecif(NRow).PozOboznach, colStrokaSpecif(NRow).SymName))
        shpRow.Shapes(NRow & ".2").text = IIf(colStrokaSpecif(NRow).NazvanieDB = "", " ", colStrokaSpecif(NRow).NazvanieDB)
        shpRow.Shapes(NRow & ".3").text = IIf(colStrokaSpecif(NRow).KolVo = "", " ", colStrokaSpecif(NRow).KolVo)
        shpRow.Shapes(NRow & ".4").text = IIf(colStrokaSpecif(NRow).ArtikulDB = "", " ", colStrokaSpecif(NRow).ArtikulDB)
        If shpRow.Shapes(NRow & ".3").text = " " Then
            shpRow.Shapes(NRow & ".2").CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "1" 'По центру
            shpRow.Shapes(NRow & ".2").CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = visItalic + visUnderLine 'Курсив+Подчеркивание
        End If
        shpRow.Shapes(NRow & ".2").CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
        shpRow.Shapes(NRow & ".4").CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
    Next
Return

End Sub

Public Sub FindKabeliShemyToExcel()
'------------------------------------------------------------------------------------------------------------
' Macros        : FindKabeliShemyToExcel - Поиск кабелей на схеме и заполнение полей кабельного журнала
'------------------------------------------------------------------------------------------------------------
    Dim clsStrokaKJ As classStrokaKabelnogoJurnala
    Dim colStrokaKJ As Collection
    Dim colCxem As Collection
    Dim nCount As Double
    Dim strColKey As String
    Dim vsoPage As Visio.Page
    Dim shpKabel As Visio.Shape
    Dim shpKabelPL As Visio.Shape
    Dim shpSensor As Visio.Shape
    Dim NazvanieShkafa As String   'Нумерация элементов идет в пределах одного шкафа
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim i As Integer
    Dim bExcelInit As Boolean
    Dim mNum() As String
    Dim Cxema As classCxema
    Dim xx As Integer
    '-------Вывод EXCEL---------
    Dim apx As Excel.Application
    Dim wb As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    Dim sPath, sFile As String
    Dim NameSheet As String
    Dim str As Integer
    Dim mStr() As String
    '-------Вывод на Лист-------
    Dim shpPerechenElementov As Visio.Shape
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    Dim ncell As Integer
    Dim NRow As Integer
    
    PageName = cListNameCxema  'Имена листов
    
    Set colCxem = New Collection
    Set Cxema = New classCxema
    Set Cxema.colListov = New Collection
'    Set clsStrokaKJ = New classStrokaKabelnogoJurnala
'    Set colStrokaKJ = New Collection

    If obVseCxKJ Then
        For i = 0 To cmbxNazvanieShkafaKJ.ListCount - 1
            NazvanieShkafa = cmbxNazvanieShkafaKJ.List(i)
            Set colStrokaKJ = New Collection
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.name Like PageName & "*" Then
                    GoSub FillcolStrokaKJ
                End If
            Next
'            If obNaListCxKJ Then
'                KJColToArray colStrokaKJ
'                fill_table_KJ
'            Else
                If Not bExcelInit Then
                    GoSub OutExcelInitKJ
                Else
                    GoSub OutExcelNextKJ
                End If
                wb.Save
'            End If
        Next
    ElseIf obVybCxKJ Then
        NazvanieShkafa = cmbxNazvanieShkafaKJ.text
        Set colStrokaKJ = New Collection
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.name Like PageName & "*" Then
                GoSub FillcolStrokaKJ
            End If
        Next
        If obNaListCxKJ Then
            KJColToArray colStrokaKJ
            fill_table_KJ
        Else
            GoSub OutExcelInitKJ
            wb.Save
        End If
    End If

Exit Sub

'-----------------------------------------------------------------------------------
FillcolStrokaKJ:
    For Each shpKabel In vsoPage.Shapes    'Перебираем все шейпы на листе
        If ShapeSATypeIs(shpKabel, typeCxemaCable) Then    'Берем только кабели схемы
            If GetNazvanie(shpKabel.Cells("User.LinkToBox").ResultStr(0), 2) = NazvanieShkafa Then
                Set clsStrokaKJ = New classStrokaKabelnogoJurnala
                Set shpSensor = FindSensorFromKabel(shpKabel)
                Set shpKabelPL = ShapeByGUID(shpKabel.Cells("Hyperlink.Kabel.ExtraInfo").ResultStr(0))
                clsStrokaKJ.Oboznach = IIf(shpKabel.Cells("Prop.BukvOboz").Result(0), shpKabel.Cells("Prop.SymName").ResultStr(0) & shpKabel.Cells("Prop.Number").Result(0), shpKabel.Cells("Prop.Number").Result(0))
                clsStrokaKJ.Nachalo = shpKabel.Cells("User.LinkToBox").ResultStr(0)
                clsStrokaKJ.Konec = Replace(shpKabel.Cells("User.LinkToSensor").ResultStr(0), vbLf, "")
                clsStrokaKJ.Trassa = GetTrassa(shpKabelPL)
                clsStrokaKJ.Marka = shpKabel.Cells("Prop.TipKab").ResultStr(0)
                clsStrokaKJ.Sechenie = shpKabel.Cells("Prop.WireCount").Result(0) & "x" & shpKabel.Cells("Prop.mm2").ResultStr(0)
                clsStrokaKJ.Dlina = Round(shpKabel.Cells("Prop.Dlina").Result(0), 1)
    
                colStrokaKJ.Add clsStrokaKJ, clsStrokaKJ.Oboznach
            End If
        End If
    Next
Return

OutExcelInitKJ:
    Set apx = CreateObject("Excel.Application")
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set wb = apx.Workbooks.Open(sFile)

    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "СП_" & un & ".xls"
    apx.Visible = True
    bExcelInit = True

OutExcelNextKJ:
    str = colStrokaKJ.Count
    NameSheet = NazvanieShkafa & "_КЖ"
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets(NameSheet).Delete
    apx.DisplayAlerts = True
    'Отключаем On Error Resume Next
    err.Clear
    On Error GoTo 0
    'добавляем новый
    apx.Sheets("КЖ").Visible = True
    apx.Sheets("КЖ").Copy After:=apx.Sheets(apx.Worksheets.Count)
    apx.Sheets("КЖ").Visible = False
    apx.Sheets("КЖ (2)").name = NameSheet
    
    lLastRow = apx.Sheets(NameSheet).Cells(apx.Rows.Count, 1).End(xlDown).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets(NameSheet).Activate
    apx.ActiveSheet.Rows("7:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A4:J6").ClearContents

    
    wb.Activate
    apx.ActiveSheet.Range("K3") = Format(Now(), "yyyy.mm.dd hh:mm:ss")
'    apx.ActiveSheet.Range("D3:D65536").NumberFormat = "@"

    
    If str < 5 Then nstr = 5 Else nstr = str
    apx.ActiveSheet.Rows("5:" & nstr + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    For xx = 1 To str
        wb.Sheets(NameSheet).Cells(xx + 3, 1) = colStrokaKJ(xx).Oboznach '1 Обозначение кабеля, провода
        wb.Sheets(NameSheet).Cells(xx + 3, 2) = " " & colStrokaKJ(xx).Nachalo '2 Трасса - Начало
        wb.Sheets(NameSheet).Cells(xx + 3, 3) = " " & colStrokaKJ(xx).Konec '3 Трасса - Конец
        wb.Sheets(NameSheet).Cells(xx + 3, 4) = colStrokaKJ(xx).Trassa '4 Участок трассы кабеля, провода
        wb.Sheets(NameSheet).Cells(xx + 3, 5) = colStrokaKJ(xx).Marka '5 Кабель, провод - по проекту - Марка
        wb.Sheets(NameSheet).Cells(xx + 3, 6) = colStrokaKJ(xx).Sechenie '6 Кабель, провод - по проекту - Кол., число и сечение жил
        wb.Sheets(NameSheet).Cells(xx + 3, 7) = CSng(colStrokaKJ(xx).Dlina) '7 Кабель, провод - по проекту - Длина, м.
        'wb.Sheets(NameSheet).Range("A" & (xx + 3)).Select 'для наглядности
    Next

'    WB.Sheets(NameSheet).Range("K2:L2").HorizontalAlignment = xlRight
'    WB.Sheets(NameSheet).Range("K2:L2").VerticalAlignment = xlCenter
    apx.ActiveSheet.Range("A4:I" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).WrapText = False
    apx.ActiveSheet.Range("A4:I" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид
'    apx.ActiveSheet.Range("B4:B" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).HorizontalAlignment = xlLeft
'    apx.ActiveSheet.Range("K4:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).NumberFormat = "#,##0"
'    For i = 7 To 12: apx.ActiveSheet.Range("K2:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row).Borders(i).Weight = 2: Next
'    apx.ActiveSheet.Range("K2:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).Columns.AutoFit
    apx.ActiveSheet.Range("A1").Select

    
    Set colStrokaKJ = New Collection
    
'    WB.Save
'    WB.Close SaveChanges:=True
'    apx.Quit
'    MsgBox "Спецификация экспортирована в файл SP_2_Visio.xls на лист " & NameSheet, vbInformation
Return

End Sub

Public Function GetTrassa(shpKabelPL As Visio.Shape) As String
'------------------------------------------------------------------------------------------------------------
' Function      : GetTrassa - Возвращает строку с трассой прокладки кабеля на плане
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    If Not shpKabelPL Is Nothing Then
        For i = 1 To shpKabelPL.Section(visSectionScratch).Count
            If cbShortName Then
                Select Case shpKabelPL.Cells("Scratch.A" & i).ResultStr(0)
                    Case "Лоток"
                        GetTrassa = GetTrassa & "Л. " & shpKabelPL.Cells("Scratch.B" & i).ResultStr(0) & " (" & shpKabelPL.Cells("Scratch.C" & i).ResultStr(0) & "м.), "
                    Case "Гофра"
                        GetTrassa = GetTrassa & "Г. " & shpKabelPL.Cells("Scratch.B" & i).ResultStr(0) & " (" & shpKabelPL.Cells("Scratch.C" & i).ResultStr(0) & "м.), "
                End Select
            Else
                GetTrassa = GetTrassa & shpKabelPL.Cells("Scratch.A" & i).ResultStr(0) & " " & shpKabelPL.Cells("Scratch.B" & i).ResultStr(0) & " (" & shpKabelPL.Cells("Scratch.C" & i).ResultStr(0) & "м.), "
            End If
        Next
        GetTrassa = Left(GetTrassa, Len(GetTrassa) - 2)
    Else
        GetTrassa = ""
    End If
End Function

Sub Fill_cmbxNazvanieShkafa()
    Dim colNameCxema As Collection
    Dim i As Integer
    
    Set colNameCxema = GetColNazvanieShkafa
    cmbxNazvanieShkafa.Clear
    cmbxNazvanieShkafaKJ.Clear
    For i = 1 To colNameCxema.Count
        cmbxNazvanieShkafa.AddItem colNameCxema.Item(i)
        cmbxNazvanieShkafaKJ.AddItem colNameCxema.Item(i)
    Next
    If colNameCxema.Count > 0 Then
        cmbxNazvanieShkafa.ListIndex = 0
        cmbxNazvanieShkafaKJ.ListIndex = 0
    End If
End Sub

Sub Fill_cmbxNazvanieFSA()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieFSA.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieFSA.AddItem mstrPropPageSheet(i)
    Next
    cmbxNazvanieFSA.text = ""
End Sub

Private Sub btnCloseCx_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnCloseFSA_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnCloseCxKJ_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnClosePlan_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnCloseVid_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub