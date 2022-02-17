

Dim NazvanieFSA As String
Dim NazvanieShemy As String

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

Private Sub obTekListCx_Click()
    frameOutListCx.Visible = True
End Sub

Private Sub obTekListFSA_Click()
    frameOutListFSA.Visible = True
End Sub

Private Sub obVseCx_Click()
    frameOutListCx.Visible = False
End Sub

Private Sub obVseFSA_Click()
    frameOutListFSA.Visible = False
End Sub

Private Sub obVybCx_Click()
    frameOutListCx.Visible = False
End Sub

Private Sub obVybFSA_Click()
    frameOutListFSA.Visible = False
End Sub

Private Sub UserForm_Initialize()
    
    Fill_cmbxNazvanieShemy
    Fill_cmbxNazvanieFSA
    
    cmbxNazvanieShemy.style = fmStyleDropDownList
    cmbxNazvanieFSA.style = fmStyleDropDownList
    
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieShemy", 0) Then
        NazvanieShemy = ActivePage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0)
        cmbxNazvanieShemy.Text = NazvanieShemy
    End If
    If ActivePage.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        NazvanieFSA = ActivePage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
        cmbxNazvanieFSA.Text = NazvanieFSA
    End If

    With mpRazdel
        .Left = Me.Left
        .Top = Me.Top
        .Width = Me.Width
        .Height = Me.Height
        .Value = IIf(NazvanieFSA = "", 0, 1)
    End With

    If NazvanieShemy <> "" Then
        obVybCx.Value = True
    End If
    If NazvanieFSA <> "" Then
        obVybFSA.Value = True
    End If
    
    obNaListFSA.Value = True
    obNaListCx.Value = True
End Sub

Public Sub FindElementShemyToExcel()
'------------------------------------------------------------------------------------------------------------
' Macros        : FindElementShemyToExcel - Поиск элементов схемы и заполнение полей спецификации

                '
                '
'------------------------------------------------------------------------------------------------------------
    Dim clsStrokaSpecif As classStrokaSpecifikacii
    Dim colStrokaSpecif As Collection
    Dim colCxem As Collection
    Dim nCount As Double
    Dim strColKey As String
    Dim vsoPage As Visio.Page
    Dim vsoShapeOnPage As Visio.Shape
    Dim NazvanieShemy As String   'Нумерация элементов идет в пределах одной схемы (одного номера схемы)
    Dim UserType As Integer     'Тип элемента схемы: клемма, провод, реле
    Dim PageName As String      'Имена листов где возможна нумерация
    Dim i As Integer
    Dim mNum() As String
    Dim Cxema As classCxema
    Dim xx As Integer
    '-------Вывод EXCEL---------
    Dim apx As Excel.Application
    Dim WB As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    Dim sPath, sFile As String
    Dim NameSheet As String
    Dim str As Integer
    Dim Mstr() As String
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

    For i = 1 To cmbxNazvanieShemy.ListCount
        NazvanieShemy = cmbxNazvanieShemy.List(i - 1)
        Cxema.NameCxema = NazvanieShemy
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.name Like PageName & "*" Then
                If NazvanieShemy = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy").ResultStr(0) Then
                    Cxema.colListov.Add vsoPage, vsoPage.name
                End If
            End If
        Next
        If Cxema.colListov.Count > 0 Then
            colCxem.Add Cxema, NazvanieShemy
        End If
        Set Cxema = New classCxema
        Set Cxema.colListov = New Collection
    Next
    
    i = 0
    If obVseCx Then
        For Each Cxema In colCxem
            NazvanieShemy = Cxema.NameCxema
            For Each vsoPage In Cxema.colListov
                GoSub ShpOnPage
            Next
            If i > 0 Then
                GoSub OutExcelNext
            Else
                GoSub OutExcel
            End If
            i = i + 1
        Next
        WB.Save
    ElseIf obVybCx Then
        NazvanieShemy = cmbxNazvanieShemy.Text
        For Each vsoPage In colCxem(NazvanieShemy).colListov
            GoSub ShpOnPage
        Next
        GoSub OutExcel
        WB.Save
    ElseIf obTekListCx Then
        Set vsoPage = ActivePage
        GoSub ShpOnPage
        If obVExcelCx Then
            GoSub OutExcel
            WB.Save
        Else 'obNaListCx
            GoSub OutList
        End If
    End If

Exit Sub

'-----------------------------------------------------------------------------------
ShpOnPage:
    For Each vsoShapeOnPage In vsoPage.Shapes    'Перебираем все шейпы на листе
        If ShapeSAType(vsoShapeOnPage) > 1 Then   'Берем только шейпы САПР АСУ
            UserType = ShapeSAType(vsoShapeOnPage)
            Set clsStrokaSpecif = New classStrokaSpecifikacii
            clsStrokaSpecif.SymName = vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0)
            clsStrokaSpecif.SAType = vsoShapeOnPage.Cells("User.SAType").Result(0)
            clsStrokaSpecif.NazvanieDB = vsoShapeOnPage.Cells("Prop.NazvanieDB").ResultStr(0)
            clsStrokaSpecif.ArtikulDB = vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
            clsStrokaSpecif.ProizvoditelDB = vsoShapeOnPage.Cells("Prop.ProizvoditelDB").ResultStr(0)
            clsStrokaSpecif.CenaDB = vsoShapeOnPage.Cells("Prop.CenaDB").ResultStr(0)
            clsStrokaSpecif.EdDB = vsoShapeOnPage.Cells("Prop.EdDB").ResultStr(0)
            clsStrokaSpecif.KolVo = 1
            clsStrokaSpecif.PozOboznach = vsoShapeOnPage.Cells("Prop.Number").ResultStr(0)
            clsStrokaSpecif.KodPoziciiDB = vsoShapeOnPage.Cells("User.KodPoziciiDB").Formula
            strColKey = vsoShapeOnPage.Cells("Prop.SymName").ResultStr(0) & ";" & vsoShapeOnPage.Cells("User.SAType").Result(0) & ";" & vsoShapeOnPage.Cells("Prop.ArtikulDB").ResultStr(0)
            
            Select Case UserType
                Case typeCableSH 'Кабели на схеме электрической
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
                Case typeTerm 'Клеммы
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

                Case typeCoil, typeParent, typeElement, typePLCParent, typePLCModParent, typeSensor, typeActuator ', typeElectroOneWire, typeElectroPlan, typeOPSPlan 'Остальные элементы
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
    Next
Return

SortReplace:
    'Сортировка номеров и замена последовательных позиционных обозначений
    For Each clsStrokaSpecif In colStrokaSpecif
        clsStrokaSpecif.PozOboznach = SortNumInString(clsStrokaSpecif.PozOboznach)
        clsStrokaSpecif.PozOboznach = ReplaceSequenceInString(clsStrokaSpecif.PozOboznach)
    Next
Return

OutExcel:
    
    Set apx = CreateObject("Excel.Application")
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "Ошибка"
        Exit Sub
    End If
    
    Set WB = apx.Workbooks.Open(sFile)

    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "СП_" & un & ".xls"
    apx.Visible = True

OutExcelNext:

    GoSub SortReplace
    
    str = colStrokaSpecif.Count
    If obTekListCx Then
        NameSheet = NazvanieShemy & "_" & vsoPage.name
    Else
        NameSheet = NazvanieShemy
    End If
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets(NameSheet).Delete
    apx.DisplayAlerts = True
    'добавляем новый
    apx.Sheets("СП").Copy After:=apx.Sheets(apx.Worksheets.Count)
    
    apx.Sheets("СП (2)").name = NameSheet
    
    lLastRow = apx.Sheets(NameSheet).Cells(apx.Rows.Count, 1).End(xlUp).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets(NameSheet).Activate
    apx.ActiveSheet.Rows("6:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A3:I5").ClearContents

    
    WB.Activate
    apx.ActiveSheet.Range("J1") = Format(Now(), "yyyy.mm.dd hh:mm:ss")
    apx.ActiveSheet.Range("D3:D65536").NumberFormat = "@"
    For xx = 1 To str
        If colStrokaSpecif(xx).ArtikulDB Like "Набор_*" Then
            Mstr = Split(colStrokaSpecif(xx).KodPoziciiDB, "/")
            NElemNabora = AddSostavNaboraIzBD(colStrokaSpecif, colStrokaSpecif(xx).KolVo, Mstr(0), xx)
            str = str + NElemNabora - 1
            colStrokaSpecif.Remove xx
        End If
    Next
    
    If str < 5 Then nstr = 5 Else nstr = str
    apx.ActiveSheet.Rows("5:" & nstr + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    For xx = 1 To str
        WB.Sheets(NameSheet).Cells(xx + 2, 1) = "=A" & xx + 1 & "+1" '1 Позиция
        WB.Sheets(NameSheet).Cells(xx + 2, 2) = colStrokaSpecif(xx).NazvanieDB '2 Наименование и техническая характеристика
        WB.Sheets(NameSheet).Cells(xx + 2, 3) = colStrokaSpecif(xx).ArtikulDB '3 Тип, марка, обозначение документа, опросного листа
        WB.Sheets(NameSheet).Cells(xx + 2, 4) = PozNameInString(colStrokaSpecif(xx).PozOboznach, colStrokaSpecif(xx).SymName) '4 Код оборудования, изделия, материала
        WB.Sheets(NameSheet).Cells(xx + 2, 5) = colStrokaSpecif(xx).ProizvoditelDB '5 Завод-изготовитель
        WB.Sheets(NameSheet).Cells(xx + 2, 6) = colStrokaSpecif(xx).EdDB '6 Единица измерения
        WB.Sheets(NameSheet).Cells(xx + 2, 7) = colStrokaSpecif(xx).KolVo '7 Количество
        'WB.Sheets(NameSheet).Cells(xx + 2, 8) = colStrokaSpecif(xx) '8 Масса единицы, кг
        'WB.Sheets(NameSheet).Cells(xx + 2, 9) = colStrokaSpecif(xx) '9 Примечание
        WB.Sheets(NameSheet).Cells(xx + 2, 11) = CSng(colStrokaSpecif(xx).CenaDB)  'Цена
        WB.Sheets(NameSheet).Cells(xx + 2, 12) = "=K" & xx + 2 & "*G" & xx + 2
        'wb.Sheets(NameSheet).Range("A" & (xx + 2)).Select 'для наглядности
    Next

    WB.Sheets(NameSheet).Range("A3") = 1
    WB.Sheets(NameSheet).Range("K2") = "Цена"
    WB.Sheets(NameSheet).Range("L2") = "Сумма"
    WB.Sheets(NameSheet).Range("K2:L2").HorizontalAlignment = xlRight
    WB.Sheets(NameSheet).Range("K2:L2").VerticalAlignment = xlCenter
    apx.ActiveSheet.Range("A3:I" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).WrapText = False
    apx.ActiveSheet.Range("A3:I" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид
    apx.ActiveSheet.Range("B3:B" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).HorizontalAlignment = xlLeft
    apx.ActiveSheet.Range("K3:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).NumberFormat = "#,##0"
    apx.ActiveSheet.Range("L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row + 1).FormulaLocal = "=СУММ(L3:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row & ")"
    For i = 7 To 12: Range("K2:L" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlUp).Row).Borders(i).Weight = 2: Next
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
    str = colStrokaSpecif.Count
    For NRow = 1 To str
        If colStrokaSpecif(NRow).ArtikulDB Like "Набор_*" Then
            Mstr = Split(colStrokaSpecif(NRow).KodPoziciiDB, "/")
            NElemNabora = AddSostavNaboraIzBD(colStrokaSpecif, colStrokaSpecif(NRow).KolVo, Mstr(0), NRow)
            If NRow < 5 Then nstr = 5 Else nstr = NRow
            str = str + NElemNabora - 1
            colStrokaSpecif.Remove NRow
        End If
    Next
    str = colStrokaSpecif.Count
    If str > 30 Then str = 30: MsgBox "Элементов на листе больше, чем строк в таблице(30): " & colStrokaSpecif.Count & vbNewLine & vbNewLine & "Используйте вывод в Excel для разбивки на несколько таблиц", vbExclamation, "Перечень элементов"
    For NRow = 1 To str
        Set shpRow = shpPerechenElementov.Shapes("row" & NRow)
        shpRow.Shapes(NRow & ".1").Text = PozNameInString(colStrokaSpecif(NRow).PozOboznach, colStrokaSpecif(NRow).SymName)
        shpRow.Shapes(NRow & ".2").Text = colStrokaSpecif(NRow).NazvanieDB
        shpRow.Shapes(NRow & ".3").Text = colStrokaSpecif(NRow).KolVo
        shpRow.Shapes(NRow & ".4").Text = colStrokaSpecif(NRow).ArtikulDB
        If shpRow.Shapes(NRow & ".3").Text = " " Then
            shpRow.Shapes(NRow & ".2").CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "1" 'По центру
            shpRow.Shapes(NRow & ".2").CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = visItalic + visUnderLine 'Курсив+Подчеркивание
        End If
        shpRow.Shapes(NRow & ".2").CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
        shpRow.Shapes(NRow & ".4").CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
    Next
Return

End Sub


Sub Fill_cmbxNazvanieShemy()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieShemy.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieShemy.AddItem mstrPropPageSheet(i)
    Next
    cmbxNazvanieShemy.Text = ""
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
    cmbxNazvanieFSA.Text = ""
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
