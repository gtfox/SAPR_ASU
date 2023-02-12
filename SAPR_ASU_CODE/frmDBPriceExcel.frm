'------------------------------------------------------------------------------------------------------------
' Module        : frmDBPriceExcel - Форма поиска и задания данных для элемента схемы из Баз Данных оборудования расположенных в отдельных файлах Excel. Каждый файл - отдельный производитель
' Author        : gtfox
' Date          : 2023.01.30
' Description   : Выбор данных из Excel прайс листа, фильтрация по категориям и полнотекстовый поиск
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#End If
Private Const LVM_FIRST As Long = &H1000   ' 4096
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)   ' 4126
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public glShape As Visio.Shape 'шейп из модуля DB
Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи
Dim mstrShpData(5) As String
Public SA_nRows As Double
Public bBlock As Boolean


'Private Sub txtNazvanie2_Change()
'    Find_ItemsByText
'End Sub

Private Sub UserForm_Initialize() ' инициализация формы
    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи
    
    lstvTablePrice.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvTablePrice.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTablePrice.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTablePrice.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
    lstvTablePrice.ColumnHeaders.Add , , "Ед." ' SubItems(3)
'    lstvTablePrice.ColumnHeaders.Add , , "    " ' SubItems(4)
    
    cmbxMagazin.Clear
    cmbxMagazin.AddItem "ЭТМ"
    cmbxMagazin.AddItem "АВС"
    cmbxMagazin.ListIndex = 0
    
    cmbxProizvoditel.style = fmStyleDropDownList
    cmbxKategoriya.style = fmStyleDropDownList
    cmbxGruppa.style = fmStyleDropDownList
    cmbxPodgruppa.style = fmStyleDropDownList
    cmbxMagazin.style = fmStyleDropDownList

    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
    
    tbtnFiltr.Caption = ChrW(9650)
    tbtnBD = True
    SA_nRows = Visio.ActiveDocument.DocumentSheet.Cells("User.SA_nRows").Result(0)

    InitExcelDB
    FillExcel_mProizvoditel
    FillExcel_cmbxProizvoditel cmbxProizvoditel, True

    Load frmDBIzbrannoeExcel

End Sub

Sub run(vsoShape As Visio.Shape) 'Приняли шейп из модуля DB
    Dim ArtikulDB As String

    Set glShape = vsoShape 'И определили его как глолбальный в форме frmDBPriceExcel
    ArtikulDB = glShape.Cells("Prop.ArtikulDB").ResultStr(0)
    If ArtikulDB <> "" Then
        bBlock = True
        For i = 0 To cmbxProizvoditel.ListCount - 1
            If cmbxProizvoditel.List(i, 0) = glShape.Cells("Prop.ProizvoditelDB").ResultStr(0) Then cmbxProizvoditel.ListIndex = i
        Next
        Fill_cmbxProizvoditel
        If cmbxProizvoditel.ListIndex <> -1 And Not (ArtikulDB Like "Набор_*") Then
            txtArtikul.Value = ArtikulDB
            tbtnFiltr.Value = False
            Find_ItemsByText
            txtArtikul.Value = ""
            bBlock = False
            frmDBPriceExcel.Show
        Else
            bBlock = False
            frmDBIzbrannoeExcel.bBlock = True
            frmDBIzbrannoeExcel.txtArtikul.Value = ArtikulDB
            frmDBIzbrannoeExcel.tbtnFiltr.Value = False
            frmDBIzbrannoeExcel.Find_ItemsByText
            frmDBIzbrannoeExcel.txtArtikul.Value = ""
            frmDBIzbrannoeExcel.bBlock = False
            frmDBIzbrannoeExcel.Show
        End If
    Else
        frmDBPriceExcel.Show
    End If
End Sub

Sub Fill_cmbxProizvoditel()
    Dim UserRange As Excel.Range
    For i = 0 To UBound(mProizvoditel)
        If cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) = mProizvoditel(i).Proizvoditel Then
            If mProizvoditel(i).FileName <> "" Then 'пустое имя файла - пропускаем
                If Not wbExcelPrice Is Nothing Then wbExcelPrice.Close SaveChanges:=False
                If mProizvoditel(i).FileName Like ":" Then
                    Set wbExcelPrice = oExcelApp.Workbooks.Open(mProizvoditel(i).FileName) 'абсолютный адрес
                Else
                    Set wbExcelPrice = oExcelApp.Workbooks.Open(sSAPath & mProizvoditel(i).FileName) 'относительный
                End If
                Set wshPrice = wbExcelPrice.Worksheets(mProizvoditel(i).NameListExcel)
                Set CurentPrice = mProizvoditel(i)
                MaxColumn = WorksheetFunction.Max(CurentPrice.Artikul, CurentPrice.Nazvanie, CurentPrice.Cena, CurentPrice.Ed, CurentPrice.Kategoriya, CurentPrice.Gruppa, CurentPrice.Podgruppa)
                MinColumn = WorksheetFunction.Min(CurentPrice.Artikul, CurentPrice.Nazvanie, CurentPrice.Cena, CurentPrice.Ed, CurentPrice.Kategoriya, CurentPrice.Gruppa, CurentPrice.Podgruppa)
                Set RangePrice = wshPrice.Range(wshPrice.Cells(CurentPrice.FirstRow, MaxColumn), wshPrice.Cells(CurentPrice.LastRow, MinColumn))
                ClearFilter wshPrice
                Exit For
            End If
        End If
    Next
End Sub

Private Sub cmbxProizvoditel_Change()
    If Not bBlock Then
        Fill_cmbxProizvoditel
        ClearFilter wshPrice
        UpdateAllCmbxFilters
        lstvTablePrice.ListItems.Clear
'        Find_ItemsByText
    End If
End Sub

Private Sub Filter_CmbxChange(Ncmbx As Integer)
    Dim fltrMode As Integer
    
    '-------------------ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
    Select Case Ncmbx
        Case 1
            RangePrice.AutoFilter Field:=CurentPrice.Kategoriya, Criteria1:=cmbxKategoriya.List(cmbxKategoriya.ListIndex, 0) 'Категория
            RangePrice.AutoFilter Field:=CurentPrice.Gruppa 'Группа
            RangePrice.AutoFilter Field:=CurentPrice.Podgruppa 'Подгруппа
            UpdateCmbxFiltersPrice cmbxGruppa, CurentPrice.Gruppa
            UpdateCmbxFiltersPrice cmbxPodgruppa, CurentPrice.Podgruppa
        Case 2
            RangePrice.AutoFilter Field:=CurentPrice.Gruppa, Criteria1:=cmbxGruppa.List(cmbxGruppa.ListIndex, 0) 'Группа
            If cmbxKategoriya.ListIndex = -1 Then
                RangePrice.AutoFilter Field:=CurentPrice.Kategoriya
                UpdateCmbxFiltersPrice cmbxKategoriya, CurentPrice.Kategoriya
            Else
                RangePrice.AutoFilter Field:=CurentPrice.Kategoriya, Criteria1:=cmbxKategoriya.List(cmbxKategoriya.ListIndex, 0) 'Категория
            End If
            UpdateCmbxFiltersPrice cmbxPodgruppa, CurentPrice.Podgruppa
        Case 3
            '-------------------ФИЛЬТРАЦИЯ Подгруппы при разных (Категория || Группа)------------------------------------------------
            '*    К   Гр
            '0    0   0
            '1    0   1
            '2    1   0
            '3    1   1
            
            fltrMode = IIf(cmbxKategoriya.ListIndex = -1, 0, 2) + IIf(cmbxGruppa.ListIndex = -1, 0, 1)
            RangePrice.AutoFilter Field:=CurentPrice.Podgruppa, Criteria1:=cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 0) 'Подгруппа
            Select Case fltrMode
                Case 0
                    RangePrice.AutoFilter Field:=CurentPrice.Kategoriya 'Категория
                    RangePrice.AutoFilter Field:=CurentPrice.Gruppa 'Группа
                    UpdateCmbxFiltersPrice cmbxKategoriya, CurentPrice.Kategoriya
                    UpdateCmbxFiltersPrice cmbxGruppa, CurentPrice.Gruppa
                Case 1
                    RangePrice.AutoFilter Field:=CurentPrice.Kategoriya 'Категория
                    RangePrice.AutoFilter Field:=CurentPrice.Gruppa, Criteria1:=cmbxGruppa.List(cmbxGruppa.ListIndex, 0) 'Группа
                    UpdateCmbxFiltersPrice cmbxKategoriya, CurentPrice.Kategoriya
                Case 2
                    RangePrice.AutoFilter Field:=CurentPrice.Kategoriya, Criteria1:=cmbxKategoriya.List(cmbxKategoriya.ListIndex, 0) 'Категория
                    RangePrice.AutoFilter Field:=CurentPrice.Gruppa 'Группа
                    UpdateCmbxFiltersPrice cmbxGruppa, CurentPrice.Gruppa
                Case 3
                    RangePrice.AutoFilter Field:=CurentPrice.Kategoriya, Criteria1:=cmbxKategoriya.List(cmbxKategoriya.ListIndex, 0) 'Категория
                    RangePrice.AutoFilter Field:=CurentPrice.Gruppa, Criteria1:=cmbxGruppa.List(cmbxGruppa.ListIndex, 0) 'Группа
                Case Else
            End Select
            '-------------------/ФИЛЬТРАЦИЯ Подгруппы при разных (Категория || Группа)------------------------------------------------
        Case Else
            RangePrice.AutoFilter Field:=CurentPrice.Kategoriya 'Категория
            RangePrice.AutoFilter Field:=CurentPrice.Gruppa 'Группа
            RangePrice.AutoFilter Field:=CurentPrice.Podgruppa 'Подгруппа
            UpdateAllCmbxFilters
    End Select
    '-------------------/ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
   
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(wshPrice, lstvTablePrice)
    ReSize

End Sub

Private Sub UpdateCmbxFiltersPrice(cmbxComboBox As ComboBox, nColumn As Long)
    'nColumn = CurentPrice.Kategoriya - Категория
    'nColumn = CurentPrice.Gruppa - Группа
    'nColumn = CurentPrice.Podgruppa - Подгруппа
    Dim UserRange As Excel.Range
    Dim lLastRow As Long
    Dim i As Integer
    Dim mFilter() As String
    
    bBlock = True
    wshTemp.Cells.ClearContents
    lLastRow = wshPrice.Cells(wshPrice.Rows.Count, 1).End(xlUp).Row
    If lLastRow > 1 Then
        wshPrice.Range(wshPrice.Cells(CurentPrice.FirstRow, nColumn), wshPrice.Cells(lLastRow, nColumn)).Copy wshTemp.Cells(1, 1)
        Set UserRange = wshTemp.Range(wshTemp.Cells(1, 1), wshTemp.Cells(lLastRow - 1, 1))
        UserRange.RemoveDuplicates Columns:=1, Header:=xlNo
        lLastRow = wshTemp.Cells(wshTemp.Rows.Count, 1).End(xlUp).Row
        If lLastRow > 0 Then
            cmbxComboBox.Clear
            For i = 1 To lLastRow
                cmbxComboBox.AddItem wshTemp.Cells(i, 1)
            Next
        End If
    Else
        cmbxComboBox.Clear
    End If
    bBlock = False
End Sub


'Полнотекстовый поиск
Sub Find_ItemsByText()
    Dim RangeToFilter As Excel.Range
    
    Set RangeToFilter = RangePrice
    
    If txtArtikul.Value = "" Then
        RangeToFilter.AutoFilter Field:=CurentPrice.Artikul
    Else
        RangeToFilter.AutoFilter Field:=CurentPrice.Artikul, Criteria1:="=*" & txtArtikul.Value & "*"
    End If
    
    If txtNazvanie2.Value = "" Then
        RangeToFilter.AutoFilter Field:=CurentPrice.Nazvanie
    Else
        RangeToFilter.AutoFilter Field:=CurentPrice.Nazvanie, Criteria1:="=*" & Replace(txtNazvanie2.Value, " ", "*") & "*"
    End If
    
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(wshPrice, lstvTablePrice)
    
    UpdateAllCmbxFilters
    
    ReSize
    
End Sub

'Заполняет lstvTable данными из БД
Public Function Fill_lstvTable(wSheets As Excel.Worksheet, lstvTable As ListView) As String
    Dim RangeToFill As Excel.Range
    Dim RangeResult As Excel.Range
    Dim RangeRow As Excel.Range
    Dim i As Double
    Dim itmx As ListItem
    Set RangeToFill = wSheets.AutoFilter.Range
    'исключаем из диапазона автофильтра первую строку (Offset),
    'берем видимые строки(SpecialCells(xlCellTypeVisible).EntireRow)
    'и в цикле перебираем эти сроки
    On Error GoTo err1
    Set RangeResult = RangeToFill.Offset(1, 0).ReSize(RangeToFill.Rows.Count - 1, RangeToFill.Columns.Count).SpecialCells(xlCellTypeVisible).EntireRow
    lstvTable.ListItems.Clear
    lstvTablePrice.Visible = False
    For Each RangeRow In RangeResult.Rows
        Set itmx = lstvTable.ListItems.Add(, , RangeRow.Cells(1, CurentPrice.Artikul))  'Артикул
        itmx.SubItems(1) = RangeRow.Cells(1, CurentPrice.Nazvanie) 'Название
        itmx.SubItems(2) = RangeRow.Cells(1, CurentPrice.Cena) 'Цена
        itmx.SubItems(3) = RangeRow.Cells(1, CurentPrice.Ed) 'Единица
'        itmx.SubItems(4) = "              "
        i = i + 1
        If i = SA_nRows Then Exit For
    Next
    lstvTablePrice.Visible = True
    Fill_lstvTable = RangeResult.Rows.Count & ".  Показано: " & i
    Exit Function
err1:
    lstvTable.ListItems.Clear
End Function

'Очистка фильтров
Sub ClearFilter(wshWorkSheet As Excel.Worksheet)
    wshWorkSheet.Cells(CurentPrice.FirstRow, CurentPrice.Artikul).AutoFilter
    wshWorkSheet.Cells(CurentPrice.FirstRow, CurentPrice.Artikul).AutoFilter Field:=CurentPrice.Artikul
End Sub

'Заполнение всех фильтров
Sub UpdateAllCmbxFilters()
    UpdateCmbxFiltersPrice cmbxKategoriya, CurentPrice.Kategoriya
    UpdateCmbxFiltersPrice cmbxGruppa, CurentPrice.Gruppa
    UpdateCmbxFiltersPrice cmbxPodgruppa, CurentPrice.Podgruppa
End Sub

'Добавить в избранное
Private Sub btnFavAdd_Click()
    If mstrShpData(1) <> "" Then
        If Not bBlock Then
            bBlock = True
            tbtnFav = False
            bBlock = False
            Me.Hide
        End If
        Load frmDBAddToIzbrannoeExcel
        frmDBAddToIzbrannoeExcel.run mstrShpData(0), Replace(mstrShpData(1), """", """"""), mstrShpData(2), mstrShpData(4), mstrShpData(3)
    End If
End Sub

'Добавить в набор
Private Sub btnNabAdd_Click()
    If mstrShpData(1) <> "" Then
        Me.Hide
        Load frmDBAddToNaborExcel
        frmDBAddToNaborExcel.run mstrShpData(0), Replace(mstrShpData(1), """", """"""), mstrShpData(2), mstrShpData(4), mstrShpData(3)
    End If
End Sub

Private Sub lstvTablePrice_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Если в таблице ткнуть на строку с номером больше 30000 то сюда попадет первая строка!!!
    mstrShpData(0) = Item             'Артикул
    mstrShpData(1) = Item.SubItems(1) 'Название
    mstrShpData(2) = Item.SubItems(2) 'Цена
    mstrShpData(3) = Item.SubItems(3) 'Единица
    mstrShpData(4) = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) 'Производитель 'cmbxProizvoditel.Value
End Sub

Private Sub lstvTablePrice_DblClick()
    Dim vsoShape As Visio.Shape
    
    With glShape
        .Cells("User.KodProizvoditelyaDB").Formula = """"""
        .Cells("User.KodPoziciiDB").Formula = """"""
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(1), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(0) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrShpData(2) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrShpData(3) & """"
    End With
    
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> glShape And ShapeSATypeIs(vsoShape, ShapeSAType(glShape)) Then
                With vsoShape
                    .Cells("User.KodProizvoditelyaDB").Formula = """"""
                    .Cells("User.KodPoziciiDB").Formula = """"""
                    .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(1), """", """""") & """"
                    .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(0) & """"
                    .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
                    .Cells("Prop.CenaDB").Formula = """" & mstrShpData(2) & """"
                    .Cells("Prop.EdDB").Formula = """" & mstrShpData(3) & """"
                End With
            End If
        Next
    End If
    
    btnClose_Click

End Sub

Private Sub ReSize() ' изменение формы. Зависит от длины в lstvTablePrice
    Dim TablePriceWidth As Single
    
    lstvTablePrice.Width = frmMinWdth

    lblHeaders_Click
    
    If lstvTablePrice.ListItems.Count < 1 Then Exit Sub

    If lstvTablePrice.ListItems(1).Width > frmMinWdth Then
        TablePriceWidth = lstvTablePrice.ListItems(1).Width
    Else
        TablePriceWidth = frmMinWdth
    End If
    
    lstvTablePrice.Width = TablePriceWidth + 20
    frameTab.Width = lstvTablePrice.Width + 10
    
    frameFilters.Width = frameTab.Width
    Me.Width = frameTab.Width + 14
    cmbxKategoriya.Width = frameFilters.Width - cmbxKategoriya.Left - 6
    cmbxGruppa.Width = frameFilters.Width - cmbxGruppa.Left - 6
    cmbxPodgruppa.Width = frameFilters.Width - cmbxPodgruppa.Left - 6
    btnClose.Left = Me.Width - btnClose.Width - 10
    tbtnFiltr.Left = Me.Width - tbtnFiltr.Width - 10
    btnNabAdd.Left = btnClose.Left - btnNabAdd.Width - 10
    btnFavAdd.Left = btnNabAdd.Left - btnFavAdd.Width - 2
    btnETM.Left = btnFavAdd.Left - btnETM.Width - 2
    btnAVS.Left = btnETM.Left
    cmbxMagazin.Left = btnClose.Left
    frameProizvoditel.Width = btnETM.Left - frameProizvoditel.Left - 6
    cmbxProizvoditel.Width = frameProizvoditel.Width - 12
    lblResult.Left = frameTab.Width - lblResult.Width
    btnFind.Left = frameTab.Width - btnFind.Width - 6
    frameNazvanie.Width = btnFind.Left - frameNazvanie.Left - 6
'    txtNazvanie1.Width = frameNazvanie.Width / 4
'    txtNazvanie2.Left = txtNazvanie1.Left + txtNazvanie1.Width
'    txtNazvanie2.Width = (frameNazvanie.Width - 16) / 2
'    txtNazvanie3.Left = txtNazvanie2.Left + txtNazvanie2.Width
'    txtNazvanie3.Width = frameNazvanie.Width / 4
    txtNazvanie2.Left = 3
    txtNazvanie2.Width = frameNazvanie.Width - 9


End Sub

Private Sub tbtnFiltr_Click()
    If tbtnFiltr.Value Then
        frameFilters.Height = 84
        tbtnFiltr.Caption = ChrW(9650) 'вверх
    Else
        frameFilters.Height = 0
        tbtnFiltr.Caption = ChrW(9660) 'вниз
        If Not bBlock Then
            ClearFilter wshPrice
            Find_ItemsByText
        End If
    End If
    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
End Sub

Private Sub cmbxMagazin_Change()
    Select Case cmbxMagazin.ListIndex
        Case 0 'ЭТМ
            btnETM.Visible = True
            btnAVS.Visible = False
        Case 1 'АВС
            btnETM.Visible = False
            btnAVS.Visible = True
        Case Else
            btnETM.Visible = True
            btnAVS.Visible = False
    End Select
End Sub

Private Sub btnETM_Click()
    MagazinInfo mstrShpData(3), cmbxMagazin.ListIndex
End Sub

Private Sub btnAVS_Click()
    MagazinInfo mstrShpData(3), cmbxMagazin.ListIndex
End Sub

Private Sub btnFind_Click()
    Find_ItemsByText
End Sub

Private Sub cmbxKategoriya_Change()
    If Not bBlock Then Filter_CmbxChange 1
End Sub

Private Sub cmbxGruppa_Change()
    If Not bBlock Then Filter_CmbxChange 2
End Sub

Private Sub cmbxPodgruppa_Change()
    If Not bBlock Then Filter_CmbxChange 3
End Sub

Private Sub tbtnBD_Click()
    tbtnBD = True
End Sub

Private Sub tbtnFav_Click()
    If Not bBlock Then
        bBlock = True
        tbtnFav = False
        bBlock = False
        Me.Hide
        frmDBIzbrannoeExcel.Find_ItemsByText
        frmDBIzbrannoeExcel.Show
    End If
End Sub

Private Sub lblContent_Click() ' выровнять ширину столбцов по содержимому
   Dim colNum As Long
   For colNum = 0 To lstvTablePrice.ColumnHeaders.Count - 1
      Call SendMessage(lstvTablePrice.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvTablePrice.ColumnHeaders.Count - 1
      Call SendMessage(lstvTablePrice.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

Private Sub lstvTablePrice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку
    With lstvTablePrice
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
End Sub

Sub btnClose_Click() ' выгрузка формы
    Unload frmDBIzbrannoeExcel
    oExcelApp.Application.Quit
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

