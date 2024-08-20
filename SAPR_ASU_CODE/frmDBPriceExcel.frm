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
Dim FilePathName As String
Dim bInit As Boolean

Private Sub UserForm_Initialize() ' инициализация формы
    If Not bInit Then
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
        bInit = True
    End If
    InitCustomCCPMenu Me 'Контекстное меню для TextBox
End Sub

Sub run(vsoShape As Visio.Shape) 'Приняли шейп из модуля DB
    Dim ArtikulDB As String

    InitIzbrannoeExcelDB
'    FillExcel_cmbxProizvoditel cmbxProizvoditel, True
    
    Set glShape = vsoShape 'И определили его как глолбальный в форме frmDBPriceExcel
    ArtikulDB = glShape.Cells("Prop.ArtikulDB").ResultStr(0)
    
    'Открываем избранное, без поиска введенного в элемент, артикула

            bBlock = True
            cmbxProizvoditel.ListIndex = -1
            Load frmDBIzbrannoeExcel
'            frmDBIzbrannoeExcel.txtArtikul.Value = ArtikulDB
''            frmDBIzbrannoeExcel.tbtnFiltr.Value = False

            frmDBIzbrannoeExcel.txtArtikul.Value = ""
            frmDBIzbrannoeExcel.Find_ItemsByText_ADO
            bBlock = False
            frmDBIzbrannoeExcel.Show


'    If ArtikulDB <> "" Then
'        bBlock = True
'        For i = 0 To cmbxProizvoditel.ListCount - 1
'            If cmbxProizvoditel.List(i, 0) = glShape.Cells("Prop.ProizvoditelDB").ResultStr(0) Then cmbxProizvoditel.ListIndex = i
'        Next
'        If cmbxProizvoditel.ListIndex <> -1 And Not (ArtikulDB Like "Набор_*") Then
'            SetVarProizvoditelPrice
'            txtArtikul.Value = ArtikulDB
'            tbtnFiltr.Value = False
'            Find_ItemsByText_ADO
'            txtArtikul.Value = ""
'            bBlock = False
'            frmDBPriceExcel.Show
'        Else
'            bBlock = True
'            cmbxProizvoditel.ListIndex = -1
'            Load frmDBIzbrannoeExcel
'            frmDBIzbrannoeExcel.txtArtikul.Value = ArtikulDB
''            frmDBIzbrannoeExcel.tbtnFiltr.Value = False
'            frmDBIzbrannoeExcel.Find_ItemsByText_ADO
'            frmDBIzbrannoeExcel.txtArtikul.Value = ""
'            bBlock = False
'            frmDBIzbrannoeExcel.Show
'        End If
'    Else
'        ExcelAppQuit oExcelAppIzbrannoe
'        KillSAExcelProcess
'        frmDBPriceExcel.Show
'    End If
End Sub

Private Sub cmbxProizvoditel_Change()
    If Not bBlock Then
        SetVarProizvoditelPrice
        ClearFilter wshPrice
        'закрыть прайс
        ExcelAppQuit oExcelAppPrice
        KillSAExcelProcess
        Reset_FiltersCmbx_ADO
        lstvTablePrice.ListItems.Clear
    End If
End Sub

Sub SetVarProizvoditelPrice()
    Dim UserRange As Excel.Range
    Dim wshTemp As Excel.Worksheet
    
    For i = 0 To UBound(mProizvoditel)
        If cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) = mProizvoditel(i).Proizvoditel Then
            If mProizvoditel(i).FileName <> "" Then 'пустое имя файла - пропускаем
                Set PriceSettings = mProizvoditel(i)
                ExcelAppQuit oExcelAppPrice
                KillSAExcelProcess
                SetVarPrice
                Exit For
            End If
        End If
    Next
    Set wshTemp = Nothing
End Sub

Sub SetVarPrice()
    Set oExcelAppPrice = CreateObject("Excel.Application")
'    oExcelAppPrice.WindowState = xlMinimized
'    oExcelAppPrice.Visible = True
    Set wbExcelPrice = oExcelAppPrice.Workbooks.Open(PriceSettings.FileName)
    Set wshTemp = GetSheetExcel(wbExcelPrice, ExcelTemp)
    wshTemp.Cells.ClearContents
    Set wshPrice = Nothing
    Set wshPrice = wbExcelPrice.Worksheets(PriceSettings.NameListExcel)
    wshPrice.Range("A1").AutoFilter Field:=1
    Set RangePrice = wshPrice.AutoFilter.Range
End Sub

Sub Reset_FiltersCmbx_ADO()
    Dim SQLQuery As String
    If cmbxProizvoditel.ListIndex = -1 Then Exit Sub
    bBlock = True
    SQLQuery = "SELECT DISTINCT Категория FROM [" & PriceSettings.NameListExcel & "$];"
    Fill_ComboBox_ADO PriceSettings.FileName, SQLQuery, cmbxKategoriya
    SQLQuery = "SELECT DISTINCT Группа FROM [" & PriceSettings.NameListExcel & "$];"
    Fill_ComboBox_ADO PriceSettings.FileName, SQLQuery, cmbxGruppa
    SQLQuery = "SELECT DISTINCT Подгруппа FROM [" & PriceSettings.NameListExcel & "$];"
    Fill_ComboBox_ADO PriceSettings.FileName, SQLQuery, cmbxPodgruppa
    bBlock = False
    lstvTablePrice.ListItems.Clear
    lblResult.Caption = "Найдено записей: 0"
End Sub

 Sub Filter_CmbxChange_ADO(Ncmbx As Integer)
    Dim SQLQuery As String
    Dim fltrWhere As String
    Dim fltrKategoriya As String
    Dim fltrGruppa As String
    Dim fltrPodgruppa As String
    Dim fltrMode As Integer

    If cmbxKategoriya.ListIndex = -1 Then
        fltrKategoriya = ""
    Else
        fltrKategoriya = "Категория='" & cmbxKategoriya & "'"
    End If
    If cmbxGruppa.ListIndex = -1 Then
        fltrGruppa = ""
    Else
        fltrGruppa = "Группа='" & cmbxGruppa & "'"
    End If
    If cmbxPodgruppa.ListIndex = -1 Then
        fltrPodgruppa = ""
    Else
        fltrPodgruppa = "Подгруппа='" & cmbxPodgruppa & "'"
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
            fltrWhere = ""
        Case 1
            fltrWhere = " WHERE " & fltrPodgruppa
            bCallUpdatecmbxKategoriya = True
            bCallUpdatecmbxGruppa = True
        Case 2
            fltrWhere = " WHERE " & fltrGruppa
        Case 3
            fltrWhere = " WHERE " & fltrGruppa & " AND " & fltrPodgruppa
            bCallUpdatecmbxKategoriya = True
        Case 4
            fltrWhere = " WHERE " & fltrKategoriya
        Case 5
            fltrWhere = " WHERE " & fltrKategoriya & " AND " & fltrPodgruppa
            bCallUpdatecmbxGruppa = True
        Case 6
            fltrWhere = " WHERE " & fltrKategoriya & " AND " & fltrGruppa
        Case 7
            fltrWhere = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & " AND " & fltrPodgruppa
        Case Else
            fltrWhere = ""
            fltrKategoriya = ""
            fltrGruppa = ""
            fltrPodgruppa = ""
    End Select
'-------------------ФИЛЬТРАЦИЯ БЕЗ ПРИОРИТЕТА (Нет иерархии: Категория || Группа || Подгруппа)------------------------------------------------

'-------------------ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
    Select Case Ncmbx
        Case 1
            fltrWhere = " WHERE " & fltrKategoriya
            fltrGruppa = ""
            fltrPodgruppa = ""
            bBlock = True
            cmbxGruppa.Clear
            cmbxPodgruppa.Clear
            bBlock = False
            bCallUpdatecmbxGruppa = True
            bCallUpdatecmbxPodgruppa = True
        Case 2
            fltrWhere = IIf(fltrKategoriya = "", " WHERE " & fltrGruppa, " WHERE " & fltrKategoriya & " AND " & fltrGruppa)
            fltrPodgruppa = ""
            bBlock = True
            cmbxPodgruppa.Clear
            bBlock = False
            bCallUpdatecmbxPodgruppa = True
        Case 3
            'Работают варианты 1,3,5,7 из ФИЛЬТРАЦИЯ БЕЗ ПРИОРИТЕТА
        Case Else
            fltrWhere = ""
            fltrKategoriya = ""
            fltrGruppa = ""
            fltrPodgruppa = ""
    End Select
'-------------------ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
    SQLQuery = "SELECT * FROM [" & PriceSettings.NameListExcel & "$] " & fltrWhere & ";"
    lstvTablePrice.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable_ADO(PriceSettings.FileName, SQLQuery, lstvTablePrice)
    lstvTablePrice.Visible = True
    Fill_FiltersByResultSQLQuery_ADO
    ReSize
End Sub

Sub Find_ItemsByText_ADO()
    Dim SQLQuery As String
    Dim findMode As Integer
    Dim findWhat As String
    Dim fltrWhere As String
    Dim findArtikul As String
    Dim findNazvanie As String
    
    If cmbxProizvoditel.ListIndex = -1 Then Exit Sub
    
    If txtArtikul.Value = "" Then
        findArtikul = ""
    Else
        findArtikul = "Артикул LIKE '%" & txtArtikul.Value & "%'"
    End If
    
    If txtNazvanie2.Value = "" Then
        findNazvanie = ""
    Else
        findNazvanie = "Название LIKE '%" & Replace(txtNazvanie2.Value, " ", "%") & "%'"
    End If
    
    findMode = IIf(findArtikul = "", 0, 2) + IIf(findNazvanie = "", 0, 1)

    '*   Арт Наз
    '0   0   0
    '1   0   1
    '2   1   0
    '3   1   1

    Select Case findMode
        Case 0
            findWhat = ""
        Case 1
            findWhat = " WHERE " & findNazvanie
        Case 2
            findWhat = " WHERE " & findArtikul
        Case 3
            findWhat = " WHERE " & findArtikul & " AND " & findNazvanie
        Case Else
            findWhat = ""
    End Select
    
    '---ФИЛЬТРАЦИЯ БЕЗ ПРИОРИТЕТА (Нет иерархии: Категория || Группа || Подгруппа)---
    
    If cmbxKategoriya.ListIndex = -1 Then
        fltrKategoriya = ""
    Else
        fltrKategoriya = "Категория='" & cmbxKategoriya & "'"
    End If
    If cmbxGruppa.ListIndex = -1 Then
        fltrGruppa = ""
    Else
        fltrGruppa = "Группа='" & cmbxGruppa & "'"
    End If
    If cmbxPodgruppa.ListIndex = -1 Then
        fltrPodgruppa = ""
    Else
        fltrPodgruppa = "Подгруппа='" & cmbxPodgruppa & "'"
    End If
    
    fltrWhere = IIf(fltrKategoriya = "", "", " AND " & fltrKategoriya) & _
                IIf(fltrGruppa = "", "", " AND " & fltrGruppa) & _
                IIf(fltrPodgruppa = "", "", " AND " & fltrPodgruppa)
    SQLQuery = "SELECT * FROM [" & PriceSettings.NameListExcel & "$] " & findWhat & fltrWhere & ";"
    lstvTablePrice.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable_ADO(PriceSettings.FileName, SQLQuery, lstvTablePrice)
    lstvTablePrice.Visible = True
    bCallUpdatecmbxKategoriya = True
    bCallUpdatecmbxGruppa = True
    bCallUpdatecmbxPodgruppa = True
    Fill_FiltersByResultSQLQuery_ADO
    ReSize
End Sub

Sub Fill_FiltersByResultSQLQuery_ADO()
    Dim SQLQuery As String
    Dim i As Integer
    Dim scmbxKategoriyaValue As String
    Dim scmbxGruppaValue As String
    Dim scmbxPodgruppaValue As String
    
    bBlock = True
    If bCallUpdatecmbxKategoriya Then
        scmbxKategoriyaValue = cmbxKategoriya
        SQLQuery = "SELECT DISTINCT Категория FROM (" & sLastSQLQuery & ");"
        Fill_ComboBox_ADO PriceSettings.FileName, SQLQuery, cmbxKategoriya
        For i = 0 To cmbxKategoriya.ListCount - 1
            If cmbxKategoriya.List(i, 0) = scmbxKategoriyaValue Then cmbxKategoriya.ListIndex = i
        Next
        bCallUpdatecmbxKategoriya = False
    End If
    If bCallUpdatecmbxGruppa Then
        scmbxGruppaValue = cmbxGruppa
        SQLQuery = "SELECT DISTINCT Группа FROM (" & sLastSQLQuery & ");"
        Fill_ComboBox_ADO PriceSettings.FileName, SQLQuery, cmbxGruppa
        For i = 0 To cmbxGruppa.ListCount - 1
            If cmbxGruppa.List(i, 0) = scmbxGruppaValue Then cmbxGruppa.ListIndex = i
        Next
        bCallUpdatecmbxGruppa = False
    End If
    If bCallUpdatecmbxPodgruppa Then
        scmbxPodgruppaValue = cmbxPodgruppa
        SQLQuery = "SELECT DISTINCT Подгруппа FROM (" & sLastSQLQuery & ");"
        Fill_ComboBox_ADO PriceSettings.FileName, SQLQuery, cmbxPodgruppa
        For i = 0 To cmbxPodgruppa.ListCount - 1
            If cmbxPodgruppa.List(i, 0) = scmbxPodgruppaValue Then cmbxPodgruppa.ListIndex = i
        Next
        bCallUpdatecmbxPodgruppa = False
    End If
    bBlock = False
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
    mstrShpData(4) = cmbxProizvoditel 'Производитель
End Sub

Private Sub lstvTablePrice_DblClick()
    Dim vsoShape As Visio.Shape
    Set vsoShape = glShape
    GoSub SetDB
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> glShape And ShapeSATypeIs(vsoShape, ShapeSAType(glShape)) Then
                GoSub SetDB
            End If
        Next
    End If
    
    btnClose_Click
    Exit Sub
    
SetDB:
    On Error GoTo errGuard
    With vsoShape
        .Cells("User.KodProizvoditelyaDB").Formula = """"""
        .Cells("User.KodPoziciiDB").Formula = """"""
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(1), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(0) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrShpData(2) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrShpData(3) & """"
    End With
    err.Clear
    On Error GoTo 0
    Return
    
errGuard:
    With vsoShape
        .Cells("Prop.NazvanieDB").FormulaForce = """" & Replace(mstrShpData(1), """", """""") & """"
        .Cells("Prop.NazvanieDB.Type").FormulaForce = 0
        .Cells("Prop.NazvanieDB.Format").FormulaForce = """"""
        .Cells("Prop.ArtikulDB").FormulaForce = """" & mstrShpData(0) & """"
        .Cells("Prop.ArtikulDB.Type").FormulaForce = 0
        .Cells("Prop.ArtikulDB.Format").FormulaForce = """"""
        .Cells("Prop.ProizvoditelDB").FormulaForce = """" & mstrShpData(4) & """"
        .Cells("Prop.ProizvoditelDB.Type").FormulaForce = 0
        .Cells("Prop.ProizvoditelDB.Format").FormulaForce = """"""
        .Cells("Prop.CenaDB").FormulaForce = """" & mstrShpData(2) & """"
        .Cells("Prop.CenaDB.Type").FormulaForce = 0
        .Cells("Prop.CenaDB.Format").FormulaForce = """"""
        .Cells("Prop.EdDB").FormulaForce = """" & mstrShpData(3) & """"
        .Cells("Prop.EdDB.Type").FormulaForce = 0
        .Cells("Prop.EdDB.Format").FormulaForce = """"""
    End With
    Return
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
            txtNazvanie2.Value = ""
            txtArtikul.Value = ""
            Reset_FiltersCmbx_ADO
            lstvTablePrice.ListItems.Clear
            bBlock = True
            cmbxKategoriya.ListIndex = -1
            cmbxGruppa.ListIndex = -1
            cmbxPodgruppa.ListIndex = -1
            bBlock = False
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
    FindArticulInBrowser mstrShpData(0), cmbxMagazin.ListIndex
End Sub

Private Sub btnAVS_Click()
    FindArticulInBrowser mstrShpData(0), cmbxMagazin.ListIndex
End Sub

Private Sub btnFind_Click()
    Find_ItemsByText_ADO
End Sub

Private Sub cmbxKategoriya_Change()
    If Not bBlock Then Filter_CmbxChange_ADO 1
End Sub

Private Sub cmbxGruppa_Change()
    If Not bBlock Then Filter_CmbxChange_ADO 2
End Sub

Private Sub cmbxPodgruppa_Change()
    If Not bBlock Then Filter_CmbxChange_ADO 3
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
        
        ExcelAppQuit oExcelAppPrice
        KillSAExcelProcess
        If frmDBIzbrannoeExcel.lstvTableIzbrannoe.ListItems.Count = 0 Then
            frmDBIzbrannoeExcel.Find_ItemsByText_ADO
        End If
'        InitCustomCCPMenu frmDBIzbrannoeExcel 'Контекстное меню для TextBox
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
    ExcelAppQuit oExcelAppPrice
    ExcelAppQuit oExcelAppIzbrannoe
    KillSAExcelProcess
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub
