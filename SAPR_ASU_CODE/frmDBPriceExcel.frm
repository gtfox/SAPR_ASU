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
Dim mstrShpData(6) As String
Public bBlock As Boolean
Dim NameQueryDef As String


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

    Dim SQLQuery As String

    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель, Производители.КодПроизводителя " & _
                "FROM Производители;"
                
    Fill_cmbxProizvoditel DBNameIzbrannoe, SQLQuery, cmbxProizvoditel, True

    Load frmDBIzbrannoeExcel
    'frmDBIzbrannoeExcel.Find_ItemsByText

End Sub

Sub run(vsoShape As Visio.Shape) 'Приняли шейп из модуля DB
    Dim ArtikulDB As String

    Set glShape = vsoShape 'И определили его как глолбальный в форме frmDBPriceExcel
    ArtikulDB = glShape.Cells("Prop.ArtikulDB").ResultStr(0)
    If ArtikulDB <> "" Then
        bBlock = True
        For i = 0 To cmbxProizvoditel.ListCount - 1
            If cmbxProizvoditel.List(i, 2) = glShape.Cells("User.KodProizvoditelyaDB").Result(0) Then cmbxProizvoditel.ListIndex = i
        Next
        
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
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(DBName, SQLQuery, NameQueryDef, lstvTablePrice)
'lstvTablePrice.Visible = True

    Fill_FiltersByResultSQLQuery DBName, fltrKategoriya, fltrGruppa, fltrPodgruppa

    ReSize

    'Find_ItemsByText
    
End Sub

Sub Fill_FiltersByResultSQLQuery(DBName As String, fltrKategoriya As String, fltrGruppa As String, fltrPodgruppa As String)
    Dim SQLQuery As String

    If fltrKategoriya = "" Then
        SQLQuery = "SELECT FilterSQLQuery.КатегорииКод, Категории.Категория " & _
                    "FROM Категории INNER JOIN FilterSQLQuery ON Категории.КодКатегории = FilterSQLQuery.КатегорииКод " & _
                    "GROUP BY FilterSQLQuery.КатегорииКод, Категории.Категория;"
        Fill_ComboBox DBName, SQLQuery, cmbxKategoriya
    End If
    
    If fltrGruppa = "" Then
        SQLQuery = "SELECT FilterSQLQuery.ГруппыКод, Группы.Группа " & _
                    "FROM Группы INNER JOIN FilterSQLQuery ON Группы.КодГруппы = FilterSQLQuery.ГруппыКод " & _
                    "GROUP BY FilterSQLQuery.ГруппыКод, Группы.Группа;"
        Fill_ComboBox DBName, SQLQuery, cmbxGruppa
    End If
    
    If fltrPodgruppa = "" Then
        SQLQuery = "SELECT FilterSQLQuery.ПодгруппыКод, Подгруппы.Подгруппа " & _
                    "FROM Подгруппы INNER JOIN FilterSQLQuery ON Подгруппы.КодПодгруппы = FilterSQLQuery.ПодгруппыКод " & _
                    "GROUP BY FilterSQLQuery.ПодгруппыКод, Подгруппы.Подгруппа;"
        Fill_ComboBox DBName, SQLQuery, cmbxPodgruppa
    End If

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
        lblResult.Caption = "Найдено записей: " & Fill_lstvTable(DBName, SQLQuery, NameQueryDef, lstvTablePrice)
'lstvTablePrice.Visible = True
        Fill_FiltersByResultSQLQuery DBName, "", "", ""
    Else
        NameQueryDef = ""
        SQLQuery = "SELECT FilterSQLQuery.КодПозиции, FilterSQLQuery.Артикул, FilterSQLQuery.Название, FilterSQLQuery.Цена, FilterSQLQuery.КатегорииКод, FilterSQLQuery.ГруппыКод, FilterSQLQuery.ПодгруппыКод, FilterSQLQuery.ПроизводительКод, FilterSQLQuery.ЕдиницыКод, FilterSQLQuery.Единица " & _
                   "FROM Единицы INNER JOIN FilterSQLQuery ON Единицы.КодЕдиницы = FilterSQLQuery.ЕдиницыКод " & findWHERE & ";"
'lstvTablePrice.Visible = False
        lblResult.Caption = "Найдено записей: " & Fill_lstvTable(DBName, SQLQuery, NameQueryDef, lstvTablePrice)
'lstvTablePrice.Visible = True
    End If

    ReSize
 
End Sub

Private Sub Reset_FiltersCmbx()
    Dim DBName As String
    Dim SQLQuery As String
    If cmbxProizvoditel.ListIndex = -1 Then Exit Sub
    bBlock = True
    DBName = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1)
    SQLQuery = "SELECT Категории.КодКатегории, Категории.Категория " & _
                "FROM Категории;"
    Fill_ComboBox DBName, SQLQuery, cmbxKategoriya
    SQLQuery = "SELECT Группы.КодГруппы, Группы.Группа " & _
                "FROM Группы;"
    Fill_ComboBox DBName, SQLQuery, cmbxGruppa
    SQLQuery = "SELECT Подгруппы.КодПодгруппы, Подгруппы.Подгруппа " & _
                "FROM Подгруппы;"
    Fill_ComboBox DBName, SQLQuery, cmbxPodgruppa
    bBlock = False
    lstvTablePrice.ListItems.Clear
    lblResult.Caption = "Найдено записей: 0"
End Sub

Private Sub btnFavAdd_Click()
    Dim Mstr() As String
    If mstrShpData(2) <> "" Then
        If Not bBlock Then
            bBlock = True
            tbtnFav = False
            bBlock = False
            Me.Hide
        End If
        Load frmDBAddToIzbrannoe
        Mstr = Split(Replace(mstrShpData(1), """", ""), "/")
        frmDBAddToIzbrannoe.run mstrShpData(3), Replace(mstrShpData(2), """", """"""), mstrShpData(5), cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2), Mstr(2)
    End If
End Sub

Private Sub btnNabAdd_Click()
    Dim Mstr() As String
    If mstrShpData(2) <> "" Then
        Me.Hide
        Load frmDBAddToNabor
        Mstr = Split(Replace(mstrShpData(1), """", ""), "/")
        frmDBAddToNabor.run mstrShpData(3), Replace(mstrShpData(2), """", """"""), mstrShpData(5), cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2), Mstr(2)
    End If
End Sub

Private Sub lstvTablePrice_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Если в таблице ткнуть на строку с номером больше 30000 то сюда попадет первая строка!!!
    mstrShpData(0) = cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2)
    mstrShpData(1) = Item.Key
    mstrShpData(2) = Item.SubItems(1)
    mstrShpData(3) = Item
    mstrShpData(4) = cmbxProizvoditel.Value
    mstrShpData(5) = Item.SubItems(2)
    mstrShpData(6) = Item.SubItems(3)

End Sub

Private Sub lstvTablePrice_DblClick()
    Dim vsoShape As Visio.Shape
    
    With glShape
        .Cells("User.KodProizvoditelyaDB").Formula = mstrShpData(0)
        .Cells("User.KodPoziciiDB").Formula = Replace(mstrShpData(1), """", "")
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(2), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(3) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrShpData(5) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrShpData(6) & """"
    End With
    
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> glShape And ShapeSATypeIs(vsoShape, ShapeSAType(glShape)) Then
                With vsoShape
                    .Cells("User.KodProizvoditelyaDB").Formula = mstrShpData(0)
                    .Cells("User.KodPoziciiDB").Formula = Replace(mstrShpData(1), """", "")
                    .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(2), """", """""") & """"
                    .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(3) & """"
                    .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
                    .Cells("Prop.CenaDB").Formula = """" & mstrShpData(5) & """"
                    .Cells("Prop.EdDB").Formula = """" & mstrShpData(6) & """"
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
        Reset_FiltersCmbx
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

Private Sub cmbxProizvoditel_Change()
   If Not bBlock Then Reset_FiltersCmbx
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
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

