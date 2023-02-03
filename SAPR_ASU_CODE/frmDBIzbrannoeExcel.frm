
'------------------------------------------------------------------------------------------------------------
' Module        : frmDBIzbrannoeExcel - Форма поиска и задания данных для элемента схемы из БД Избранное. В одном файле разные производители.
' Author        : gtfox
' Date          : 2023.01.30
' Description   : Выбор данных из БД Избранное, фильтрация по категориям и полнотекстовый поиск
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

Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи
Dim mstrShpData(7) As String
Public bBlock As Boolean
Dim NameQueryDef As String
Dim mstrVybPozVNabore(7) As String




Private Sub UserForm_Initialize() ' инициализация формы

    lstvTableIzbrannoe.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvTableIzbrannoe.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTableIzbrannoe.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTableIzbrannoe.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
    lstvTableIzbrannoe.ColumnHeaders.Add , , "Ед." ' SubItems(3)
    lstvTableIzbrannoe.ColumnHeaders.Add , , "Производитель" ' SubItems(4)
    lstvTableIzbrannoe.ColumnHeaders.Add , , "    " ' SubItems(5)
   
    lstvTableNabor.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvTableNabor.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTableNabor.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTableNabor.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
    lstvTableNabor.ColumnHeaders.Add , , "Ед." ' SubItems(3)
    lstvTableNabor.ColumnHeaders.Add , , "Производитель" ' SubItems(4)
    lstvTableNabor.ColumnHeaders.Add , , "Кол-во" ' SubItems(5)
    lstvTableNabor.ColumnHeaders.Add , , "    " ' SubItems(6)

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
    Me.Top = 350
    lblResult.Top = Me.Height - 35
    
    tbtnFiltr.Caption = ChrW(9650)
'    tbtnBD = False
    tbtnFav = True

    Dim SQLQuery As String

    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель, Производители.КодПроизводителя " & _
                "FROM Производители;"
                
    FillExcel_cmbxProizvoditel DBNameIzbrannoeExcel, SQLQuery, cmbxProizvoditel
    
    Reset_FiltersCmbx

End Sub

Private Sub Filter_CmbxChange(Ncmbx As Integer)
    Dim SQLQuery As String
    Dim fltrKategoriya As String
    Dim fltrGruppa As String
    Dim fltrPodgruppa As String
    Dim fltrProizvoditel As String
    Dim fltrMode As Integer
    Dim fltrWHERE As String
    Dim DBName As String

    If cmbxKategoriya.ListIndex = -1 Then
        fltrKategoriya = ""
    Else
        fltrKategoriya = "Избранное.КатегорииКод=" & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1)
    End If
    If cmbxGruppa.ListIndex = -1 Then
        fltrGruppa = ""
    Else
        fltrGruppa = "Избранное.ГруппыКод=" & cmbxGruppa.List(cmbxGruppa.ListIndex, 1)
    End If
    If cmbxPodgruppa.ListIndex = -1 Then
        fltrPodgruppa = ""
    Else
        fltrPodgruppa = "Избранное.ПодгруппыКод=" & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1)
    End If
    
    If cmbxProizvoditel.ListIndex = -1 Then
        fltrProizvoditel = ""
    Else
        fltrProizvoditel = "" & IIf(cmbxProizvoditel.ListIndex = -1, "", " AND Производители.Производитель=" & """" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) & """")
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
            If cmbxProizvoditel.ListIndex = -1 Then
                fltrWHERE = ""
            Else
                fltrWHERE = "" & IIf(cmbxProizvoditel.ListIndex = -1, "", " WHERE Производители.Производитель=" & """" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) & """")
            End If
        Case 1
            fltrWHERE = " WHERE " & fltrPodgruppa & fltrProizvoditel
        Case 2
            fltrWHERE = " WHERE " & fltrGruppa & fltrProizvoditel
        Case 3
            fltrWHERE = " WHERE " & fltrGruppa & " AND " & fltrPodgruppa & fltrProizvoditel
        Case 4
            fltrWHERE = " WHERE " & fltrKategoriya & fltrProizvoditel
        Case 5
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrPodgruppa & fltrProizvoditel
        Case 6
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & fltrProizvoditel
        Case 7
            fltrWHERE = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & " AND " & fltrPodgruppa & fltrProizvoditel
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
            fltrWHERE = " WHERE " & fltrKategoriya & fltrProizvoditel
            fltrGruppa = ""
            fltrPodgruppa = ""
            bBlock = True
            cmbxGruppa.Clear
            cmbxPodgruppa.Clear
            bBlock = False
        Case 2
            fltrWHERE = IIf(fltrKategoriya = "", " WHERE " & fltrGruppa, " WHERE " & fltrKategoriya & " AND " & fltrGruppa) & fltrProizvoditel
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


        SQLQuery = "SELECT Избранное.КодПозиции, Избранное.Артикул, Избранное.Название, Избранное.Цена, Избранное.КатегорииКод, Избранное.ГруппыКод, Избранное.ПодгруппыКод, Избранное.ПроизводительКод, Производители.Производитель, Избранное.ЕдиницыКод, Единицы.Единица " & _
                   "FROM Единицы INNER JOIN (Производители INNER JOIN Избранное ON Производители.КодПроизводителя = Избранное.ПроизводительКод) ON Единицы.КодЕдиницы = Избранное.ЕдиницыКод " & fltrWHERE & ";"
                
    DBName = DBNameIzbrannoeExcel
    
    NameQueryDef = "FilterSQLQuery"
    
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(DBName, SQLQuery, NameQueryDef, lstvTableIzbrannoe, 1)

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
    Dim findProizvoditel As String
    Dim findArtikul As String
    Dim findNazvanie As String

    DBName = DBNameIzbrannoeExcel
    
    If txtArtikul.Value = "" Then
        findArtikul = ""
    Else
        findArtikul = "Избранное.Артикул like ""*" & txtArtikul.Value & "*"""
    End If
    
    If txtNazvanie1.Value = "" And txtNazvanie2.Value = "" And txtNazvanie3.Value = "" Then
        findNazvanie = ""
    Else
        findNazvanie = "Избранное.Название like ""*" & txtNazvanie1.Value & "*" & Replace(txtNazvanie2.Value, " ", "*") & "*" & txtNazvanie3.Value & "*"""
    End If
    
    If cmbxProizvoditel.ListIndex = -1 Then
        findProizvoditel = ""
    Else
        findProizvoditel = "" & IIf(cmbxProizvoditel.ListIndex = -1, "", " AND Производители.Производитель=" & """" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) & """")
    End If
    
    findMode = IIf(findArtikul = "", 0, 2) + IIf(findNazvanie = "", 0, 1)

    '*   Арт Наз
    '0   0   0
    '1   0   1
    '2   1   0
    '3   1   1

    Select Case findMode
        Case 0
            If cmbxProizvoditel.ListIndex = -1 Then
                findWHERE = ""
            Else
                findWHERE = "" & IIf(cmbxProizvoditel.ListIndex = -1, "", " WHERE Производители.Производитель=" & """" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0) & """")
            End If
        Case 1
            findWHERE = " WHERE " & findNazvanie & findProizvoditel
        Case 2
            findWHERE = " WHERE " & findArtikul & findProizvoditel
        Case 3
            findWHERE = " WHERE " & findArtikul & " AND " & findNazvanie & findProizvoditel
        Case Else
            findWHERE = ""
    End Select

    If cmbxKategoriya.ListIndex = -1 And cmbxGruppa.ListIndex = -1 And cmbxPodgruppa.ListIndex = -1 Then
        NameQueryDef = "FilterSQLQuery"
        SQLQuery = "SELECT Избранное.КодПозиции, Избранное.Артикул, Избранное.Название, Избранное.Цена, Избранное.КатегорииКод, Избранное.ГруппыКод, Избранное.ПодгруппыКод, Избранное.ПроизводительКод, Производители.Производитель, Избранное.ЕдиницыКод, Единицы.Единица " & _
                   "FROM Единицы INNER JOIN (Производители INNER JOIN Избранное ON Производители.КодПроизводителя = Избранное.ПроизводительКод) ON Единицы.КодЕдиницы = Избранное.ЕдиницыКод " & findWHERE & ";"
        lblResult.Caption = "Найдено записей: " & Fill_lstvTable(DBName, SQLQuery, NameQueryDef, lstvTableIzbrannoe, 1)
        Fill_FiltersByResultSQLQuery DBName, "", "", ""
    Else
        NameQueryDef = ""
        SQLQuery = "SELECT FilterSQLQuery.КодПозиции, FilterSQLQuery.Артикул, FilterSQLQuery.Название, FilterSQLQuery.Цена, FilterSQLQuery.КатегорииКод, FilterSQLQuery.ГруппыКод, FilterSQLQuery.ПодгруппыКод, FilterSQLQuery.ПроизводительКод, Производители.Производитель, FilterSQLQuery.ЕдиницыКод, Единицы.Единица " & _
                   "FROM Единицы INNER JOIN (Производители INNER JOIN FilterSQLQuery ON Производители.КодПроизводителя = FilterSQLQuery.ПроизводительКод) ON Единицы.КодЕдиницы = FilterSQLQuery.ЕдиницыКод " & findWHERE & ";"
        lblResult.Caption = "Найдено записей: " & Fill_lstvTable(DBName, SQLQuery, NameQueryDef, lstvTableIzbrannoe, 1)
    End If

    ReSize
 
End Sub

Private Sub btnFavDel_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись из избранного?" & vbCrLf & vbCrLf & "Артикул: " & mstrShpData(3) & vbCrLf & "Название: " & mstrShpData(2) & vbCrLf & "Цена: " & mstrShpData(5) & vbCrLf & "Производитель: " & mstrShpData(4), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Избранного") = vbYes Then
        If mstrShpData(6) <> "" Then
            DBName = DBNameIzbrannoeExcel
            SQLQuery = "DELETE Избранное.* " & _
                        "FROM Избранное " & _
                        "WHERE Избранное.КодПозиции=" & mstrShpData(6) & ";"
            ExecuteSQL DBName, SQLQuery
            lstvTableNabor.ListItems.Clear
            Find_ItemsByText
        End If
    End If
End Sub

Private Sub btnNabDel_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись из набора?" & vbCrLf & vbCrLf & "Артикул: " & mstrVybPozVNabore(3) & vbCrLf & "Название: " & mstrVybPozVNabore(2) & vbCrLf & "Цена: " & mstrVybPozVNabore(5) & vbCrLf & "Производитель: " & mstrVybPozVNabore(4), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Набора") = vbYes Then
        If mstrVybPozVNabore(6) <> "" Then
            DBName = DBNameIzbrannoeExcel
            SQLQuery = "DELETE Наборы.* " & _
                        "FROM Наборы " & _
                        "WHERE Наборы.КодПозиции=" & mstrVybPozVNabore(6) & ";"
            ExecuteSQL DBName, SQLQuery
            lblSostav.Caption = "Состав набора: " & Fill_lstvTableNabor(DBName, mstrShpData(6), lstvTableNabor)

            SQLQuery = "UPDATE Избранное SET Избранное.Цена = """ & CalcCenaNabora(lstvTableNabor) & """" & _
                        " WHERE Избранное.КодПозиции = " & mstrShpData(6) & ";"
            ExecuteSQL DBName, SQLQuery
            Find_ItemsByText
        End If
    End If
End Sub

Private Sub Reset_FiltersCmbx()
    Dim oExcelApp As Excel.Application
    Dim wbExcel As Excel.Workbook
    Dim wshTemp As Excel.Worksheet
    Dim sPath As String
    Dim UserRange As Excel.Range
    Dim lLastRow As Long
    Dim i As Integer
    Dim j As Integer
    Dim mFilter() As String

    Set oExcelApp = CreateObject("Excel.Application")
    sPath = Visio.ActiveDocument.path
    Set wbExcel = oExcelApp.Workbooks.Open(sPath & DBNameIzbrannoeExcel)
    Set wshTemp = oExcelApp.Worksheets(Exceltemp)

    lLastRow = oExcelApp.Sheets(ExcelIzbrannoe).Cells(oExcelApp.Sheets(ExcelIzbrannoe).Rows.Count, 1).End(xlUp).Row
    oExcelApp.Worksheets(ExcelIzbrannoe).Range("F2:H" & lLastRow).Copy oExcelApp.Worksheets(Exceltemp).Range("A1")
    Set UserRange = wshTemp.Range("A1:C" & lLastRow - 1)
    UserRange.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlNo
    
    bBlock = True
    
    lLastRow = wshTemp.Cells(wshTemp.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lLastRow
        cmbxKategoriya.AddItem wshTemp.Cells(i, 1)
    Next
    lLastRow = wshTemp.Cells(wshTemp.Rows.Count, 2).End(xlUp).Row
    For i = 1 To lLastRow
        cmbxGruppa.AddItem wshTemp.Cells(i, 2)
    Next
    lLastRow = wshTemp.Cells(wshTemp.Rows.Count, 3).End(xlUp).Row
    For i = 1 To lLastRow
        cmbxPodgruppa.AddItem wshTemp.Cells(i, 3)
    Next
    
    wbExcel.Close SaveChanges:=False
    oExcelApp.Application.Quit
    
    bBlock = False
    lstvTableIzbrannoe.ListItems.Clear
    lblResult.Caption = "Найдено записей: 0"


End Sub

Private Sub lstvTableIzbrannoe_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Если в таблице ткнуть на строку с номером больше 30000 то сюда попадет первая строка!!!
    Dim Mstr() As String
    Dim colNum As Long
    
    Mstr = Split(Replace(Item.Key, """", ""), "/")

    mstrShpData(0) = Mstr(1)
    mstrShpData(1) = Item.Key
    mstrShpData(2) = Item.SubItems(1)
    mstrShpData(3) = Item
    mstrShpData(4) = Item.SubItems(4)
    mstrShpData(5) = Item.SubItems(2)
    mstrShpData(6) = Mstr(0)
    mstrShpData(7) = Item.SubItems(3)
    
    If Item.ForeColor = NaboryColor Then
        lblSostav.Caption = "Состав набора: " & Fill_lstvTableNabor(DBNameIzbrannoeExcel, mstrShpData(6), lstvTableNabor)
        lstvTableNabor.Width = frmMinWdth
        'выровнять ширину столбцов по заголовкам
        For colNum = 0 To lstvTableNabor.ColumnHeaders.Count - 1
            Call SendMessage(lstvTableNabor.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
        Next
        Me.Height = lstvTableNabor.Top + lstvTableNabor.Height + 26
    Else
        lstvTableNabor.ListItems.Clear
        Me.Height = frameTab.Top + frameTab.Height + 36
        lblSostav.Caption = ""
    End If
    
    ReSize
    
End Sub

Private Sub lstvTableIzbrannoe_DblClick()
    Dim vsoShape As Visio.Shape
    
    With frmDBPriceAccess.glShape
        .Cells("User.KodProizvoditelyaDB").Formula = mstrShpData(0)
        .Cells("User.KodPoziciiDB").Formula = Replace(mstrShpData(1), """", "")
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(2), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(3) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrShpData(5) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrShpData(7) & """"
    End With
    
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> frmDBPriceAccess.glShape And ShapeSATypeIs(vsoShape, ShapeSAType(frmDBPriceAccess.glShape)) Then
                With vsoShape
                    .Cells("User.KodProizvoditelyaDB").Formula = mstrShpData(0)
                    .Cells("User.KodPoziciiDB").Formula = Replace(mstrShpData(1), """", "")
                    .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrShpData(2), """", """""") & """"
                    .Cells("Prop.ArtikulDB").Formula = """" & mstrShpData(3) & """"
                    .Cells("Prop.ProizvoditelDB").Formula = """" & mstrShpData(4) & """"
                    .Cells("Prop.CenaDB").Formula = """" & mstrShpData(5) & """"
                    .Cells("Prop.EdDB").Formula = """" & mstrShpData(7) & """"
                End With
            End If
        Next
    End If
    
    btnClose_Click
    
End Sub

Private Sub lstvTableNabor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim Mstr() As String

    Mstr = Split(Replace(Item.Key, """", ""), "/")
    
    mstrVybPozVNabore(0) = Mstr(1)
    mstrVybPozVNabore(1) = Item.Key
    mstrVybPozVNabore(2) = Item.SubItems(1)
    mstrVybPozVNabore(3) = Item
    mstrVybPozVNabore(4) = Item.SubItems(4)
    mstrVybPozVNabore(5) = Item.SubItems(2)
    mstrVybPozVNabore(6) = Mstr(0)
    mstrVybPozVNabore(7) = Item.SubItems(3)
    
End Sub

Private Sub lstvTableNabor_DblClick()
    Dim vsoShape As Visio.Shape
    
    With frmDBPriceAccess.glShape
        .Cells("User.KodProizvoditelyaDB").Formula = mstrVybPozVNabore(0)
        .Cells("User.KodPoziciiDB").Formula = Replace(mstrVybPozVNabore(1), """", "")
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrVybPozVNabore(2), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrVybPozVNabore(3) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrVybPozVNabore(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrVybPozVNabore(5) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrVybPozVNabore(7) & """"
    End With
    
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> frmDBPriceAccess.glShape And ShapeSATypeIs(vsoShape, ShapeSAType(frmDBPriceAccess.glShape)) Then
                With vsoShape
                    .Cells("User.KodProizvoditelyaDB").Formula = mstrVybPozVNabore(0)
                    .Cells("User.KodPoziciiDB").Formula = Replace(mstrVybPozVNabore(1), """", "")
                    .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrVybPozVNabore(2), """", """""") & """"
                    .Cells("Prop.ArtikulDB").Formula = """" & mstrVybPozVNabore(3) & """"
                    .Cells("Prop.ProizvoditelDB").Formula = """" & mstrVybPozVNabore(4) & """"
                    .Cells("Prop.CenaDB").Formula = """" & mstrVybPozVNabore(5) & """"
                    .Cells("Prop.EdDB").Formula = """" & mstrVybPozVNabore(7) & """"
                End With
            End If
        Next
    End If
    
    btnClose_Click
    
End Sub



Private Sub ReSize() ' изменение формы. Зависит от длины в lstvTableIzbrannoe
    Dim TableIzbrannoeWidth As Single
    Dim TableNaborWidth As Single
    
    lstvTableIzbrannoe.Width = frmMinWdth

'    lblContent_Click
    lblHeaders_Click

    If lstvTableIzbrannoe.ListItems.Count < 1 Then Exit Sub
    
    TableIzbrannoeWidth = lstvTableIzbrannoe.ListItems(1).Width
    
    If lstvTableNabor.ListItems.Count < 1 Then
        TableNaborWidth = 0
    Else
        TableNaborWidth = lstvTableNabor.ListItems(1).Width
    End If

    If TableIzbrannoeWidth > TableNaborWidth Then
        If TableIzbrannoeWidth < frmMinWdth Then
            TableIzbrannoeWidth = frmMinWdth
        End If
    Else
        If TableNaborWidth > frmMinWdth Then
            TableIzbrannoeWidth = TableNaborWidth
        Else
            TableIzbrannoeWidth = frmMinWdth
        End If
    End If
    
    lstvTableIzbrannoe.Width = TableIzbrannoeWidth
    
    lstvTableNabor.Width = lstvTableIzbrannoe.Width
    frameTab.Width = lstvTableIzbrannoe.Width + 10
    
    frameFilters.Width = frameTab.Width
    Me.Width = frameTab.Width + 14
    cmbxKategoriya.Width = frameFilters.Width - cmbxKategoriya.Left - 6
    cmbxGruppa.Width = frameFilters.Width - cmbxGruppa.Left - 6
    cmbxPodgruppa.Width = frameFilters.Width - cmbxPodgruppa.Left - 6
    btnClose.Left = Me.Width - btnClose.Width - 10
    tbtnFiltr.Left = Me.Width - tbtnFiltr.Width - 10
    btnNabDel.Left = btnClose.Left - btnNabDel.Width - 10
    btnFavDel.Left = btnNabDel.Left - btnFavDel.Width - 2
    btnETM.Left = btnFavDel.Left - btnETM.Width - 2
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
        cmbxProizvoditel.ListIndex = -1
        Reset_FiltersCmbx
    End If
    lblSostav.Caption = ""
    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
    lblSostav.Top = frameTab.Top + 222
    lstvTableNabor.Top = lblSostav.Top + 12
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
   If Not bBlock Then Find_ItemsByText
End Sub

Private Sub tbtnFav_Click()
    tbtnFav = True
End Sub

Private Sub tbtnBD_Click()
    If Not bBlock Then
        bBlock = True
        tbtnBD = False
        bBlock = False
        Me.Hide
        frmDBPriceAccess.Show
    End If
End Sub

Private Sub lblContent_Click() ' выровнять ширину столбцов по содержимому
   Dim colNum As Long
   For colNum = 0 To lstvTableIzbrannoe.ColumnHeaders.Count - 1
      Call SendMessage(lstvTableIzbrannoe.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvTableIzbrannoe.ColumnHeaders.Count - 1
      Call SendMessage(lstvTableIzbrannoe.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

Private Sub lstvTableIzbrannoe_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку
    With lstvTableIzbrannoe
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
End Sub

Sub btnClose_Click() ' выгрузка формы
    Unload frmDBPriceAccess
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

