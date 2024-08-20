
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
Dim mstrShpData(5) As String
Dim mstrVybPozVNabore(7) As String
Dim bInit As Boolean

Private Sub UserForm_Initialize() ' инициализация формы
    Dim SQLQuery As String
    
    If Not bInit Then
        lstvTableIzbrannoe.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
        lstvTableIzbrannoe.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
        lstvTableIzbrannoe.ColumnHeaders.Add , , "Название" ' SubItems(1)
        lstvTableIzbrannoe.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
        lstvTableIzbrannoe.ColumnHeaders.Add , , "Ед." ' SubItems(3)
        lstvTableIzbrannoe.ColumnHeaders.Add , , "Производитель" ' SubItems(4)
    '    lstvTableIzbrannoe.ColumnHeaders.Add , , "    " ' SubItems(5)
       
        lstvTableNabor.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
        lstvTableNabor.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
        lstvTableNabor.ColumnHeaders.Add , , "Название" ' SubItems(1)
        lstvTableNabor.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
        lstvTableNabor.ColumnHeaders.Add , , "Ед." ' SubItems(3)
        lstvTableNabor.ColumnHeaders.Add , , "Производитель" ' SubItems(4)
        lstvTableNabor.ColumnHeaders.Add , , "Кол-во" ' SubItems(5)
    '    lstvTableNabor.ColumnHeaders.Add , , "    " ' SubItems(6)
    
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
        
        
           
        bBlock = True
        SQLQuery = "SELECT DISTINCT Производитель FROM [" & ExcelIzbrannoe & "$];"
        Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxProizvoditel
        bBlock = False
        
'        FillExcel_cmbxProizvoditel cmbxProizvoditel
        
    '    ClearFilter wshIzbrannoe
    '    ClearFilter wshNabory
        bInit = True
    End If
    InitCustomCCPMenu Me 'Контекстное меню для TextBox
End Sub

Sub Reset_FiltersCmbx_ADO()
    Dim SQLQuery As String
    bBlock = True
    SQLQuery = "SELECT DISTINCT Категория FROM [" & ExcelIzbrannoe & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxKategoriya
    SQLQuery = "SELECT DISTINCT Группа FROM [" & ExcelIzbrannoe & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxGruppa
    SQLQuery = "SELECT DISTINCT Подгруппа FROM [" & ExcelIzbrannoe & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxPodgruppa
    SQLQuery = "SELECT DISTINCT Производитель FROM [" & ExcelIzbrannoe & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxProizvoditel
    bBlock = False
    lstvTableIzbrannoe.ListItems.Clear
    lblResult.Caption = "Найдено записей: 0"
End Sub

 Sub Filter_CmbxChange_ADO(Ncmbx As Integer)
    Dim SQLQuery As String
    Dim fltrWhere As String
    Dim fltrProizvoditel As String
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
    
    If cmbxProizvoditel.ListIndex = -1 Then
        fltrProizvoditel = ""
    Else
        fltrProizvoditel = "Производитель='" & cmbxProizvoditel & "'"
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
            fltrWhere = IIf(fltrProizvoditel = "", "", " WHERE " & fltrProizvoditel)
        Case 1
            fltrWhere = " WHERE " & fltrPodgruppa & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
            bCallUpdatecmbxKategoriya = True
            bCallUpdatecmbxGruppa = True
        Case 2
            fltrWhere = " WHERE " & fltrGruppa & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
        Case 3
            fltrWhere = " WHERE " & fltrGruppa & " AND " & fltrPodgruppa & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
            bCallUpdatecmbxKategoriya = True
        Case 4
            fltrWhere = " WHERE " & fltrKategoriya & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
        Case 5
            fltrWhere = " WHERE " & fltrKategoriya & " AND " & fltrPodgruppa & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
            bCallUpdatecmbxGruppa = True
        Case 6
            fltrWhere = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
        Case 7
            fltrWhere = " WHERE " & fltrKategoriya & " AND " & fltrGruppa & " AND " & fltrPodgruppa & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
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
            fltrWhere = " WHERE " & fltrKategoriya & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
            fltrGruppa = ""
            fltrPodgruppa = ""
            bBlock = True
            cmbxGruppa.Clear
            cmbxPodgruppa.Clear
            bBlock = False
            bCallUpdatecmbxGruppa = True
            bCallUpdatecmbxPodgruppa = True
        Case 2
            fltrWhere = IIf(fltrKategoriya = "", " WHERE " & fltrGruppa, " WHERE " & fltrKategoriya & " AND " & fltrGruppa) & IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
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
    SQLQuery = "SELECT * FROM [" & ExcelIzbrannoe & "$] " & fltrWhere & ";"
    lstvTableIzbrannoe.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable_ADO(IzbrannoeSettings.FileName, SQLQuery, lstvTableIzbrannoe, 1)
    lstvTableIzbrannoe.Visible = True
    Fill_FiltersByResultSQLQuery_ADO
    ReSize
End Sub

'Полнотекстовый поиск
Sub Find_ItemsByText_ADO()
    Dim SQLQuery As String
    Dim findMode As Integer
    Dim findWhat As String
    Dim findArtikul As String
    Dim findNazvanie As String
    Dim fltrWhere As String
    Dim fltrProizvoditel As String
    Dim fltrKategoriya As String
    Dim fltrGruppa As String
    Dim fltrPodgruppa As String

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
    If cmbxProizvoditel.ListIndex = -1 Then
        fltrProizvoditel = ""
    Else
        fltrProizvoditel = "Производитель='" & cmbxProizvoditel & "'"
    End If
    
    fltrWhere = IIf(fltrKategoriya = "", "", " AND " & fltrKategoriya) & _
                IIf(fltrGruppa = "", "", " AND " & fltrGruppa) & _
                IIf(fltrPodgruppa = "", "", " AND " & fltrPodgruppa) & _
                IIf(fltrProizvoditel = "", "", " AND " & fltrProizvoditel)
    SQLQuery = "SELECT * FROM [" & ExcelIzbrannoe & "$] " & findWhat & fltrWhere & ";"
    lstvTableIzbrannoe.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable_ADO(IzbrannoeSettings.FileName, SQLQuery, lstvTableIzbrannoe, 1)
    lstvTableIzbrannoe.Visible = True
    bCallUpdatecmbxKategoriya = True
    bCallUpdatecmbxGruppa = True
    bCallUpdatecmbxPodgruppa = True
    bCallUpdatecmbxProizvoditel = True
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
        Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxKategoriya
        For i = 0 To cmbxKategoriya.ListCount - 1
            If cmbxKategoriya.List(i, 0) = scmbxKategoriyaValue Then cmbxKategoriya.ListIndex = i
        Next
        bCallUpdatecmbxKategoriya = False
    End If
    If bCallUpdatecmbxGruppa Then
        scmbxGruppaValue = cmbxGruppa
        SQLQuery = "SELECT DISTINCT Группа FROM (" & sLastSQLQuery & ");"
        Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxGruppa
        For i = 0 To cmbxGruppa.ListCount - 1
            If cmbxGruppa.List(i, 0) = scmbxGruppaValue Then cmbxGruppa.ListIndex = i
        Next
        bCallUpdatecmbxGruppa = False
    End If
    If bCallUpdatecmbxPodgruppa Then
        scmbxPodgruppaValue = cmbxPodgruppa
        SQLQuery = "SELECT DISTINCT Подгруппа FROM (" & sLastSQLQuery & ");"
        Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxPodgruppa
        For i = 0 To cmbxPodgruppa.ListCount - 1
            If cmbxPodgruppa.List(i, 0) = scmbxPodgruppaValue Then cmbxPodgruppa.ListIndex = i
        Next
        bCallUpdatecmbxPodgruppa = False
    End If
    bBlock = False
End Sub

Private Sub btnFavDel_Click()
    Dim UserRange As Excel.Range
    InitIzbrannoeExcelDB
    If MsgBox("Удалить запись из избранного?" & vbCrLf & vbCrLf & "Артикул: " & mstrShpData(0) & vbCrLf & "Название: " & mstrShpData(1) & vbCrLf & "Цена: " & mstrShpData(2) & vbCrLf & "Производитель: " & mstrShpData(4), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Избранного") = vbYes Then
        Set UserRange = wshIzbrannoe.Columns(1).Find(What:=mstrShpData(0), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        UserRange.EntireRow.Delete
        lstvTableNabor.ListItems.Clear
        Do  'Чистим состав набора
            Set UserRange = wshNabory.Columns(7).Find(What:=mstrShpData(0), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If Not UserRange Is Nothing Then
                UserRange.EntireRow.Delete
            End If
        Loop While Not UserRange Is Nothing
        wbExcelIzbrannoe.Save
        ExcelAppQuit oExcelAppIzbrannoe
        KillSAExcelProcess
        Find_ItemsByText_ADO
    End If
End Sub

Private Sub btnNabDel_Click()
    Dim UserRange As Excel.Range
    Dim NewCena As Double
    Dim SQLQuery As String
    InitIzbrannoeExcelDB
    If MsgBox("Удалить запись из набора?" & vbCrLf & vbCrLf & "Артикул: " & mstrVybPozVNabore(0) & vbCrLf & "Название: " & mstrVybPozVNabore(1) & vbCrLf & "Цена: " & mstrVybPozVNabore(2) & vbCrLf & "Производитель: " & mstrVybPozVNabore(4), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Набора") = vbYes Then
        Set UserRange = wshNabory.Columns(1).Find(What:=mstrVybPozVNabore(0), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        UserRange.EntireRow.Delete
        wbExcelIzbrannoe.Save
        ExcelAppQuit oExcelAppIzbrannoe
        KillSAExcelProcess
        SQLQuery = "SELECT * FROM [" & ExcelNabory & "$]  WHERE Набор='" & mstrShpData(0) & "';"
        lblSostav.Caption = "Состав набора: " & Fill_lstvTable_ADO(IzbrannoeSettings.FileName, SQLQuery, lstvTableNabor, 2)
        NewCena = CalcCenaNabora(lstvTableNabor)
        InitIzbrannoeExcelDB
        Set UserRange = wshIzbrannoe.Columns(1).Find(What:=mstrShpData(0), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If (UserRange Is Nothing) Or (UserRange.Value = Empty) Then
            MsgBox "Набор не найден в избранном" & vbCrLf & vbCrLf & "Набор: " & cmbxNabor, vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
        Else
            wshIzbrannoe.Cells(UserRange.Row, 3) = NewCena
        End If
        lstvTableNabor.Width = frmMinWdth
        'выровнять ширину столбцов по заголовкам
        For colNum = 0 To lstvTableNabor.ColumnHeaders.Count - 1
            Call SendMessage(lstvTableNabor.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
        Next
        Me.Height = lstvTableNabor.Top + lstvTableNabor.Height + 26
        wbExcelIzbrannoe.Save
        ExcelAppQuit oExcelAppIzbrannoe
        KillSAExcelProcess
        Find_ItemsByText_ADO
    End If
End Sub

Private Sub lstvTableIzbrannoe_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Если в таблице ткнуть на строку с номером больше 30000 то сюда попадет первая строка!!!
    Dim colNum As Long
    Dim SQLQuery As String
    
    mstrShpData(0) = Item
    mstrShpData(1) = Item.SubItems(1)
    mstrShpData(2) = Item.SubItems(2)
    mstrShpData(3) = Item.SubItems(3)
    mstrShpData(4) = Item.SubItems(4)

    If Item.ForeColor = NaboryColor Then
        SQLQuery = "SELECT * FROM [" & ExcelNabory & "$]  WHERE Набор='" & Item & "';"
        lblSostav.Caption = "Состав набора: " & Fill_lstvTable_ADO(IzbrannoeSettings.FileName, SQLQuery, lstvTableNabor, 2)
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
    
    Set vsoShape = frmDBPriceExcel.glShape
    GoSub SetDB
    
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> frmDBPriceExcel.glShape And ShapeSATypeIs(vsoShape, ShapeSAType(frmDBPriceExcel.glShape)) Then
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

Private Sub lstvTableNabor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mstrVybPozVNabore(0) = Item
    mstrVybPozVNabore(1) = Item.SubItems(1)
    mstrVybPozVNabore(2) = Item.SubItems(2)
    mstrVybPozVNabore(3) = Item.SubItems(3)
    mstrVybPozVNabore(4) = Item.SubItems(4)
    mstrVybPozVNabore(5) = Item.SubItems(5)
End Sub

Private Sub lstvTableNabor_DblClick()
    Dim vsoShape As Visio.Shape
    
    Set vsoShape = frmDBPriceExcel.glShape
    GoSub SetDB
  
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> frmDBPriceExcel.glShape And ShapeSATypeIs(vsoShape, ShapeSAType(frmDBPriceExcel.glShape)) Then
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
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrVybPozVNabore(1), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrVybPozVNabore(0) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrVybPozVNabore(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrVybPozVNabore(2) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrVybPozVNabore(3) & """"
    End With
    err.Clear
    On Error GoTo 0
    Return
    
errGuard:
    With vsoShape
        .Cells("Prop.NazvanieDB").FormulaForce = """" & Replace(mstrVybPozVNabore(1), """", """""") & """"
        .Cells("Prop.NazvanieDB.Type").FormulaForce = 0
        .Cells("Prop.NazvanieDB.Format").FormulaForce = """"""
        .Cells("Prop.ArtikulDB").FormulaForce = """" & mstrVybPozVNabore(0) & """"
        .Cells("Prop.ArtikulDB.Type").FormulaForce = 0
        .Cells("Prop.ArtikulDB.Format").FormulaForce = """"""
        .Cells("Prop.ProizvoditelDB").FormulaForce = """" & mstrVybPozVNabore(4) & """"
        .Cells("Prop.ProizvoditelDB.Type").FormulaForce = 0
        .Cells("Prop.ProizvoditelDB.Format").FormulaForce = """"""
        .Cells("Prop.CenaDB").FormulaForce = """" & mstrVybPozVNabore(2) & """"
        .Cells("Prop.CenaDB.Type").FormulaForce = 0
        .Cells("Prop.CenaDB.Format").FormulaForce = """"""
        .Cells("Prop.EdDB").FormulaForce = """" & mstrVybPozVNabore(3) & """"
        .Cells("Prop.EdDB.Type").FormulaForce = 0
        .Cells("Prop.EdDB.Format").FormulaForce = """"""
    End With
    Return
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
    Dim RangeToFilter As Excel.Range
    Dim lLastRow As Long

    If tbtnFiltr.Value Then
        frameFilters.Height = 84
        tbtnFiltr.Caption = ChrW(9650) 'вверх
    Else
        Reset_FiltersCmbx_ADO
        frameFilters.Height = 0
        tbtnFiltr.Caption = ChrW(9660) 'вниз
        cmbxProizvoditel.ListIndex = -1
        bBlock = True
        cmbxKategoriya.ListIndex = -1
        cmbxGruppa.ListIndex = -1
        cmbxPodgruppa.ListIndex = -1
        bBlock = False
        txtNazvanie2.Value = ""
        txtArtikul.Value = ""
        Find_ItemsByText_ADO
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
    If mstrShpData(0) Like "Набор_*" Then
        FindArticulInBrowser mstrVybPozVNabore(0), cmbxMagazin.ListIndex
    Else
        FindArticulInBrowser mstrShpData(0), cmbxMagazin.ListIndex
    End If
End Sub

Private Sub btnAVS_Click()
    If mstrShpData(0) Like "Набор_*" Then
        FindArticulInBrowser mstrVybPozVNabore(0), cmbxMagazin.ListIndex
    Else
        FindArticulInBrowser mstrShpData(0), cmbxMagazin.ListIndex
    End If
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

Private Sub cmbxProizvoditel_Change()
    If Not bBlock Then Filter_CmbxChange_ADO 3
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
        frmDBPriceExcel.Show
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
    ExcelAppQuit oExcelAppIzbrannoe
    ExcelAppQuit oExcelAppPrice
    KillSAExcelProcess
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload frmDBPriceExcel
    Unload Me
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub
