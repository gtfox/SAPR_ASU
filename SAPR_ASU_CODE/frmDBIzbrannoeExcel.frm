
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

Private Sub UserForm_Initialize() ' инициализация формы

    InitCustomCCPMenu Me 'Контекстное меню для TextBox
    
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
    
'    Set oIzbrannoeRecordSet As New ADODB.Recordset'= CreateObject("ADODB.Recordset")
'    Set oIzbrannoeConn = CreateObject("ADODB.Connection")
    ADODB_Excel_Connect oIzbrannoeConn, sSAPath & DBNameIzbrannoeExcel
    
    FillExcel_cmbxProizvoditel cmbxProizvoditel
    
    ClearFilter wshIzbrannoe
    ClearFilter wshNabory

End Sub

Private Sub Filter_CmbxChange(Ncmbx As Integer)
    Dim RangeToFilter As Excel.Range
    Dim lLastRow As Long
    Dim fltrMode As Integer

    lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
    Set RangeToFilter = wshIzbrannoe.Range("A2:H" & lLastRow)
    
    'ФИЛЬТРАЦИЯ
    RuleFilterCmbx wshIzbrannoe, RangeToFilter, Me, IzbrannoeSettings, Ncmbx
    lstvTableIzbrannoe.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(oIzbrannoeRecordSet, oIzbrannoeConn, wshIzbrannoe, lstvTableIzbrannoe, IzbrannoeSettings, 1)
    lstvTableIzbrannoe.Visible = True
    ReSize

End Sub


'Полнотекстовый поиск
Sub Find_ItemsByText()
    Dim RangeToFilter As Excel.Range
    Dim lLastRow As Long

    lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
    Set RangeToFilter = wshIzbrannoe.Range("A2:H" & lLastRow)
    
    If txtArtikul.Value = "" Then
        RangeToFilter.AutoFilter Field:=IzbrannoeSettings.StolbArtikul
    Else
        RangeToFilter.AutoFilter Field:=IzbrannoeSettings.StolbArtikul, Criteria1:="=*" & txtArtikul.Value & "*"
    End If
    
    If txtNazvanie2.Value = "" Then
        RangeToFilter.AutoFilter Field:=IzbrannoeSettings.StolbNazvanie
    Else
        RangeToFilter.AutoFilter Field:=IzbrannoeSettings.StolbNazvanie, Criteria1:="=*" & Replace(txtNazvanie2.Value, " ", "*") & "*"
    End If
    lstvTableIzbrannoe.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(oIzbrannoeRecordSet, oIzbrannoeConn, wshIzbrannoe, lstvTableIzbrannoe, IzbrannoeSettings, 1)
    lstvTableIzbrannoe.Visible = True
    UpdateAllCmbxFilters wshIzbrannoe, Me, IzbrannoeSettings
    
    ReSize

End Sub

Private Sub btnFavDel_Click()
    Dim UserRange As Excel.Range
    If MsgBox("Удалить запись из избранного?" & vbCrLf & vbCrLf & "Артикул: " & mstrShpData(0) & vbCrLf & "Название: " & mstrShpData(1) & vbCrLf & "Цена: " & mstrShpData(2) & vbCrLf & "Производитель: " & mstrShpData(4), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Избранного") = vbYes Then
        Set UserRange = wshIzbrannoe.Columns(1).Find(What:=mstrShpData(0), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        UserRange.EntireRow.Delete
        lstvTableNabor.ListItems.Clear
        Find_ItemsByText
    End If
End Sub

Private Sub btnNabDel_Click()
    Dim UserRange As Excel.Range
    Dim NewCena As Double
    If MsgBox("Удалить запись из набора?" & vbCrLf & vbCrLf & "Артикул: " & mstrVybPozVNabore(0) & vbCrLf & "Название: " & mstrVybPozVNabore(1) & vbCrLf & "Цена: " & mstrVybPozVNabore(2) & vbCrLf & "Производитель: " & mstrVybPozVNabore(4), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Набора") = vbYes Then
        Set UserRange = wshNabory.Columns(1).Find(What:=mstrVybPozVNabore(0), LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        UserRange.EntireRow.Delete
        lblSostav.Caption = "Состав набора: " & Fill_lstvTable(oIzbrannoeRecordSet, oIzbrannoeConn, wshNabory, lstvTableNabor, IzbrannoeSettings, 2)
        NewCena = CalcCenaNabora(lstvTableNabor)
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
        Find_ItemsByText
    End If
End Sub

Private Sub lstvTableIzbrannoe_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Если в таблице ткнуть на строку с номером больше 30000 то сюда попадет первая строка!!!
    Dim colNum As Long
    Dim RangeToFilter As Excel.Range
    Dim lLastRow As Long

    mstrShpData(0) = Item
    mstrShpData(1) = Item.SubItems(1)
    mstrShpData(2) = Item.SubItems(2)
    mstrShpData(3) = Item.SubItems(3)
    mstrShpData(4) = Item.SubItems(4)

    If Item.ForeColor = NaboryColor Then
        lLastRow = wshNabory.Cells(wshNabory.Rows.Count, 1).End(xlUp).Row
        Set RangeToFilter = wshNabory.Range("A2:H" & lLastRow)
        RangeToFilter.AutoFilter Field:=7, Criteria1:=Item
        lblSostav.Caption = "Состав набора: " & Fill_lstvTable(oIzbrannoeRecordSet, oIzbrannoeConn, wshNabory, lstvTableNabor, IzbrannoeSettings, 2)
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
    
    With frmDBPriceExcel.glShape
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
            If vsoShape <> frmDBPriceExcel.glShape And ShapeSATypeIs(vsoShape, ShapeSAType(frmDBPriceExcel.glShape)) Then
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
    
    With frmDBPriceExcel.glShape
        .Cells("User.KodProizvoditelyaDB").Formula = """"""
        .Cells("User.KodPoziciiDB").Formula = """"""
        .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrVybPozVNabore(1), """", """""") & """"
        .Cells("Prop.ArtikulDB").Formula = """" & mstrVybPozVNabore(0) & """"
        .Cells("Prop.ProizvoditelDB").Formula = """" & mstrVybPozVNabore(4) & """"
        .Cells("Prop.CenaDB").Formula = """" & mstrVybPozVNabore(2) & """"
        .Cells("Prop.EdDB").Formula = """" & mstrVybPozVNabore(3) & """"
    End With
    
    If ActiveWindow.Selection.Count > 1 Then
        For Each vsoShape In ActiveWindow.Selection
            If vsoShape <> frmDBPriceExcel.glShape And ShapeSATypeIs(vsoShape, ShapeSAType(frmDBPriceExcel.glShape)) Then
                With vsoShape
                    .Cells("User.KodProizvoditelyaDB").Formula = """"""
                    .Cells("User.KodPoziciiDB").Formula = """"""
                    .Cells("Prop.NazvanieDB").Formula = """" & Replace(mstrVybPozVNabore(1), """", """""") & """"
                    .Cells("Prop.ArtikulDB").Formula = """" & mstrVybPozVNabore(0) & """"
                    .Cells("Prop.ProizvoditelDB").Formula = """" & mstrVybPozVNabore(4) & """"
                    .Cells("Prop.CenaDB").Formula = """" & mstrVybPozVNabore(2) & """"
                    .Cells("Prop.EdDB").Formula = """" & mstrVybPozVNabore(3) & """"
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
    Dim RangeToFilter As Excel.Range
    Dim lLastRow As Long

    If tbtnFiltr.Value Then
        frameFilters.Height = 84
        tbtnFiltr.Caption = ChrW(9650) 'вверх
    Else
        lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
        Set RangeToFilter = wshIzbrannoe.Range("A2:H" & lLastRow)
        frameFilters.Height = 0
        tbtnFiltr.Caption = ChrW(9660) 'вниз
        cmbxProizvoditel.ListIndex = -1
        ClearFilter wshIzbrannoe
        ClearFilter wshNabory
        Find_ItemsByText
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
    Dim RangeToFilter As Excel.Range
    Dim lLastRow As Long
    
    If Not bBlock Then
        lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
        Set RangeToFilter = wshIzbrannoe.Range("A2:H" & lLastRow)
        'Фильтр по Производителю (не обновляется по результатам фильтрации)
        If cmbxProizvoditel.ListIndex = -1 Then
            RangeToFilter.AutoFilter Field:=5
        Else
            RangeToFilter.AutoFilter Field:=5, Criteria1:=cmbxProizvoditel
        End If
    End If
    lstvTableIzbrannoe.Visible = False
    lblResult.Caption = "Найдено записей: " & Fill_lstvTable(oIzbrannoeRecordSet, oIzbrannoeConn, wshIzbrannoe, lstvTableIzbrannoe, IzbrannoeSettings, 1)
    lstvTableIzbrannoe.Visible = True
    ReSize
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
        InitCustomCCPMenu frmDBPriceExcel 'Контекстное меню для TextBox
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
    If oIzbrannoeRecordSet.State = adStateOpen Then oIzbrannoeRecordSet.Close
    oIzbrannoeConn.Close
    ExcelAppExit
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload frmDBPriceExcel
    Unload Me
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub
