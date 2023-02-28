#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#End If
Private Const LVM_FIRST As Long = &H1000   ' 4096
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)   ' 4126
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Sub UserForm_Initialize()

    InitCustomCCPMenu Me 'Контекстное меню для TextBox

    lstvTableNabor.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvTableNabor.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTableNabor.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTableNabor.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
    lstvTableNabor.ColumnHeaders.Add , , "Ед." ' SubItems(3)
    lstvTableNabor.ColumnHeaders.Add , , "Производитель" ' SubItems(4)
    lstvTableNabor.ColumnHeaders.Add , , "Кол-во" ' SubItems(5)
    lstvTableNabor.ColumnHeaders.Add , , "    " ' SubItems(6)

'    cmbxProizvoditel.style = fmStyleDropDownList
    cmbxNabor.style = fmStyleDropDownList
    cmbxEdinicy.style = fmStyleDropDownList

    FillExcel_cmbxProizvoditel cmbxProizvoditel

End Sub

Sub run(Artikul As String, Nazvanie As String, Cena As String, Proizvoditel As String, Edinica As String)
    Dim SQLQuery As String
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie
    txtCena.Value = Cena
    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 0) = Proizvoditel Then cmbxProizvoditel.ListIndex = i
    Next

    SQLQuery = "SELECT ЕдиницыИзмерения FROM [" & ExcelEdinicyIzmereniya & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxEdinicy
    
    For i = 0 To cmbxEdinicy.ListCount - 1
        If cmbxEdinicy.List(i, 0) = Edinica Then cmbxEdinicy.ListIndex = i
    Next

    SQLQuery = "SELECT DISTINCT Набор FROM [" & ExcelNabory & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, cmbxNabor

'    InitCustomCCPMenu frmDBAddToNaborExcel 'Контекстное меню для TextBox
    frmDBAddToNaborExcel.Show
End Sub

'Public Sub FillCmbxNabor(cmbxComboBox As ComboBox)
'    Dim UserRange As Excel.Range
'    Dim lLastRow As Long
'    Dim i As Integer
'    Dim wshTemp As Excel.Worksheet
'
'    Set wshTemp = wbExcelIzbrannoe.Worksheets(ExcelTemp)
'    wshTemp.Cells.ClearContents
'    lLastRow = wshNabory.Cells(wshNabory.Rows.Count, 7).End(xlUp).Row
'    If lLastRow > 1 Then
'        wshNabory.Range("G2:G" & lLastRow).Copy wshTemp.Cells(1, 1)
'        Set UserRange = wshTemp.Range("A1:A" & lLastRow - 1)
'        UserRange.RemoveDuplicates Columns:=1, Header:=xlNo
'        lLastRow = wshTemp.Cells(wshTemp.Rows.Count, 1).End(xlUp).Row
'        If lLastRow > 0 Then
'            cmbxComboBox.Clear
'            For i = 1 To lLastRow
'                cmbxComboBox.AddItem wshTemp.Cells(i, 1)
'            Next
'        End If
'    Else
'        cmbxComboBox.Clear
'    End If
'    Set wshTemp = Nothing
'End Sub

Private Sub btnAdd_Click()
    Dim EstProizvoditel As Boolean
    Dim NewCena As Double
    Dim UserRange As Excel.Range
    InitIzbrannoeExcelDB
'    If cmbxNabor.ListIndex = -1 Then Exit Sub
    wshNabory.Activate
    ClearFilter wshNabory
    ClearFilter wshIzbrannoe
    lLastRow = wshNabory.Cells(wshNabory.Rows.Count, 1).End(xlUp).Row
    wshNabory.Cells(lLastRow + 1, 1) = txtArtikul.Value
    wshNabory.Cells(lLastRow + 1, 2) = txtNazvanie.Value
    wshNabory.Cells(lLastRow + 1, 3) = CDbl(txtCena.Value)
    wshNabory.Cells(lLastRow + 1, 4) = cmbxEdinicy
    wshNabory.Cells(lLastRow + 1, 5) = cmbxProizvoditel
    wshNabory.Cells(lLastRow + 1, 6) = CDbl(txtKolichestvo.Value)
    wshNabory.Cells(lLastRow + 1, 7) = cmbxNabor

    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 0) = cmbxProizvoditel Then EstProizvoditel = True
    Next
    'Добавляем производитля в базу
    If Not EstProizvoditel Then
        wshNastrojkiPrajsov.Activate
        lLastRow = wshNastrojkiPrajsov.Cells(wshNastrojkiPrajsov.Rows.Count, 1).End(xlUp).Row
        wshNastrojkiPrajsov.Cells(lLastRow + 1, 1) = cmbxProizvoditel
        wbExcelIzbrannoe.Save
    End If
    
    NewCena = CalcCenaNabora(lstvTableNabor) + CDbl(txtCena.Value) * CInt(txtKolichestvo.Value)

    Set UserRange = wshIzbrannoe.Columns(1).Find(What:=cmbxNabor, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If (UserRange Is Nothing) Or (UserRange.Value = Empty) Then
        MsgBox "Набор не найден в избранном" & vbCrLf & vbCrLf & "Набор: " & cmbxNabor, vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
    Else
        wshIzbrannoe.Cells(UserRange.Row, 3) = NewCena
    End If
    
    wbExcelIzbrannoe.Save

    'Обновляем cmbxProizvoditel на случай, если был добавлен/удален производитель
    FillExcel_mProizvoditel
    FillExcel_cmbxProizvoditel frmDBIzbrannoeExcel.cmbxProizvoditel
    ExcelAppQuit oExcelAppIzbrannoe
    KillSAExcelProcess
    
    Unload Me
    frmDBIzbrannoeExcel.txtArtikul.Value = cmbxNabor
    frmDBIzbrannoeExcel.Find_ItemsByText_ADO
    frmDBIzbrannoeExcel.txtArtikul.Value = ""
    frmDBIzbrannoeExcel.lstvTableNabor.ListItems.Clear
    frmDBIzbrannoeExcel.Height = frmDBIzbrannoeExcel.frameTab.Top + frmDBIzbrannoeExcel.frameTab.Height + 36
    frmDBIzbrannoeExcel.lblSostav.Caption = ""
'    InitCustomCCPMenu frmDBIzbrannoeExcel 'Контекстное меню для TextBox
    frmDBIzbrannoeExcel.Show
End Sub

Private Sub cmbxNabor_Change()
    Load_lstvTableNabor
End Sub

Sub Load_lstvTableNabor()
    Dim SQLQuery As String
    Dim colNum As Long
    
    If cmbxNabor.ListIndex > -1 Then
        SQLQuery = "SELECT * FROM [" & ExcelNabory & "$]  WHERE Набор='" & cmbxNabor & "';"
        lblSostav.Caption = "Состав набора: " & Fill_lstvTable_ADO(IzbrannoeSettings.FileName, SQLQuery, lstvTableNabor, 2)
    End If
    'выровнять ширину столбцов по заголовкам
    For colNum = 0 To lstvTableNabor.ColumnHeaders.Count - 1
        Call SendMessage(lstvTableNabor.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Private Sub txtCena_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 And KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub txtKolichestvo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 And KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub CommandButton5_Click()
    Dim UserRange As Excel.Range
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel & vbCrLf & vbCrLf & "Из избранного будут удалены все товары этого производителя", vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Производителей") = vbYes Then
        If cmbxProizvoditel <> "" Then
            InitIzbrannoeExcelDB
            
            Do  'Чистим избранное от записей удаляемого производителя
                Set UserRange = wshIzbrannoe.Columns(5).Find(What:=cmbxProizvoditel, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
                If Not UserRange Is Nothing Then
                    UserRange.EntireRow.Delete
                End If
            Loop While Not UserRange Is Nothing
            
            Set UserRange = wshNastrojkiPrajsov.Columns(1).Find(What:=cmbxProizvoditel, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If (UserRange Is Nothing) Or (UserRange.Value = Empty) Then
                MsgBox "Производитель не найден в базе" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel, vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
            Else
                UserRange.EntireRow.Delete
                wbExcelIzbrannoe.Save
                FillExcel_mProizvoditel
                ExcelAppQuit oExcelAppIzbrannoe
                KillSAExcelProcess
            End If
        End If
        UserForm_Initialize
    End If
End Sub

Private Sub CommandButton8_Click()
    Me.Hide
    Load frmDBAddNaborExcel
    frmDBAddNaborExcel.run txtArtikul.Value, txtNazvanie.Value, cmbxProizvoditel
End Sub

Private Sub btnClose_Click()
    Unload Me
'    InitCustomCCPMenu frmDBPriceExcel 'Контекстное меню для TextBox
    frmDBPriceExcel.Show
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub