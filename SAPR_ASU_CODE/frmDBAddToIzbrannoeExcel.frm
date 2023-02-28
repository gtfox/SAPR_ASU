Sub UserForm_Initialize()

    InitCustomCCPMenu Me 'Контекстное меню для TextBox

    FillExcel_cmbxProizvoditel cmbxProizvoditel
'    cmbxProizvoditel.style = fmStyleDropDownList
'    cmbxKategoriya.style = fmStyleDropDownList
'    cmbxGruppa.style = fmStyleDropDownList
'    cmbxPodgruppa.style = fmStyleDropDownList
    cmbxEdinicy.style = fmStyleDropDownList
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

    Reset_FiltersCmbx_ADO
'    InitCustomCCPMenu frmDBAddToIzbrannoeExcel 'Контекстное меню для TextBox
    frmDBAddToIzbrannoeExcel.Show
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
    bBlock = False
End Sub

Private Sub btnAdd_Click()
    Dim EstProizvoditel As Boolean
    'Запись данных в лист Избранное
    InitIzbrannoeExcelDB
    ClearFilter wshIzbrannoe
    wshIzbrannoe.Activate
    lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
    wshIzbrannoe.Cells(lLastRow + 1, 1) = txtArtikul.Value
    wshIzbrannoe.Cells(lLastRow + 1, 2) = txtNazvanie.Value
    wshIzbrannoe.Cells(lLastRow + 1, 3) = CDbl(txtCena.Value)
    wshIzbrannoe.Cells(lLastRow + 1, 4) = cmbxEdinicy '.List(cmbxEdinicy.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 5) = cmbxProizvoditel '.List(cmbxProizvoditel.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 6) = IIf(cmbxKategoriya = "", "Нет категории", cmbxKategoriya) '.List(cmbxKategoriya.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 7) = IIf(cmbxGruppa = "", "Нет группы", cmbxGruppa) '.List(cmbxGruppa.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 8) = IIf(cmbxPodgruppa = "", "Нет подгруппы", cmbxPodgruppa) '.List(cmbxPodgruppa.ListIndex, 0)
    
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
    
    wbExcelIzbrannoe.Save
    
    'Обновляем cmbxProizvoditel на случай, если был добавлен/удален производитель
    FillExcel_mProizvoditel
    FillExcel_cmbxProizvoditel frmDBIzbrannoeExcel.cmbxProizvoditel
    ExcelAppQuit oExcelAppIzbrannoe
    KillSAExcelProcess
                
    Unload Me
    frmDBIzbrannoeExcel.txtArtikul.Value = txtArtikul.Value
    frmDBIzbrannoeExcel.Find_ItemsByText_ADO
    frmDBIzbrannoeExcel.txtArtikul.Value = ""
    frmDBIzbrannoeExcel.lstvTableNabor.ListItems.Clear
    frmDBIzbrannoeExcel.Height = frmDBIzbrannoeExcel.frameTab.Top + frmDBIzbrannoeExcel.frameTab.Height + 36
    frmDBIzbrannoeExcel.lblSostav.Caption = ""
'    InitCustomCCPMenu frmDBIzbrannoeExcel 'Контекстное меню для TextBox
    frmDBIzbrannoeExcel.Show
End Sub

Private Sub txtCena_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 And KeyAscii <> 46) Then KeyAscii = 0
End Sub

'Private Sub cmbxPodgruppa_Change()
'    If cmbxPodgruppa = "Наборы" Then
'        MsgBox "Подгруппа ""Наборы"" используется только для наборов", vbOKOnly + vbInformation, "САПР-АСУ: Информация"
'        cmbxPodgruppa.ListIndex = 0
'    End If
'End Sub

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

Private Sub btnClose_Click()
    Unload Me
'    InitCustomCCPMenu frmDBPriceExcel 'Контекстное меню для TextBox
    frmDBPriceExcel.Show
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub