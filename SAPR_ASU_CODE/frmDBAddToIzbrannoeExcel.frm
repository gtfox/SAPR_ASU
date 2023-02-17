Sub UserForm_Initialize()

    InitCustomCCPMenu Me 'Контекстное меню для TextBox

    FillExcel_cmbxProizvoditel cmbxProizvoditel
'    cmbxProizvoditel.style = fmStyleDropDownList
'    cmbxKategoriya.style = fmStyleDropDownList
'    cmbxGruppa.style = fmStyleDropDownList
'    cmbxPodgruppa.style = fmStyleDropDownList
    cmbxEdinicy.style = fmStyleDropDownList
End Sub

Sub run(Artikul As String, Nazvanie As String, Cena As String, ProizvoditelID As String, EdinicaID As String)
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie
    txtCena.Value = Cena
    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 0) = ProizvoditelID Then cmbxProizvoditel.ListIndex = i
    Next

    FillCmbxEdinicy cmbxEdinicy

    For i = 0 To cmbxEdinicy.ListCount - 1
        If cmbxEdinicy.List(i, 0) = EdinicaID Then cmbxEdinicy.ListIndex = i
    Next

    ClearFilter wshIzbrannoe
    ClearFilter wshNabory
    UpdateAllCmbxFilters wshIzbrannoe, frmDBAddToIzbrannoeExcel, IzbrannoeSettings
'    InitCustomCCPMenu frmDBAddToIzbrannoeExcel 'Контекстное меню для TextBox
    frmDBAddToIzbrannoeExcel.Show
End Sub

Private Sub btnAdd_Click()
    'Запись данных в лист Избранное
    wshIzbrannoe.Activate
    lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
    wshIzbrannoe.Cells(lLastRow + 1, 1) = txtArtikul.Value
    wshIzbrannoe.Cells(lLastRow + 1, 2) = txtNazvanie.Value
    wshIzbrannoe.Cells(lLastRow + 1, 3) = CDbl(txtCena.Value)
    wshIzbrannoe.Cells(lLastRow + 1, 4) = cmbxEdinicy '.List(cmbxEdinicy.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 5) = cmbxProizvoditel '.List(cmbxProizvoditel.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 6) = cmbxKategoriya '.List(cmbxKategoriya.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 7) = cmbxGruppa '.List(cmbxGruppa.ListIndex, 0)
    wshIzbrannoe.Cells(lLastRow + 1, 8) = cmbxPodgruppa '.List(cmbxPodgruppa.ListIndex, 0)
    wbExcelIzbrannoe.Save
    
    'Обновляем cmbxProizvoditel в 2-х формах на случай, если был добавлен/удален производитель
    FillExcel_mProizvoditel
    frmDBPriceExcel.FillExcel_cmbxProizvoditel frmDBPriceExcel.cmbxProizvoditel, True
    frmDBIzbrannoeExcel.FillExcel_cmbxProizvoditel frmDBIzbrannoeExcel.cmbxProizvoditel
    
    Unload Me
    frmDBIzbrannoeExcel.txtArtikul.Value = txtArtikul.Value
    frmDBIzbrannoeExcel.Find_ItemsByText
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

Private Sub cmbxPodgruppa_Change()
    If cmbxPodgruppa = "Наборы" Then
        MsgBox "Подгруппа ""Наборы"" используется только для наборов", vbOKOnly + vbInformation, "САПР-АСУ: Информация"
        cmbxPodgruppa.ListIndex = 0
    End If
End Sub

Private Sub CommandButton5_Click()
    Dim UserRange As Excel.Range
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel, vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Производителей") = vbYes Then
        If cmbxProizvoditel <> "" Then
            Set UserRange = wshNastrojkiPrajsov.Columns(1).Find(What:=cmbxProizvoditel, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If (UserRange Is Nothing) Or (UserRange.Value = Empty) Then
                MsgBox "Производитель не найден в базе" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel, vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
            Else
                UserRange.EntireRow.Delete
                wbExcelIzbrannoe.Save
                
                'Обновляем cmbxProizvoditel в 2-х формах на случай, если был добавлен/удален производитель
                FillExcel_mProizvoditel
                frmDBPriceExcel.FillExcel_cmbxProizvoditel frmDBPriceExcel.cmbxProizvoditel, True
                frmDBIzbrannoeExcel.FillExcel_cmbxProizvoditel frmDBIzbrannoeExcel.cmbxProizvoditel
                
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