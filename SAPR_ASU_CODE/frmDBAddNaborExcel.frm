Sub UserForm_Initialize()
    InitCustomCCPMenu Me 'Контекстное меню для TextBox
    FillExcel_cmbxProizvoditel cmbxProizvoditel
End Sub

Sub run(Artikul As String, Nazvanie As String, ProizvoditelID As String)
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie

    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 0) = ProizvoditelID Then cmbxProizvoditel.ListIndex = i
    Next
    frmDBIzbrannoeExcel.ClearFilter wshIzbrannoe
    frmDBIzbrannoeExcel.ClearFilter wshNabory
    frmDBIzbrannoeExcel.UpdateCmbxFiltersIzbrannoe cmbxKategoriya, 1
    frmDBIzbrannoeExcel.UpdateCmbxFiltersIzbrannoe cmbxGruppa, 2
    frmDBIzbrannoeExcel.UpdateCmbxFiltersIzbrannoe cmbxPodgruppa, 3
    InitCustomCCPMenu frmDBAddNaborExcel 'Контекстное меню для TextBox
    frmDBAddNaborExcel.Show
End Sub

Private Sub btnAdd_Click()
    'Запись данных в лист Избранное
    wshIzbrannoe.Activate
    lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
    wshIzbrannoe.Cells(lLastRow + 1, 1) = "Набор_" & txtArtikul.Value
    wshIzbrannoe.Cells(lLastRow + 1, 2) = txtNazvanie.Value
    wshIzbrannoe.Cells(lLastRow + 1, 3) = 0
    wshIzbrannoe.Cells(lLastRow + 1, 4) = "шт."
    wshIzbrannoe.Cells(lLastRow + 1, 5) = cmbxProizvoditel
    wshIzbrannoe.Cells(lLastRow + 1, 6) = cmbxKategoriya
    wshIzbrannoe.Cells(lLastRow + 1, 7) = cmbxGruppa
    wshIzbrannoe.Cells(lLastRow + 1, 8) = cmbxPodgruppa
    wshNabory.Range("G" & wshNabory.Cells(wshNabory.Rows.Count, 7).End(xlUp).Row + 1) = "Набор_" & txtArtikul.Value
    wbExcelIzbrannoe.Save
    Unload Me
    frmDBAddToNaborExcel.FillCmbxNabor frmDBAddToNaborExcel.cmbxNabor
    frmDBAddToNaborExcel.Show
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
                FillExcel_mProizvoditel
            End If
        End If
        UserForm_Initialize
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
    InitCustomCCPMenu frmDBAddToNaborExcel 'Контекстное меню для TextBox
    frmDBAddToNaborExcel.Show
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub