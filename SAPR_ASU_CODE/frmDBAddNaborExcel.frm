Sub UserForm_Initialize()
    InitCustomCCPMenu Me 'Контекстное меню для TextBox
    FillExcel_cmbxProizvoditel cmbxProizvoditel
End Sub

Sub run(Artikul As String, Nazvanie As String, Proizvoditel As String)
    Dim SQLQuery As String
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie

    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 0) = Proizvoditel Then cmbxProizvoditel.ListIndex = i
    Next
    
    Reset_FiltersCmbx_ADO
'    InitCustomCCPMenu frmDBAddNaborExcel 'Контекстное меню для TextBox
    frmDBAddNaborExcel.Show
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
    Dim SQLQuery As String
    Dim EstProizvoditel As Boolean
    'Запись данных в лист Избранное
    InitIzbrannoeExcelDB
    ClearFilter wshIzbrannoe
    wshIzbrannoe.Activate
    lLastRow = wshIzbrannoe.Cells(wshIzbrannoe.Rows.Count, 1).End(xlUp).Row
    wshIzbrannoe.Cells(lLastRow + 1, 1) = "Набор_" & txtArtikul.Value
    wshIzbrannoe.Cells(lLastRow + 1, 2) = txtNazvanie.Value
    wshIzbrannoe.Cells(lLastRow + 1, 3) = 0
    wshIzbrannoe.Cells(lLastRow + 1, 4) = "шт."
    wshIzbrannoe.Cells(lLastRow + 1, 5) = cmbxProizvoditel
    wshIzbrannoe.Cells(lLastRow + 1, 6) = IIf(cmbxKategoriya = "", "Нет категории", cmbxKategoriya)
    wshIzbrannoe.Cells(lLastRow + 1, 7) = IIf(cmbxGruppa = "", "Нет группы", cmbxGruppa)
    wshIzbrannoe.Cells(lLastRow + 1, 8) = IIf(cmbxPodgruppa = "", "Нет подгруппы", cmbxPodgruppa)
    wshNabory.Range("G" & wshNabory.Cells(wshNabory.Rows.Count, 7).End(xlUp).Row + 1) = "Набор_" & txtArtikul.Value
    
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
    SQLQuery = "SELECT DISTINCT Набор FROM [" & ExcelNabory & "$];"
    Fill_ComboBox_ADO IzbrannoeSettings.FileName, SQLQuery, frmDBAddToNaborExcel.cmbxNabor
    frmDBAddToNaborExcel.Show
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

Private Sub btnClose_Click()
    Unload Me
'    InitCustomCCPMenu frmDBAddToNaborExcel 'Контекстное меню для TextBox
    frmDBAddToNaborExcel.Show
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub