




Sub UserForm_Initialize()
    Dim SQLQuery As String

    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель, Производители.КодПроизводителя " & _
                "FROM Производители;"
                
    Fill_cmbxProizvoditel DBNameIzbrannoe, SQLQuery, cmbxProizvoditel
    
    cmbxProizvoditel.style = fmStyleDropDownList
    cmbxKategoriya.style = fmStyleDropDownList
    cmbxGruppa.style = fmStyleDropDownList

End Sub

Sub run(Artikul As String, Nazvanie As String, Cena As String, ProizvoditelID As String)
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie

    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 2) = ProizvoditelID Then cmbxProizvoditel.ListIndex = i
    Next
    Reset_FiltersCmbx
    frmDBAddNabor.Show
End Sub

Private Sub btnAdd_Click()
    Dim DBName As String
    Dim SQLQuery As String
    DBName = DBNameIzbrannoe
    SQLQuery = "INSERT INTO Избранное ( Артикул, Название, Цена, КатегорииКод, ГруппыКод, ПодгруппыКод, ПроизводительКод, ЕдиницыКод ) " & _
                "SELECT ""Набор_" & txtArtikul.Value & """, """ & txtNazvanie.Value & """, """ & "0" & """, " & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1) & ", " & cmbxGruppa.List(cmbxGruppa.ListIndex, 1) & ", " & "2" & " ," & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ", " & "1" & ";"
    ExecuteSQL DBName, SQLQuery
    Unload Me
    frmDBAddToNabor.Reload_cmbxNabor
    frmDBAddToNabor.Show
End Sub

Sub Reset_FiltersCmbx()
    Dim DBName As String
    Dim SQLQuery As String

    DBName = DBNameIzbrannoe
    SQLQuery = "SELECT Категории.КодКатегории, Категории.Категория " & _
                "FROM Категории;"
    Fill_ComboBox DBName, SQLQuery, cmbxKategoriya
    SQLQuery = "SELECT Группы.КодГруппы, Группы.Группа " & _
                "FROM Группы;"
    Fill_ComboBox DBName, SQLQuery, cmbxGruppa
    cmbxKategoriya.ListIndex = 0
    cmbxGruppa.ListIndex = 0
End Sub

Private Sub CommandButton1_Click()
    frmDBAddGroup.Caption = "Создать производителя"
    frmDBAddGroup.lblName = "Имя производителя:"
    frmDBAddGroup.chbxAddFile.Visible = True
    
    frmDBAddGroup.run 5
End Sub

Private Sub CommandButton2_Click()
    frmDBAddGroup.Caption = "Создать категорию"
    frmDBAddGroup.lblName = "Имя категории:"
    frmDBAddGroup.chbxAddFile.Visible = False
    frmDBAddGroup.run 6
End Sub

Private Sub CommandButton3_Click()
    frmDBAddGroup.Caption = "Создать группу"
    frmDBAddGroup.lblName = "Имя группы:"
    frmDBAddGroup.chbxAddFile.Visible = False
    frmDBAddGroup.run 7
End Sub

Private Sub CommandButton5_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Производителей") = vbYes Then
        DBName = DBNameIzbrannoe
        SQLQuery = "DELETE Производители.* " & _
                    "FROM Производители " & _
                    "WHERE Производители.КодПроизводителя=" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ";"
        If Not (cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1) <> "") Then
            ExecuteSQL DBName, SQLQuery
        End If
        UserForm_Initialize
    End If
End Sub

Private Sub CommandButton6_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Категория: " & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 0), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Категорий") = vbYes Then
        DBName = DBNameIzbrannoe
        SQLQuery = "DELETE Категории.* " & _
                    "FROM Категории " & _
                    "WHERE Категории.КодКатегории=" & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1) & ";"
        If cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1) > 1 Then
            ExecuteSQL DBName, SQLQuery
        End If
        Reset_FiltersCmbx
    End If
End Sub

Private Sub CommandButton7_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Группа: " & cmbxGruppa.List(cmbxGruppa.ListIndex, 0), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Групп") = vbYes Then
        DBName = DBNameIzbrannoe
        SQLQuery = "DELETE Группы.* " & _
                    "FROM Группы " & _
                    "WHERE Группы.КодГруппы=" & cmbxGruppa.List(cmbxGruppa.ListIndex, 1) & ";"
        If cmbxGruppa.List(cmbxGruppa.ListIndex, 1) > 1 Then
            ExecuteSQL DBName, SQLQuery
        End If
        Reset_FiltersCmbx
    End If
End Sub

Private Sub btnClose_Click()
Unload Me
frmDBAddToNabor.Show
End Sub