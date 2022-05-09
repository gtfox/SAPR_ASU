




Sub UserForm_Initialize()
    Dim SQLQuery As String

    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель, Производители.КодПроизводителя " & _
                "FROM Производители;"
                
    Fill_cmbxProizvoditel DBNameIzbrannoe, SQLQuery, cmbxProizvoditel
    

    
    cmbxProizvoditel.style = fmStyleDropDownList
    cmbxKategoriya.style = fmStyleDropDownList
    cmbxGruppa.style = fmStyleDropDownList
    cmbxPodgruppa.style = fmStyleDropDownList
    cmbxEdinicy.style = fmStyleDropDownList
End Sub

Sub run(Artikul As String, Nazvanie As String, Cena As String, ProizvoditelID As String, EdinicaID As String)
    Dim SQLQuery As String
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie
    txtCena.Value = Cena
    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 2) = ProizvoditelID Then cmbxProizvoditel.ListIndex = i
    Next

    SQLQuery = "SELECT Единицы.КодЕдиницы, Единицы.Единица " & _
                "FROM Единицы;"

    Fill_ComboBox DBNameIzbrannoe, SQLQuery, cmbxEdinicy

    For i = 0 To cmbxEdinicy.ListCount - 1
        If cmbxEdinicy.List(i, 1) = EdinicaID Then cmbxEdinicy.ListIndex = i
    Next

    Reset_FiltersCmbx
    frmDBAddToIzbrannoe.Show
End Sub

Private Sub btnAdd_Click()
    Dim DBName As String
    Dim SQLQuery As String
    DBName = DBNameIzbrannoe
    SQLQuery = "INSERT INTO Избранное ( Артикул, Название, Цена, КатегорииКод, ГруппыКод, ПодгруппыКод, ПроизводительКод, ЕдиницыКод ) " & _
                "SELECT """ & txtArtikul.Value & """, """ & txtNazvanie.Value & """, """ & txtCena.Value & """, " & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1) & ", " & cmbxGruppa.List(cmbxGruppa.ListIndex, 1) & ", " & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1) & " ," & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ", " & cmbxEdinicy.List(cmbxEdinicy.ListIndex, 1) & ";"
    ExecuteSQL DBName, SQLQuery
    Unload Me
    frmDBIzbrannoe.txtArtikul.Value = txtArtikul.Value
    frmDBIzbrannoe.Find_ItemsByText
    frmDBIzbrannoe.txtArtikul.Value = ""
    frmDBIzbrannoe.lstvTableNabor.ListItems.Clear
    frmDBIzbrannoe.Height = frmDBIzbrannoe.frameTab.Top + frmDBIzbrannoe.frameTab.Height + 36
    frmDBIzbrannoe.lblSostav.Caption = ""
    frmDBIzbrannoe.Show
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
    SQLQuery = "SELECT Подгруппы.КодПодгруппы, Подгруппы.Подгруппа " & _
                "FROM Подгруппы;"
    Fill_ComboBox DBName, SQLQuery, cmbxPodgruppa
    cmbxKategoriya.ListIndex = 0
    cmbxGruppa.ListIndex = 0
    cmbxPodgruppa.ListIndex = 0
End Sub

Private Sub txtCena_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 And KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub cmbxPodgruppa_Change()
    If cmbxPodgruppa.ListIndex = 1 Then
        MsgBox "Подгруппа ""Наборы"" используется только для наборов", vbOKOnly + vbInformation, "САПР-АСУ: Информация"
        cmbxPodgruppa.ListIndex = 0
    End If
End Sub

Private Sub CommandButton1_Click()
    frmDBAddGroup.Caption = "Создать производителя"
    frmDBAddGroup.lblName = "Имя производителя:"
    frmDBAddGroup.chbxAddFile.Visible = True
    frmDBAddGroup.run 1
End Sub

Private Sub CommandButton2_Click()
    frmDBAddGroup.Caption = "Создать категорию"
    frmDBAddGroup.lblName = "Имя категории:"
    frmDBAddGroup.chbxAddFile.Visible = False
    frmDBAddGroup.run 2
End Sub

Private Sub CommandButton3_Click()
    frmDBAddGroup.Caption = "Создать группу"
    frmDBAddGroup.lblName = "Имя группы:"
    frmDBAddGroup.chbxAddFile.Visible = False
    frmDBAddGroup.run 3
End Sub

Private Sub CommandButton4_Click()
    frmDBAddGroup.Caption = "Создать подгруппу"
    frmDBAddGroup.lblName = "Имя подгруппы:"
    frmDBAddGroup.chbxAddFile.Visible = False
    frmDBAddGroup.run 4
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

Private Sub CommandButton8_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Подгруппа: " & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 0), vbYesNo + vbCritical, "САПР-АСУ: Удаление записи из Подгрупп") = vbYes Then
        DBName = DBNameIzbrannoe
        SQLQuery = "DELETE Подгруппы.* " & _
                    "FROM Подгруппы " & _
                    "WHERE Подгруппы.КодПодгруппы=" & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1) & ";"
        If cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1) > 2 Then
            ExecuteSQL DBName, SQLQuery
        End If
        Reset_FiltersCmbx
    End If
End Sub

Private Sub btnClose_Click()
Unload Me
frmDBPrice.Show
End Sub