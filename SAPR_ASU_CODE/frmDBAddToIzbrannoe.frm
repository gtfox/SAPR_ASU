Dim glShape As Visio.Shape



Sub UserForm_Initialize()
    Dim SQLQuery As String

    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель, Производители.КодПроизводителя " & _
                "FROM Производители;"
                
    Fill_cmbxProizvoditel "SAPR_ASU_Izbrannoe.accdb", SQLQuery, cmbxProizvoditel
    
    cmbxProizvoditel.style = fmStyleDropDownList
    cmbxKategoriya.style = fmStyleDropDownList
    cmbxGruppa.style = fmStyleDropDownList
    cmbxPodgruppa.style = fmStyleDropDownList
End Sub

Sub run(vsoShape As Visio.Shape, Artikul As String, Nazvanie As String, Cena As String, ProizvoditelID As String)
    Set glShape = vsoShape
    textArtikul.Value = Artikul
    textNazvanie.Value = Nazvanie
    textCena.Value = Cena
    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 2) = ProizvoditelID Then cmbxProizvoditel.ListIndex = i
    Next
    Reset_FiltersCmbx
    frmDBAddToIzbrannoe.Show
End Sub

Private Sub btnAdd_Click()
    Dim vsoShape As Visio.Shape
    Dim DBName As String
    Dim SQLQuery As String
    Set vsoShape = glShape
    DBName = "SAPR_ASU_Izbrannoe.accdb"
    SQLQuery = "INSERT INTO Избранное ( Артикул, Название, Цена, КатегорииКод, ГруппыКод, ПодгруппыКод, ПроизводительКод ) " & _
                "SELECT """ & textArtikul.Value & """, """ & textNazvanie.Value & """, """ & textCena.Value & """, " & cmbxKategoriya.ListIndex + 1 & ", " & cmbxGruppa.ListIndex + 1 & ", " & cmbxPodgruppa.ListIndex + 1 & " ," & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ";"
    ExecuteSQL DBName, SQLQuery
    Unload Me
    frmDBIzbrannoe.txtArtikul.Value = textArtikul.Value
    frmDBIzbrannoe.Find_ItemsByText
    frmDBIzbrannoe.txtArtikul.Value = ""
    frmDBIzbrannoe.run vsoShape
End Sub

Sub Reset_FiltersCmbx()
    Dim DBName As String
    Dim SQLQuery As String

    DBName = "SAPR_ASU_Izbrannoe.accdb"
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

Private Sub textCena_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 And KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub CommandButton1_Click()
    frmDBAddGroup.Caption = "Добавить производителя"
    frmDBAddGroup.lblName = "Имя производителя:"
    frmDBAddGroup.lblFile.Visible = True
    frmDBAddGroup.textFile.Visible = True
    frmDBAddGroup.btnAdd.Top = 48
    frmDBAddGroup.btnClose.Top = 48
    frmDBAddGroup.Height = 94
    frmDBAddGroup.run 1
End Sub

Private Sub CommandButton2_Click()
    frmDBAddGroup.Caption = "Добавить категорию"
    frmDBAddGroup.lblName = "Имя категории:"
    frmDBAddGroup.lblFile.Visible = False
    frmDBAddGroup.textFile.Visible = False
    frmDBAddGroup.btnAdd.Top = 30
    frmDBAddGroup.btnClose.Top = 30
    frmDBAddGroup.Height = 76
    frmDBAddGroup.run 2
End Sub

Private Sub CommandButton3_Click()
    frmDBAddGroup.Caption = "Добавить группу"
    frmDBAddGroup.lblName = "Имя группы:"
    frmDBAddGroup.lblFile.Visible = False
    frmDBAddGroup.textFile.Visible = False
    frmDBAddGroup.btnAdd.Top = 30
    frmDBAddGroup.btnClose.Top = 30
    frmDBAddGroup.Height = 76
    frmDBAddGroup.run 3
End Sub

Private Sub CommandButton4_Click()
    frmDBAddGroup.Caption = "Добавить подгруппу"
    frmDBAddGroup.lblName = "Имя подгруппы:"
    frmDBAddGroup.lblFile.Visible = False
    frmDBAddGroup.textFile.Visible = False
    frmDBAddGroup.btnAdd.Top = 30
    frmDBAddGroup.btnClose.Top = 30
    frmDBAddGroup.Height = 76
    frmDBAddGroup.run 4
End Sub

Private Sub CommandButton5_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0), vbYesNo + vbCritical, "Удаление записи из Производителей") = vbYes Then
        DBName = "SAPR_ASU_Izbrannoe.accdb"
        SQLQuery = "DELETE Производители.* " & _
                    "FROM Производители " & _
                    "WHERE Производители.КодПроизводителя=" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ";"
        ExecuteSQL DBName, SQLQuery
        UserForm_Initialize
    End If
End Sub

Private Sub CommandButton6_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Категория: " & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 0), vbYesNo + vbCritical, "Удаление записи из Категорий") = vbYes Then
        DBName = "SAPR_ASU_Izbrannoe.accdb"
        SQLQuery = "DELETE Категории.* " & _
                    "FROM Категории " & _
                    "WHERE Категории.КодКатегории=" & cmbxKategoriya.List(cmbxKategoriya.ListIndex, 1) & ";"
        ExecuteSQL DBName, SQLQuery
        Reset_FiltersCmbx
    End If
End Sub

Private Sub CommandButton7_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Группа: " & cmbxGruppa.List(cmbxGruppa.ListIndex, 0), vbYesNo + vbCritical, "Удаление записи из Групп") = vbYes Then
        DBName = "SAPR_ASU_Izbrannoe.accdb"
        SQLQuery = "DELETE Группы.* " & _
                    "FROM Группы " & _
                    "WHERE Группы.КодГруппы=" & cmbxGruppa.List(cmbxGruppa.ListIndex, 1) & ";"
        ExecuteSQL DBName, SQLQuery
        Reset_FiltersCmbx
    End If
End Sub

Private Sub CommandButton8_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Подгруппа: " & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 0), vbYesNo + vbCritical, "Удаление записи из Подгрупп") = vbYes Then
        DBName = "SAPR_ASU_Izbrannoe.accdb"
        SQLQuery = "DELETE Подгруппы.* " & _
                    "FROM Подгруппы " & _
                    "WHERE Подгруппы.КодПодгруппы=" & cmbxPodgruppa.List(cmbxPodgruppa.ListIndex, 1) & ";"
        ExecuteSQL DBName, SQLQuery
        Reset_FiltersCmbx
    End If
End Sub

Private Sub btnClose_Click()
Unload Me
frmDBIzbrannoe.run glShape
End Sub