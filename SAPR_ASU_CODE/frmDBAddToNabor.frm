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
    Dim SQLQuery As String
    
    lstvTableNabor.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvTableNabor.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTableNabor.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTableNabor.ColumnHeaders.Add , , "Цена", , lvwColumnRight ' SubItems(2)
'    lstvTableNabor.ColumnHeaders.Add , , "Единица" ' SubItems(3)
    lstvTableNabor.ColumnHeaders.Add , , "Производитель" ' SubItems(3)
    lstvTableNabor.ColumnHeaders.Add , , "Количество" ' SubItems(4)
    
    cmbxProizvoditel.style = fmStyleDropDownList
    cmbxNabor.style = fmStyleDropDownList
    cmbxEdinicy.style = fmStyleDropDownList
    
    SQLQuery = "SELECT Производители.ИмяФайлаБазы, Производители.Производитель, Производители.КодПроизводителя " & _
                "FROM Производители;"

    Fill_cmbxProizvoditel "SAPR_ASU_Izbrannoe.accdb", SQLQuery, cmbxProizvoditel


End Sub

Sub run(Artikul As String, Nazvanie As String, Cena As String, ProizvoditelID As String)
    
    txtArtikul.Value = Artikul
    txtNazvanie.Value = Nazvanie
    txtCena.Value = Cena
    For i = 0 To cmbxProizvoditel.ListCount - 1
        If cmbxProizvoditel.List(i, 2) = ProizvoditelID Then cmbxProizvoditel.ListIndex = i
    Next
    Reload_cmbxNabor
    frmDBAddToNabor.Show
End Sub

Private Sub btnAdd_Click()
    Dim DBName As String
    Dim SQLQuery As String
    DBName = "SAPR_ASU_Izbrannoe.accdb"
    
SQLQuery = "INSERT INTO Наборы ( ИзбрПозицииКод, Артикул, Название, Цена, Количество, ПроизводительКод ) " & _
            "SELECT " & cmbxNabor.List(cmbxNabor.ListIndex, 1) & ", """ & txtArtikul.Value & """, """ & txtNazvanie.Value & """, """ & txtCena.Value & """, " & txtKolichestvo.Value & ", " & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ";"

    ExecuteSQL DBName, SQLQuery
    Unload Me
    frmDBPrice.Show
End Sub

Private Sub cmbxNabor_Change()
'    If Not bBlock Then
    Load_lstvTableNabor
End Sub

Sub Reload_cmbxNabor()
    Dim SQLQuery As String
    SQLQuery = "SELECT Избранное.КодПозиции,  Избранное.Название " & _
                "FROM Избранное " & _
                "WHERE Избранное.ПодгруппыКод=2;"

    Fill_ComboBox "SAPR_ASU_Izbrannoe.accdb", SQLQuery, cmbxNabor
End Sub
Sub Load_lstvTableNabor()
    Dim colNum As Long
    lblSostav.Caption = "Состав набора: " & Fill_lstvTableNabor("SAPR_ASU_Izbrannoe.accdb", cmbxNabor.List(cmbxNabor.ListIndex, 1), lstvTableNabor)
    'выровнять ширину столбцов по заголовкам
    For colNum = 0 To lstvTableNabor.ColumnHeaders.Count - 1
        Call SendMessage(lstvTableNabor.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Private Sub txtCena_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44 And KeyAscii <> 46) Then KeyAscii = 0
End Sub

Private Sub CommandButton1_Click()
    frmDBAddGroup.Caption = "Добавить производителя"
    frmDBAddGroup.lblName = "Имя производителя:"
    frmDBAddGroup.chbxAddFile.Visible = True
    frmDBAddGroup.run 8
End Sub

Private Sub CommandButton5_Click()
    Dim DBName As String
    Dim SQLQuery As String
    If MsgBox("Удалить запись?" & vbCrLf & vbCrLf & "Производитель: " & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 0), vbYesNo + vbCritical, "Удаление записи из Производителей") = vbYes Then
        DBName = "SAPR_ASU_Izbrannoe.accdb"
        SQLQuery = "DELETE Производители.* " & _
                    "FROM Производители " & _
                    "WHERE Производители.КодПроизводителя=" & cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2) & ";"
        If Not (cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 1) <> "") Then
            ExecuteSQL DBName, SQLQuery
        End If
        UserForm_Initialize
    End If
End Sub

Private Sub CommandButton8_Click()
    Me.Hide
    Load frmDBAddNabor
    frmDBAddNabor.run txtArtikul.Value, txtNazvanie.Value, "", cmbxProizvoditel.List(cmbxProizvoditel.ListIndex, 2)
End Sub

Private Sub btnClose_Click()
Unload Me
frmDBPrice.Show
End Sub