


Dim iKey As Integer



Sub run(Key As Integer)
    iKey = Key
    lblFile.Visible = False
    txtFile.Visible = False
    btnAdd.Top = 30
    btnClose.Top = 30
    chbxAddFile.Top = 30
    Me.Height = 76
    frmDBAddGroupAccess.Show
End Sub
Private Sub btnAdd_Click()
    Dim DBName As String
    Dim SQLQuery As String
    
    DBName = DBNameIzbrannoeAccess

    Select Case iKey
        Case 1, 5, 8
            SQLQuery = "INSERT INTO Производители ( Производитель" & IIf(txtFile.Value <> "", ", ИмяФайлаБазы", "") & " ) " & _
                        "SELECT """ & txtName.Value & IIf(txtFile.Value <> "", """, """ & txtFile.Value, "") & """ ;"
        Case 2, 6
            SQLQuery = "INSERT INTO Категории ( Категория ) " & _
                        "SELECT """ & txtName.Value & """ ;"
        Case 3, 7
            SQLQuery = "INSERT INTO Группы ( Группа ) " & _
                        "SELECT """ & txtName.Value & """ ;"
        Case 4
            SQLQuery = "INSERT INTO Подгруппы ( Подгруппа ) " & _
                        "SELECT """ & txtName.Value & """ ;"
        Case Else
            
    End Select
    
    ExecuteSQL DBName, SQLQuery
    Unload Me
    
    Select Case iKey
        Case 1
            frmDBAddToIzbrannoeAccess.UserForm_Initialize
        Case 2, 3, 4
            frmDBAddToIzbrannoeAccess.Reset_FiltersCmbx
        Case 5
            frmDBAddNaborAccess.UserForm_Initialize
        Case 6, 7
            frmDBAddNaborAccess.Reset_FiltersCmbx
        Case 8
            frmDBAddToNaborAccess.UserForm_Initialize
        Case Else

    End Select
    
End Sub

Private Sub chbxAddFile_Change()
        If chbxAddFile.Value = True Then
            lblFile.Visible = True
            txtFile.Visible = True
            btnAdd.Top = 48
            btnClose.Top = 48
            chbxAddFile.Top = 48
            Me.Height = 94
            chbxAddFile.Value = True
        Else
            lblFile.Visible = False
            txtFile.Visible = False
            btnAdd.Top = 30
            btnClose.Top = 30
            chbxAddFile.Top = 30
            Me.Height = 76
            chbxAddFile.Value = False
        End If
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    InitCustomCCPMenu Me 'Контекстное меню для TextBox
End Sub

Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub