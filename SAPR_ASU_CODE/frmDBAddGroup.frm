Dim iKey As Integer

Private Sub btnClose_Click()
    Unload Me
End Sub

Sub run(Key As Integer)
    iKey = Key
    frmDBAddGroup.Show
End Sub
Private Sub btnAdd_Click()
    Dim DBName As String
    Dim SQLQuery As String
    
    DBName = "SAPR_ASU_Izbrannoe.accdb"

    Select Case iKey
        Case 1
            SQLQuery = "INSERT INTO Производители ( Производитель" & IIf(textFile.Value <> "", ", ИмяФайлаБазы", "") & " ) " & _
                        "SELECT """ & textName.Value & IIf(textFile.Value <> "", """, """ & textFile.Value, "") & """ ;"
        Case 2
            SQLQuery = "INSERT INTO Категории ( Категория ) " & _
                        "SELECT """ & textName.Value & """ ;"
        Case 3
            SQLQuery = "INSERT INTO Группы ( Группа ) " & _
                        "SELECT """ & textName.Value & """ ;"
        Case 4
            SQLQuery = "INSERT INTO Подгруппы ( Подгруппа ) " & _
                        "SELECT """ & textName.Value & """ ;"
        Case Else
            
    End Select
    
    ExecuteSQL DBName, SQLQuery
    Unload Me
    If iKey = 1 Then
        frmDBAddToIzbrannoe.UserForm_Initialize
    Else
        frmDBAddToIzbrannoe.Reset_FiltersCmbx
    End If
    
End Sub