
Private Sub btnAddPriceExcel_Click()
    If tbProizvoditel = "" Then
        MsgBox "Название производителя пустое" & vbCrLf & vbCrLf & "Необходимо ввести название производителя", vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
    Else
        WizardAddPriceExcel tbProizvoditel
        Unload Me
    End If
End Sub

Private Sub cbColor_Click()
    ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = cbColor
End Sub

Private Sub UserForm_Initialize()
    InitCustomCCPMenu Me 'Контекстное меню для TextBox
    Reload_cmbx
    With ActiveDocument.DocumentSheet
        tbSA_FR_Shifr = .Cells("User.SA_FR_Shifr").ResultStr(0)
        tbSA_FR_Zakazchik = .Cells("User.SA_FR_Zakazchik").ResultStr(0)
        tbSA_FR_OpisaniePoekta = .Cells("User.SA_FR_OpisaniePoekta").ResultStr(0)
        tbSA_FR_Stadia = .Cells("User.SA_FR_Stadia").ResultStr(0)
        tbSA_FR_ProekOrg = .Cells("User.SA_FR_ProekOrg").ResultStr(0)
        tbSA_FR_OOO = .Cells("User.SA_FR_OOO").ResultStr(0)
        tbSA_FR_OoOoOo = .Cells("User.SA_FR_OoOoOo").ResultStr(0)
        tbSA_FR_Data = .Cells("User.SA_FR_Data").ResultStr(0)
        tbSA_PoleA = .Cells("User.SA_PoleA").Result(visMillimeters)
        tbSA_PoleVert = .Cells("User.SA_PoleVert").Result(visMillimeters)
        tbSA_Pole1 = .Cells("User.SA_Pole1").Result(visMillimeters)
        tbSA_PoleGor = .Cells("User.SA_PoleGor").Result(visMillimeters)
        tbSA_Stranica = .Cells("User.SA_Stranica").ResultStr(0)
        tbSA_Adres = .Cells("User.SA_Adres").ResultStr(0)
        tbSA_FR_OffsetFrame = .Cells("User.SA_FR_OffsetFrame").Result(visMillimeters)
        tbSA_PrefElement = .Cells("User.SA_PrefElement").ResultStr(0)
        tbSA_PrefShkaf = .Cells("User.SA_PrefShkaf").ResultStr(0)
        tbSA_PrefMesto = .Cells("User.SA_PrefMesto").ResultStr(0)
        cbISO = .Cells("User.SA_ISO").Result(0)
        tbSA_nRows = .Cells("User.SA_nRows").Result(0)
    End With
    MultiPage1.Value = 0
End Sub

Sub Reload_cmbx()
    Fill_cmbx cmbxSA_FR_Razrabotal
    Fill_cmbx cmbxSA_FR_Proveril
    Fill_cmbx cmbxSA_FR_Gip
    Fill_cmbx cmbxSA_FR_NachOtdela
    Fill_cmbx cmbxSA_FR_NKontr
    Fill_cmbx cmbxSA_FR_Utverdil
    With ActiveDocument.DocumentSheet
        cmbxSA_FR_Razrabotal = .Cells("User.SA_FR_Razrabotal").ResultStr(0)
        cmbxSA_FR_Proveril = .Cells("User.SA_FR_Proveril").ResultStr(0)
        cmbxSA_FR_Gip = .Cells("User.SA_FR_Gip").ResultStr(0)
        cmbxSA_FR_NachOtdela = .Cells("User.SA_FR_NachOtdela").ResultStr(0)
        cmbxSA_FR_NKontr = .Cells("User.SA_FR_NKontr").ResultStr(0)
        cmbxSA_FR_Utverdil = .Cells("User.SA_FR_Utverdil").ResultStr(0)
    End With
End Sub

Sub Fill_cmbx(cmbxCmbx As ComboBox)
    Dim mstrFamilii() As String
    Dim i As Integer
    cmbxCmbx.Clear
    mstrFamilii = Split(ActiveDocument.DocumentSheet.Cells("User.SA_FR_Razrabotal.Prompt").ResultStr(0), ";")
    For i = 0 To UBound(mstrFamilii)
        cmbxCmbx.AddItem mstrFamilii(i)
    Next
End Sub

Sub Add_cmbx(cmbxCmbx As ComboBox)
    Dim strFamilii As String
    strFamilii = ActiveDocument.DocumentSheet.Cells("User.SA_FR_Razrabotal.Prompt").ResultStr(0) + ";" + cmbxCmbx.text
    ActiveDocument.DocumentSheet.Cells("User.SA_FR_Razrabotal.Prompt").Formula = """" + strFamilii + """"
    Reload_cmbx
End Sub

Sub Del_cmbx(cmbxCmbx As ComboBox)
    Dim strFamiliyaToDel As String
    Dim strFamilii As String
    strFamiliyaToDel = cmbxCmbx.text
    strFamilii = Replace(ActiveDocument.DocumentSheet.Cells("User.SA_FR_Razrabotal.Prompt").ResultStr(0), strFamiliyaToDel, "")
    strFamilii = Replace(strFamilii, ";;", ";")
    strFamilii = IIf(Left(strFamilii, 1) = ";", Right(strFamilii, Len(strFamilii) - 1), IIf(Right(strFamilii, 1) = ";", Left(strFamilii, Len(strFamilii) - 1), strFamilii))
    ActiveDocument.DocumentSheet.Cells("User.SA_FR_Razrabotal.Prompt").Formula = """" + strFamilii + """"
    Del_DocumentSheet strFamiliyaToDel
    Reload_cmbx
End Sub

Sub Del_DocumentSheet(strFamiliyaToDel As String)
    With ActiveDocument.DocumentSheet
        If .Cells("User.SA_FR_Razrabotal").ResultStr(0) = strFamiliyaToDel Then .Cells("User.SA_FR_Razrabotal").Formula = """"""
        If .Cells("User.SA_FR_Proveril").ResultStr(0) = strFamiliyaToDel Then .Cells("User.SA_FR_Proveril").Formula = """"""
        If .Cells("User.SA_FR_Gip").ResultStr(0) = strFamiliyaToDel Then .Cells("User.SA_FR_Gip").Formula = """"""
        If .Cells("User.SA_FR_NachOtdela").ResultStr(0) = strFamiliyaToDel Then .Cells("User.SA_FR_NachOtdela").Formula = """"""
        If .Cells("User.SA_FR_NKontr").ResultStr(0) = strFamiliyaToDel Then .Cells("User.SA_FR_NKontr").Formula = """"""
        If .Cells("User.SA_FR_Utverdil").ResultStr(0) = strFamiliyaToDel Then .Cells("User.SA_FR_Utverdil").Formula = """"""
    End With
End Sub

Private Sub CommandButton22_Click()
    Add_cmbx cmbxSA_FR_Razrabotal
End Sub
Private Sub CommandButton23_Click()
    Del_cmbx cmbxSA_FR_Razrabotal
End Sub
Private Sub CommandButton24_Click()
    Del_cmbx cmbxSA_FR_Proveril
End Sub
Private Sub CommandButton25_Click()
    Add_cmbx cmbxSA_FR_Proveril
End Sub
Private Sub CommandButton26_Click()
    Del_cmbx cmbxSA_FR_Gip
End Sub
Private Sub CommandButton27_Click()
    Add_cmbx cmbxSA_FR_Gip
End Sub
Private Sub CommandButton28_Click()
    Del_cmbx cmbxSA_FR_NachOtdela
End Sub
Private Sub CommandButton29_Click()
    Add_cmbx cmbxSA_FR_NachOtdela
End Sub
Private Sub CommandButton30_Click()
    Del_cmbx cmbxSA_FR_NKontr
End Sub
Private Sub CommandButton31_Click()
    Add_cmbx cmbxSA_FR_NKontr
End Sub
Private Sub CommandButton32_Click()
    Del_cmbx cmbxSA_FR_Utverdil
End Sub
Private Sub CommandButton33_Click()
    Add_cmbx cmbxSA_FR_Utverdil
End Sub
Private Sub CommandButton38_Click()
    With ActiveDocument.DocumentSheet
        .Cells("User.SA_FR_Shifr").Formula = """" + Replace(tbSA_FR_Shifr, """", """""") + """"
        .Cells("User.SA_FR_Zakazchik").Formula = """" + Replace(tbSA_FR_Zakazchik, """", """""") + """"
        .Cells("User.SA_FR_OpisaniePoekta").Formula = """" + Replace(tbSA_FR_OpisaniePoekta, """", """""") + """"
        .Cells("User.SA_FR_Stadia").Formula = """" + Replace(tbSA_FR_Stadia, """", """""") + """"
        .Cells("User.SA_FR_ProekOrg").Formula = """" + Replace(tbSA_FR_ProekOrg, """", """""") + """"
        .Cells("User.SA_FR_OOO").Formula = """" + Replace(tbSA_FR_OOO, """", """""") + """"
        .Cells("User.SA_FR_OoOoOo").Formula = """" + Replace(tbSA_FR_OoOoOo, """", """""") + """"
        .Cells("User.SA_FR_Razrabotal").Formula = """" + cmbxSA_FR_Razrabotal + """"
        .Cells("User.SA_FR_Proveril").Formula = """" + cmbxSA_FR_Proveril + """"
        .Cells("User.SA_FR_Gip").Formula = """" + cmbxSA_FR_Gip + """"
        .Cells("User.SA_FR_NachOtdela").Formula = """" + cmbxSA_FR_NachOtdela + """"
        .Cells("User.SA_FR_NKontr").Formula = """" + cmbxSA_FR_NKontr + """"
        .Cells("User.SA_FR_Utverdil").Formula = """" + cmbxSA_FR_Utverdil + """"
        .Cells("User.SA_FR_Data").Formula = """" + Replace(tbSA_FR_Data, """", """""") + """"
    End With
    SetTheDocInAllFrame
End Sub

Private Sub CommandButton39_Click()
    With ActiveDocument.DocumentSheet
        .Cells("User.SA_PoleA").FormulaU = CStr(tbSA_PoleA + " mm")
        .Cells("User.SA_PoleVert").Formula = CStr(tbSA_PoleVert + " mm")
        .Cells("User.SA_Pole1").Formula = CStr(tbSA_Pole1 + " mm")
        .Cells("User.SA_PoleGor").Formula = CStr(tbSA_PoleGor + " mm")
        .Cells("User.SA_Stranica").Formula = """" + tbSA_Stranica + """"
        .Cells("User.SA_Adres").Formula = """" + tbSA_Adres + """"
        .Cells("User.SA_PrefMesto").Formula = """" + tbSA_PrefMesto + """"
        .Cells("User.SA_PrefShkaf").Formula = """" + tbSA_PrefShkaf + """"
        .Cells("User.SA_PrefElement").Formula = """" + tbSA_PrefElement + """"
        .Cells("User.SA_FR_OffsetFrame").Formula = CStr(tbSA_FR_OffsetFrame + " mm")
    End With
End Sub

Private Sub CommandButton37_Click()
    ActiveDocument.DocumentSheet.Cells("User.SA_nRows").Formula = tbSA_nRows
End Sub
Private Sub cbISO_Click()
    ActiveDocument.DocumentSheet.Cells("User.SA_ISO").Formula = cbISO
End Sub
Private Sub UserForm_Terminate()
    DelCustomCCPMenu 'Удаления контекстного меню для TextBox
End Sub
