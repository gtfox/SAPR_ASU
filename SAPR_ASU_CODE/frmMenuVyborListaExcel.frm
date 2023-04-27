Option Explicit

Private Sub UserForm_Initialize()
    cmbxNazvanieLista.style = fmStyleDropDownList
    frmClose = False
End Sub

Sub run(wb As Excel.Workbook)
    Dim sht As Excel.Worksheet
    cmbxNazvanieLista.Clear
    For Each sht In wb.Worksheets
        If sht.Visible = xlSheetVisible Then cmbxNazvanieLista.AddItem sht.name
    Next
    cmbxNazvanieLista.ListIndex = 0
    Me.Show
End Sub

Private Sub btnSelSheet_Click()
    Excel_imya_lista = cmbxNazvanieLista.text
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    frmClose = True
    Unload Me
End Sub