

Option Explicit

Private Sub UserForm_Initialize()
    cmbxNazvanieLista.style = fmStyleDropDownList
    frmClose = False
    Fill_cmbxNazvanieLista
End Sub

Private Sub btnSelSheet_Click()
    Excel_imya_lista = cmbxNazvanieLista.Text
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub

Sub Fill_cmbxNazvanieLista()
    Dim sht As Excel.Worksheet
    cmbxNazvanieLista.Clear
    For Each sht In sp.Worksheets
        cmbxNazvanieLista.AddItem sht.Name
    Next
    cmbxNazvanieLista.ListIndex = 0
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    frmClose = True
    Unload Me
End Sub