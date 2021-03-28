


Public linkCatalog As String
Public linkFind As String

Sub run()
    Dim fWidth As Long

    If lblNazvanie.Width < imgKartinka.Width + frameMain.Width Then
        Me.Width = imgKartinka.Width + frameMain.Width + 18
    Else
        Me.Width = lblNazvanie.Width + 18
    End If
    
'    Me.Width = imgKartinka.Width + IIf(lblNazvanie.Width < frameMain.Width, frameMain.Width, lblNazvanie.Width) + 18
'    lblNazvanie.Left = imgKartinka.Width + 6
    btnClose.Left = Me.Width - btnClose.Width - 9
    frameMain.Left = imgKartinka.Width + 6
    If imgKartinka.Height > 162 Then
        Me.Height = imgKartinka.Height + 54
    Else
        Me.Height = 162
    End If
    
    Me.Show
End Sub

Private Sub btnCatalog_Click()
    If linkCatalog <> "" Then CreateObject("WScript.Shell").run linkCatalog
End Sub

Private Sub btnFind_Click()
    CreateObject("WScript.Shell").run linkFind
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


