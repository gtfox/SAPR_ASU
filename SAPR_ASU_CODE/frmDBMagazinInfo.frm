


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

