Option Explicit

Sub cp()
    Application.EventsEnabled = False
    ActiveWindow.Selection.Copy visCopyPasteNoTranslate
    ActivePage.Paste visCopyPasteNoTranslate
    Application.EventsEnabled = True
End Sub