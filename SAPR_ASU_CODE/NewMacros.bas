Sub MacroOff()
    'ThisDocument.BlockMacros = False
    Application.EventsEnabled = 0
End Sub
Sub MacroOn()
    Application.EventsEnabled = -1
End Sub
Sub Macro3()


End Sub
Sub Macro4()

    Application.ActiveWindow.Selection.Copy
    Application.EventsEnabled = 0
    Application.ActivePage.Paste
    DoEvents
    Application.EventsEnabled = -1
End Sub