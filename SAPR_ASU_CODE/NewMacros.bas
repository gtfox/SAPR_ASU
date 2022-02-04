'

Sub Macro1()
    ActiveWindow.DeselectAll
    SSS = Application.ActiveWindow.Page.Shapes.ItemFromID(84).UniqueID(visGetOrMakeGUID)
End Sub