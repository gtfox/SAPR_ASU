
Sub test()


AddSAPage "СВП"

End Sub
Sub Macro1()

    Application.ActiveWindow.SetViewRect -1.181102, 10.409449, 14.669291, 5.527559

    Application.ActiveWindow.SetViewRect -5.527559, 12.346457, 29.338583, 11.055118

    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(45), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(48), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(51), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(54), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(39), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(42), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(36), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(33), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(93), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(57), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(98), visSelect
    Application.ActiveWindow.Selection.Move 2.805118, -0.738189

End Sub