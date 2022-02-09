'

Sub Macro1()
ttt = ReplaceSequenceInString("1;2;4;5;6;7;8;9;10;11;12;13;14;15;25;26;27;28;29;30;31;55;56;57;58;59;60;77")
    ActiveWindow.DeselectAll
    SSS = Application.ActiveWindow.Page.Shapes.ItemFromID(84).UniqueID(visGetOrMakeGUID)
End Sub
