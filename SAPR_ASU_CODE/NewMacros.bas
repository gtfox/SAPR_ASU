'

Sub Macro1()

    Application.ActiveWindow.Page.DrawLine 1.279528, 10.925197, 2.952756, 10.925197

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape1 As Visio.Shape
    Set vsoShape1 = Application.ActiveWindow.Page.Shapes.ItemFromID(66)
    vsoShape1.DrawLine 1.673228, 0#, 1.673228, -1.082677
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape2 As Visio.Shape
    Set vsoShape2 = Application.ActiveWindow.Page.Shapes.ItemFromID(66)
    vsoShape2.DrawLine 1.673228, 0#, 0.492126, 0#
    Application.EndUndoScope UndoScopeID2, True

    Dim UndoScopeID3 As Long
    UndoScopeID3 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape3 As Visio.Shape
    Set vsoShape3 = Application.ActiveWindow.Page.Shapes.ItemFromID(66)
    vsoShape3.DrawLine 0.492126, 0#, 0.492126, 0.688976
    Application.EndUndoScope UndoScopeID3, True

    Dim UndoScopeID4 As Long
    UndoScopeID4 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape4 As Visio.Shape
    Set vsoShape4 = Application.ActiveWindow.Page.Shapes.ItemFromID(66)
    vsoShape4.DrawLine 0.492126, 0.688976, 1.181102, 0.688976
    Application.EndUndoScope UndoScopeID4, True

    Dim UndoScopeID5 As Long
    UndoScopeID5 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape5 As Visio.Shape
    Set vsoShape5 = Application.ActiveWindow.Page.Shapes.ItemFromID(66)
    vsoShape5.DrawLine 1.181102, 0.688976, 1.181102, -1.771654
    Application.EndUndoScope UndoScopeID5, True

End Sub
Sub Macro2()

    Application.ActiveWindow.Page.DrawLine 0#, 0#, 0.492126, 0#

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape1 As Visio.Shape
    Set vsoShape1 = Application.ActiveWindow.Page.Shapes.ItemFromID(67)
    vsoShape1.DrawLine 0.492126, 0#, 0.492126, 0.590551
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape2 As Visio.Shape
    Set vsoShape2 = Application.ActiveWindow.Page.Shapes.ItemFromID(67)
    vsoShape2.DrawLine 0.492126, 0.590551, 0.984252, 0.590551
    Application.EndUndoScope UndoScopeID2, True

    Dim UndoScopeID3 As Long
    UndoScopeID3 = Application.BeginUndoScope("Добавить сегмент")
    Dim vsoShape3 As Visio.Shape
    Set vsoShape3 = Application.ActiveWindow.Page.Shapes.ItemFromID(67)
    vsoShape3.DrawLine 0.984252, 0.590551, 0.984252, 1.082677
    Application.EndUndoScope UndoScopeID3, True

End Sub