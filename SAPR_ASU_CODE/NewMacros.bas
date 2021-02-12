Sub Macro7()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(81).CellsU("EndX")
    vsoCell1.GlueToPos Application.ActiveWindow.Page.Shapes.ItemFromID(189), 0#, 1#
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(81).CellsU("BeginX")
    Set vsoCell4 = Application.ActiveWindow.Page.Shapes.ItemFromID(60).CellsSRC(7, 0, 0)
    vsoCell3.GlueTo vsoCell4
    Application.EndUndoScope UndoScopeID2, True

End Sub
Sub Macro8()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(81).CellsU("EndX")
    vsoCell1.GlueToPos Application.ActiveWindow.Page.Shapes.ItemFromID(189), 0#, 1#
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(81).CellsU("BeginX")
    vsoCell3.GlueToPos Application.ActiveWindow.Page.Shapes.ItemFromID(189), 0#, 0.76869
    Application.EndUndoScope UndoScopeID2, True

End Sub
Sub Macro9()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(83).CellsU("BeginX")
    Set vsoCell2 = Application.ActiveWindow.Page.Shapes.ItemFromID(71).CellsSRC(7, 2, 0)
    vsoCell1.GlueTo vsoCell2
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(83).CellsU("EndX")
    vsoCell3.GlueToPos Application.ActiveWindow.Page.Shapes.ItemFromID(189), 0.232343, 1#
    Application.EndUndoScope UndoScopeID2, True

End Sub