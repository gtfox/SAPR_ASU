Sub Macro3()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(1249).CellsU("EndX")
    Set vsoCell2 = Application.ActiveWindow.Page.Shapes.ItemFromID(1246).CellsSRC(7, 1, 0)
    vsoCell1.GlueTo vsoCell2
    Application.EndUndoScope UndoScopeID1, True

    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Изменить размер объекта")
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(1249).CellsU("BeginX")
    Set vsoCell4 = Application.ActiveWindow.Page.Shapes.ItemFromID(477).CellsSRC(visSectionConnectionPts, visRowConnectionPts, 0)
    vsoCell3.GlueTo vsoCell4
    Application.EndUndoScope UndoScopeID2, True


End Sub