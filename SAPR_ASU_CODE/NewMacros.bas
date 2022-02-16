'

Sub Macro4()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Слой")
    Application.ActiveWindow.Page.Shapes.ItemFromID(190).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(191).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(192).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(193).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(194).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(195).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(196).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(197).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.ActiveWindow.Page.Shapes.ItemFromID(198).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """1"""
    Application.EndUndoScope UndoScopeID1, True

End Sub