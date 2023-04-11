Option Explicit
Sub Macro2()

    Application.ActiveWindow.Page.Shapes.ItemFromID(296).OpenSheetWindow

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Добавить раздел")
    Application.ActiveWindow.Shape.AddSection visSectionAction
    Application.ActiveWindow.Shape.AddRow visSectionAction, visRowLast, visTagDefault
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionMenu).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionPrompt).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionHelp).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionAction).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionChecked).FormulaForceU = "0"
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionDisabled).FormulaForceU = "0"
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionReadOnly).FormulaForceU = "FALSE"
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionInvisible).FormulaForceU = "FALSE"
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionBeginGroup).FormulaForceU = "FALSE"
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionTagName).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionButtonFace).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionSortKey).FormulaForceU = """"""
    Application.EndUndoScope UndoScopeID1, True

End Sub