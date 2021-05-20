Sub Macro21111()
    Application.CommandBars("Standard").Visible = False
    Application.CommandBars("Formatting").Visible = False
    Application.CommandBars("View").Visible = False
    Application.CommandBars("Data").Visible = False
    Application.CommandBars("Action").Visible = False
    Application.CommandBars("Stencil").Visible = False
    Application.CommandBars("Stop Recording").Visible = False
    Application.CommandBars("Snap & Glue").Visible = False
    Application.CommandBars("Developer").Visible = False
    Application.CommandBars("Drawing").Visible = False
    Application.CommandBars("Picture").Visible = False
    Application.CommandBars("Format Text").Visible = False
    Application.CommandBars("Format Shape").Visible = False
    Application.CommandBars("САПР АСУ").Visible = False
    Application.CommandBars("Standard").Visible = True
    Application.CommandBars("Formatting").Visible = True
    Application.CommandBars("Web").Visible = True
    Application.CommandBars("View").Visible = True
    Application.CommandBars("Data").Visible = True
    Application.CommandBars("Action").Visible = True
    Application.CommandBars("Layout & Routing").Visible = True
    Application.CommandBars("Stencil").Visible = True
    Application.CommandBars("Stop Recording").Visible = True
    Application.CommandBars("Snap & Glue").Visible = True
    Application.CommandBars("Developer").Visible = True
    Application.CommandBars("Reviewing").Visible = True
    Application.CommandBars("Drawing").Visible = True
    Application.CommandBars("Picture").Visible = True
    Application.CommandBars("Ink").Visible = True
    Application.CommandBars("Format Text").Visible = True
    Application.CommandBars("Format Shape").Visible = True
    Application.CommandBars("САПР АСУ").Visible = True

End Sub
Sub Macro3111()
    Application.CommandBars("Reviewing").Visible = False
    Application.CommandBars("Web").Visible = False
    Application.CommandBars("Ink").Visible = False
    Application.CommandBars("Stencil").Visible = False
    Application.CommandBars("Picture").Visible = False
    Application.CommandBars("Layout & Routing").Visible = False
    Application.CommandBars("Data").Visible = False

End Sub

Sub Macro33333()


    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Линейка и сетка")
    Dim vsoShape1 As Shape
    Set vsoShape1 = Application.ActiveWindow.Page.PageSheet
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visXRulerOrigin).FormulaU = "95 mm"
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visYRulerOrigin).FormulaU = "170 mm"
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visXGridOrigin).FormulaU = "95 mm"
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visYGridOrigin).FormulaU = "170 mm"
    Application.EndUndoScope UndoScopeID1, True

End Sub
Sub Macro43333()

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Линейка и сетка")
    Dim vsoShape1 As Shape
    Set vsoShape1 = Application.ActiveWindow.Page.PageSheet
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visXRulerOrigin).FormulaU = "0 mm"
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visYRulerOrigin).FormulaU = "0 mm"
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visXGridOrigin).FormulaU = "0 mm"
    vsoShape1.CellsSRC(visSectionObject, visRowRulerGrid, visYGridOrigin).FormulaU = "0 mm"
    Application.EndUndoScope UndoScopeID1, True

End Sub
Sub Macro1333()
    Dim vsoShape As Shape
    Set vsoShape = Application.ActiveWindow.Page.PageSheet
    With vsoShape
        .AddSection visSectionAction
        .AddRow visSectionAction, visRowLast, visTagDefault
        .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить элементы со схемы"""
        .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageVIDAddElementsFrm"")"
    End With

End Sub


Sub Macro41()

    Application.ActiveWindow.Page.PageSheet.OpenSheetWindow

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

    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionMenu).RowNameU = "Row_1r"

    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionMenu).FormulaU = xdfhxfr

    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionAction).FormulaU = "dfbh"

    Application.ActiveWindow.Shape.CellsSRC(visSectionAction, 0, visActionSortKey).FormulaU = 10

    Application.ActiveWindow.Close

End Sub
Sub Macro11()

    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(155), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(167), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(144), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(73), visSelect
    ActiveWindow.Selection.Group

    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(178), visSelect
    ActiveWindow.Selection.Ungroup

End Sub
Sub Macro2()
ActiveWindow.Selection.Item(1).Cells("Actions.Rotate.ButtonFace").FormulaU = "IF(Actions.Rotate.Action,""199"",""198"")" '128 129
ActiveWindow.Selection.Item(1).Cells("Actions.AddReference.ButtonFace").FormulaU = "2651" '1623
ActiveWindow.Selection.Item(1).Cells("Actions.Thumb.ButtonFace").FormulaU = "2871" '256
End Sub