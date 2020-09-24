
Sub wire_conn()

    ActiveWindow.DeselectAll
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(369), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(370), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(371), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(372), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(373), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(374), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(375), visSelect
    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(376), visSelect
    Application.ActiveWindow.Selection.Duplicate

    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(360).CellsU("BeginX")
    Set vsoCell2 = Application.ActiveWindow.Page.Shapes.ItemFromID(591).CellsSRC(7, 0, 0)
    vsoCell1.GlueTo vsoCell2
    Set vsoCell1 = Application.ActiveWindow.Page.Shapes.ItemFromID(360).CellsU("EndX")
    Set vsoCell2 = Application.ActiveWindow.Page.Shapes.ItemFromID(506).CellsSRC(7, 1, 0)
    vsoCell1.GlueTo vsoCell2
    Dim vsoCell3 As Visio.Cell
    Dim vsoCell4 As Visio.Cell
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(361).CellsU("BeginX")
    Set vsoCell4 = Application.ActiveWindow.Page.Shapes.ItemFromID(594).CellsSRC(7, 0, 0)
    vsoCell3.GlueTo vsoCell4
    Set vsoCell3 = Application.ActiveWindow.Page.Shapes.ItemFromID(361).CellsU("EndX")
    Set vsoCell4 = Application.ActiveWindow.Page.Shapes.ItemFromID(509).CellsSRC(7, 1, 0)
    vsoCell3.GlueTo vsoCell4
    Dim vsoCell5 As Visio.Cell
    Dim vsoCell6 As Visio.Cell
    Set vsoCell5 = Application.ActiveWindow.Page.Shapes.ItemFromID(362).CellsU("BeginX")
    Set vsoCell6 = Application.ActiveWindow.Page.Shapes.ItemFromID(597).CellsSRC(7, 0, 0)
    vsoCell5.GlueTo vsoCell6
    Set vsoCell5 = Application.ActiveWindow.Page.Shapes.ItemFromID(362).CellsU("EndX")
    Set vsoCell6 = Application.ActiveWindow.Page.Shapes.ItemFromID(564).CellsSRC(7, 1, 0)
    vsoCell5.GlueTo vsoCell6
    Dim vsoCell7 As Visio.Cell
    Dim vsoCell8 As Visio.Cell
    Set vsoCell7 = Application.ActiveWindow.Page.Shapes.ItemFromID(363).CellsU("BeginX")
    Set vsoCell8 = Application.ActiveWindow.Page.Shapes.ItemFromID(600).CellsSRC(7, 0, 0)
    vsoCell7.GlueTo vsoCell8
    Set vsoCell7 = Application.ActiveWindow.Page.Shapes.ItemFromID(363).CellsU("EndX")
    Set vsoCell8 = Application.ActiveWindow.Page.Shapes.ItemFromID(567).CellsSRC(7, 1, 0)
    vsoCell7.GlueTo vsoCell8
    Dim vsoCell9 As Visio.Cell
    Dim vsoCell10 As Visio.Cell
    Set vsoCell9 = Application.ActiveWindow.Page.Shapes.ItemFromID(364).CellsU("BeginX")
    Set vsoCell10 = Application.ActiveWindow.Page.Shapes.ItemFromID(606).CellsSRC(7, 0, 0)
    vsoCell9.GlueTo vsoCell10
    Set vsoCell9 = Application.ActiveWindow.Page.Shapes.ItemFromID(364).CellsU("EndX")
    Set vsoCell10 = Application.ActiveWindow.Page.Shapes.ItemFromID(570).CellsSRC(7, 1, 0)
    vsoCell9.GlueTo vsoCell10
    Dim vsoCell11 As Visio.Cell
    Dim vsoCell12 As Visio.Cell
    Set vsoCell11 = Application.ActiveWindow.Page.Shapes.ItemFromID(365).CellsU("BeginX")
    Set vsoCell12 = Application.ActiveWindow.Page.Shapes.ItemFromID(609).CellsSRC(7, 0, 0)
    vsoCell11.GlueTo vsoCell12
    Set vsoCell11 = Application.ActiveWindow.Page.Shapes.ItemFromID(365).CellsU("EndX")
    Set vsoCell12 = Application.ActiveWindow.Page.Shapes.ItemFromID(573).CellsSRC(7, 1, 0)
    vsoCell11.GlueTo vsoCell12
    Dim vsoCell13 As Visio.Cell
    Dim vsoCell14 As Visio.Cell
    Set vsoCell13 = Application.ActiveWindow.Page.Shapes.ItemFromID(366).CellsU("BeginX")
    Set vsoCell14 = Application.ActiveWindow.Page.Shapes.ItemFromID(612).CellsSRC(7, 0, 0)
    vsoCell13.GlueTo vsoCell14
    Set vsoCell13 = Application.ActiveWindow.Page.Shapes.ItemFromID(366).CellsU("EndX")
    Set vsoCell14 = Application.ActiveWindow.Page.Shapes.ItemFromID(576).CellsSRC(7, 1, 0)
    vsoCell13.GlueTo vsoCell14
    Dim vsoCell15 As Visio.Cell
    Dim vsoCell16 As Visio.Cell
    Set vsoCell15 = Application.ActiveWindow.Page.Shapes.ItemFromID(367).CellsU("BeginX")
    Set vsoCell16 = Application.ActiveWindow.Page.Shapes.ItemFromID(615).CellsSRC(7, 0, 0)
    vsoCell15.GlueTo vsoCell16
    Set vsoCell15 = Application.ActiveWindow.Page.Shapes.ItemFromID(367).CellsU("EndX")
    Set vsoCell16 = Application.ActiveWindow.Page.Shapes.ItemFromID(579).CellsSRC(7, 1, 0)
    vsoCell15.GlueTo vsoCell16

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