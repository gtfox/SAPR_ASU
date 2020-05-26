Attribute VB_Name = "MISC"


Sub Macro2() '15924991 Цвет листа как Splan 7

    Application.Settings.DeveloperMode = True
    Application.Settings.FreeformDrawingPrecision = 5
    Application.Settings.FreeformDrawingSmoothing = 5
    Application.DrawingPaths = "D:\YandexDisk\VISIO\SAPR_ASU"
    Application.TemplatePaths = "C:\Program Files\Radica\Electra\"
    Application.StencilPaths = "C:\Program Files\Radica\Electra\"
    Application.HelpPaths = ""
    Application.AddonPaths = ""
    Application.StartupPaths = ""
    Application.MyShapesPath = "D:\YandexDisk\VISIO\SAPR_ASU"
    Application.Settings.DrawingPageColor = 15924991 '15924991 Цвет листа в Splan 7 (белый 16777215)
    Application.Settings.FullScreenBackgroundColor = 0
    Application.Settings.EnableAutoConnect = False

End Sub



'Public Enum tList
'    A4m = 1
'    A4b = 2
'    A3m1 = 3
'    A3m2 = 4
'    A3b1 = 5
'    A3b2 = 6
'End Enum

Private Sub SetStyleGost() 'Изменение стилей под Гост

    Dim vsoStyle As Visio.Style
    Set vsoStyle = Application.ActiveDocument.Styles("EE Normal")
    vsoStyle.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.2 mm"
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaU = 93
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = 2
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "11 pt"
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterDblUnderline).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterOverline).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterStrikethru).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, 11).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterDoubleStrikethrough).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterRTLText).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterUseVertical).FormulaU = False

    Set vsoStyle = Application.ActiveDocument.Styles("Pin Normal")
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaU = 93
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = 2
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "8 pt"
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterDblUnderline).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterOverline).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterStrikethru).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, 11).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterDoubleStrikethrough).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterRTLText).FormulaU = False
    vsoStyle.CellsSRC(visSectionCharacter, 0, visCharacterUseVertical).FormulaU = False
    
    'сетка 2,5 мм
    Dim vsoShape As Shape
    Dim vsoPage As Visio.Page
    For Each vsoPage In Application.ActiveDocument.Pages
        Set vsoShape = vsoPage.PageSheet
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visXGridDensity).FormulaU = "0"
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visXGridSpacing).FormulaU = "2.5 mm"
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visYGridDensity).FormulaU = "0"
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visYGridSpacing).FormulaU = "2.5 mm"
    Next
    
    'Сила привязки к сетке в пикселях
    'Сервис -> Привязать и приклеить -> Дополнительно -> Сетка = 100
    Application.Settings.SnapStrengthGridX = 100
    Application.Settings.SnapStrengthGridY = 100
    
    Application.Settings.EnableAutoConnect = False
    
End Sub

Private Sub Tune_Stencils() 'переделка шаблонов электры под гост (перед выполнением макроса надо окрыть шаблоны и сделать их редактируемыми)

    Dim appdoc As Document
    Dim appcol As Collection
    Set appcol = New Collection
    Dim mast As Master
    Dim ss As String
        
    'выбираем нужные шаблоны для измениния
    For Each appdoc In Application.Documents
        If (appdoc.Creator = "Electra" Or appdoc.Creator = "Pneumata" Or appdoc.Creator = "Hydraula") And Not (appdoc.Title = "Electra" Or appdoc.Title = "Layout" Or appdoc.Title = "Layout 3D" Or appdoc.Title = "Reports" Or appdoc.Title = "IEC Parts" Or appdoc.Title = "Title Blocks") Then
            appcol.Add appdoc
        End If
    Next
    
    For Each appdoc In appcol
        For Each mast In appdoc.Masters
            If InStr(1, mast.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageScale).FormulaU, "in") Then 'не трогаем элемент если он в мм (значит он уже был изменён)
                
                'масштаб под гост
                mast.Shapes(1).Cells("Width").FormulaForceU = "guard(" & str(mast.Shapes(1).Cells("Width").Result(visInches) * 1.181102362) & ")"
                mast.Shapes(1).Cells("Height").FormulaForceU = "guard(" & str(mast.Shapes(1).Cells("Height").Result(visInches) * 1.181102362) & ")"
                
                If mast.Shapes(1).Shapes.Count > 0 Then
                    'скрываем описание
                    On Error Resume Next
                    mast.Shapes(1).Shapes("Desc").CellsU("HideText").FormulaU = "TRUE"
                    'поворот фигур
                    mast.Shapes(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormAngle).FormulaU = "=IF(Actions.Row_2.Action,-90 deg,0 deg)"
                    mast.Shapes(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormFlipX).FormulaU = 0
                    'только группа
                    mast.Shapes(1).CellsSRC(visSectionObject, visRowGroup, visGroupSelectMode).FormulaU = "0"
                End If
                
                'страница в милиметрах чтобы электра не запускала конвертацию in->mm
                mast.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageScale).FormulaU = "1 mm"
                mast.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = "1 mm"
                
            End If
        Next mast
        appdoc.Save
    Next appdoc

End Sub

Private Sub Pole_Spec()   'массовая установка свойств в спецификации (поля 10pt до и 5/1pt  после текста во 2 и 9 столбце )

    For i = 1 To 30
    
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".2").CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "10 pt"
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".2").CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "10 pt"
        
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".9").CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "5 pt"
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".9").CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "1 pt"
    Next i
End Sub

Private Sub Pole_VRCh() 'установка полей текста ВРЧ 10pt

For i = 1 To 15
    Application.ActiveWindow.Page.Shapes("В Р Ч").Shapes(i).Shapes(3).CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "10 pt"
Next i

End Sub

Sub AddToolBar()

    Dim Bar As CommandBar

    Set Bar = Application.CommandBars.Add(Position:=msoBarFloating, Temporary:=True) 'msoBarTop
    With Bar
        .Name = "САПР АСУ"
        .Visible = True
    End With
    
    AddButtons

End Sub

Private Sub AddButtons()

    Dim Bar As CommandBar
    Dim Button As CommandBarButton

    Set Bar = Application.CommandBars("САПР АСУ")
    
    '---Кнопка Блокировки рамки
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "БлокРамки"
        .Tag = "LockTitle"
        .OnAction = "LockTitleBlock"
        .TooltipText = "Блокировка рамки"
        .FaceID = 519
    End With
    
    Set Button = Nothing
           
End Sub

Public Sub LockTitleBlock() 'Блокировка слоя рамки

    Dim vsoLayer1 As Visio.Layer
    Set vsoLayer1 = Application.ActiveWindow.Page.Layers("TitleBlock")
    
    If vsoLayer1.CellsC(visLayerLock).FormulaU = 0 Then
        
        'Блокруем слой
        vsoLayer1.CellsC(visLayerLock).FormulaU = "1"
        vsoLayer1.CellsC(visLayerColor).FormulaU = "19"
        vsoLayer1.CellsC(visLayerSnap).FormulaU = "0"
        vsoLayer1.CellsC(visLayerGlue).FormulaU = "0"
      
        Application.CommandBars("САПР АСУ").Controls("БлокРамки").State = msoButtonDown

    Else
        
        'Разблокруем слой
        vsoLayer1.CellsC(visLayerLock).FormulaU = "0"
        vsoLayer1.CellsC(visLayerColor).FormulaU = "255"
        vsoLayer1.CellsC(visLayerSnap).FormulaU = "0"
        vsoLayer1.CellsC(visLayerGlue).FormulaU = "0"
        
        Application.CommandBars("САПР АСУ").Controls("БлокРамки").State = msoButtonUp
    End If

End Sub

Private Sub mcr1() 'добавление панельки

Set cbar1 = Application.CommandBars.Add(Name:="Custom1", Position:=msoBarFloating)
cbar1.Visible = True


Set myControl = cbar1.Controls _
    .Add(Type:=msoControlComboBox, Before:=1)
With myControl
    .AddItem Text:="First Item", Index:=1
    .AddItem Text:="Second Item", Index:=2
    .DropDownLines = 3
    .DropDownWidth = 75
    .ListHeaderCount = 0
    .OnAction = "SAPR_ASU.LockTitleBlock"
    
End With

End Sub

'Private Sub ttt()
'Dim List As tList
'List = A4m
'
'    Select Case List
'        Case tList.A4m
'            ' Process.
'            List = A3b1
'        Case tList.A4b
'            ' Process.
'        Case tList.A3b1
'            ' Process.
'        Case Else
'
'    End Select
'
'End Sub


Public Sub UpdateZoneOnPage() 'Обновить сетку коодинат на листе
    Dim tShape As Visio.Shape
    For Each tShape In ActivePage.Shapes
        If tShape = "SETKA KOORD" Then DeleteExistBlock tShape
        Next
End Sub

Public Sub UpdateZoneInDoc() 'Обновить сетку коодинат на всех листах
    Dim tShape As Visio.Shape
    Dim tPage As Visio.Page
    For Each tPage In ActiveDocument.Pages
        For Each tShape In ActivePage.Shapes
            If tShape = "SETKA KOORD" Then DeleteExistBlock tShape
        Next
    Next
End Sub

Private Sub DeleteExistBlock(shpActive As Visio.Shape) 'Удаляем существующие ячейки зон начиная с В... и с 2...
    Dim tShape As Visio.Shape
    Dim strName As String
    Dim colShapes As New Collection
    Dim Index As Integer

    'make sure all cells are spaced according to settings set in document options
    For Each tShape In shpActive.Shapes
        strName = tShape.NameU
        If strName = "Zone" Or strName = "Zones" Then 'Our title block zones nameu is "Zone" but user title block zones nameu is "Zones"
            Set colShapes = New Collection
            'remove all extra shape before being generated again
            For Index = 1 To tShape.Shapes.Count
                If InStr(1, tShape.Shapes(Index).NameU, "HZone1") > 0 Or InStr(1, tShape.Shapes(Index).NameU, "VZone1") > 0 Then
                    'add into collection before remove as all indexs will change if deleted immediately
                    If Len(tShape.Shapes(Index).NameU) <> 6 Then colShapes.Add tShape.Shapes(Index)
                End If
            Next
            'delete the shapes
            For Index = 1 To colShapes.Count
                colShapes(Index).Delete
            Next
            'generate the shapes
            CreateZoneBlocks tShape
        End If
    Next
End Sub

Private Sub CreateZoneBlocks(shpActive As Visio.Shape)   'Копируем начальные блоки и задаем их ширину
    Dim Remain As Double
    Dim LastShape As Visio.Shape
    Dim NewShape As Visio.Shape
    Dim Spacing As Double
    Dim Zone As String
    Dim Gutter As Double
    Dim curScale As Double
    Const conZoneGutterMm As Double = 0.196850393700787  '5 mm
    Const conZoneGutterIn As Double = 0.25 '6,35 mm

    On Error GoTo Skip
    curScale = ActivePage.PageSheet.CellsU("DrawingScale") / ActivePage.PageSheet.CellsU("PageScale")
    Gutter = IIf(shpActive.CellsU("Width").Units = visMillimeters, conZoneGutterMm, conZoneGutterIn) * curScale
    'read LocC and set width of first h zone
    shpActive.Shapes("HZone1").CellsU("Width").Formula = ActiveDocument.DocumentSheet.CellsU("User.LocationC") * curScale - Gutter - (0.78740157480315 - ActiveDocument.DocumentSheet.CellsU("User.OffsetFrame")) '20??=0.78740157480315
    'read LocA and set width of first v zone
    shpActive.Shapes("VZone1").CellsU("Width").Formula = ActiveDocument.DocumentSheet.CellsU("User.LocationA") * curScale - ActiveDocument.DocumentSheet.CellsU("User.OffsetFrame")
    'insert h zones
    Remain = shpActive.CellsU("Width").ResultIU - shpActive.Shapes("HZone1").CellsU("Width").ResultIU
    Set LastShape = shpActive.Shapes("HZone1")
    Spacing = ActiveDocument.DocumentSheet.CellsU("User.LocationD")
    Do While Remain > 0
        If Remain >= Spacing * curScale Then
            Set NewShape = shpActive.Drop(LastShape, 0, 0)
            NewShape.CellsU("Width").Formula = Spacing * curScale
            Remain = Remain - NewShape.CellsU("Width").ResultIU
            NewShape.CellsU("PinX").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinX +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
            NewShape.CellsU("PinY").FormulaForceU = "GUARD(IF(" + shpActive.NameID + "!Scratch.A1=0," + shpActive.NameID + "!Height-Height*0.5,Height*0.5))"
            Set LastShape = NewShape
        Else
            If Abs(Remain) < LastShape.CellsU("Height").ResultIU Then
                LastShape.CellsU("Width").Formula = LastShape.CellsU("Width").ResultIU + Abs(Remain)
            Else
                Set NewShape = shpActive.Drop(LastShape, 0, 0)
                NewShape.CellsU("Width").Formula = Abs(Remain)
                NewShape.CellsU("PinX").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinX +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
                NewShape.CellsU("PinY").FormulaForceU = "GUARD(IF(" + shpActive.NameID + "!Scratch.A1=0," + shpActive.NameID + "!Height-Height*0.5,Height*0.5))"
            End If
            Remain = 0
        End If
        DoEvents
    Loop
    'insert v zones
    Remain = shpActive.CellsU("Height").ResultIU - shpActive.Shapes("VZone1").CellsU("Width").ResultIU
    Set LastShape = shpActive.Shapes("VZone1")
    LastShape.CellsU("TxtAngle").FormulaU = "IF(" + shpActive.NameID + "!Scratch.C1=0, 0 deg, 270 deg)"
    Spacing = ActiveDocument.DocumentSheet.CellsU("User.LocationB")
    Do While Remain > 0
        If Remain >= Spacing * curScale Then
            Set NewShape = shpActive.Drop(LastShape, 0, 0)
            NewShape.CellsU("Width").Formula = Spacing * curScale
            Remain = Remain - NewShape.CellsU("Width").ResultIU
            NewShape.CellsU("PinY").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinY +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
            NewShape.CellsU("PinX").FormulaForceU = "GUARD(IF(" + shpActive.NameID + "!Scratch.B1=1," + shpActive.NameID + "!Width-Height*0.5,Height*0.5))"
            NewShape.CellsU("TxtAngle").FormulaU = "IF(" + shpActive.NameID + "!Scratch.C1=0, 0 deg, 270 deg)"
            Set LastShape = NewShape
        Else
            If Abs(Remain) < LastShape.CellsU("Height").ResultIU Then
                LastShape.CellsU("Width").Formula = LastShape.CellsU("Width").ResultIU + Abs(Remain)
            Else
                Set NewShape = shpActive.Drop(LastShape, 0, 0)
                NewShape.CellsU("Width").Formula = Abs(Remain)
                NewShape.CellsU("PinY").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinY +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
                NewShape.CellsU("PinX").FormulaForceU = "GUARD(IF(" + shpActive.NameID + "!Scratch.B1=1," + shpActive.NameID + "!Width-Height*0.5,Height*0.5))"
                NewShape.CellsU("TxtAngle").FormulaU = "IF(" + shpActive.NameID + "!Scratch.C1=0, 0 deg, 270 deg)"
            End If
            Remain = 0
        End If
        DoEvents
    Loop
Skip:
End Sub


Sub ReadCopyRight()
    Debug.Print ActiveWindow.Selection(1).Cells("Copyright").FormulaU
End Sub
Sub RegCopyright()
    On Error GoTo EMSG
    ActiveWindow.Selection(1).Cells("Copyright").FormulaU = Chr(34) & "Copyright (C) 2009 Visio Guys" & Chr(34)
    Exit Sub
EMSG:
    MsgBox err.Description
End Sub
Sub RegAllCopyright()
    Dim shp As Visio.Shape
    On Error GoTo EMSG
    For Each shp In ActivePage.Shapes
        shp.Cells("Copyright").FormulaU = Chr(34) & "Copyright (C) 2009 Visio Guys" & Chr(34)
    Next
    Exit Sub
EMSG:
    MsgBox err.Description
End Sub



