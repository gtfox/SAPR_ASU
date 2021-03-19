'------------------------------------------------------------------------------------------------------------
' Module        : Oformlenie - Сетка координат зон чертежа, блокировка рамки, стили оформления, страницы
' Author        : gtfox
' Date          : 2020.05.05
' Description   : Сборник макросов относящихся к оформлению
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------


Public Sub UpdateZoneOnPage()
'------------------------------------------------------------------------------------------------------------
' Macros        : UpdateZoneOnPage - Обновить сетку коодинат на листе
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Visio.Shape
    For Each vsoShape In ActivePage.Shapes
        If vsoShape.Name Like "SETKA KOORD*" Then UpdateZoneBlocks vsoShape
        Next
End Sub

Public Sub UpdateZoneInDoc()
'------------------------------------------------------------------------------------------------------------
' Macros        : UpdateZoneInDoc - Обновить сетку коодинат на всех листах
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Visio.Shape
    Dim vsoPage As Visio.Page
    For Each vsoPage In ActiveDocument.Pages
        For Each vsoShape In ActivePage.Shapes
            If vsoShape.Name Like "SETKA KOORD*" Then UpdateZoneBlocks vsoShape
        Next
    Next
End Sub

Private Sub UpdateZoneBlocks(shpSetkaKoord As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : UpdateZoneBlocks - Формирует сетку координат зон чертежа из вброшенного шейпа SETKA KOORD
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Visio.Shape
    Dim LastShape As Visio.Shape
    Dim NewShape As Visio.Shape
    Dim OffsetFrame As Double
    Dim Ostalos As Double
    Dim ShirinaZony As Double
    Dim PageScale As Double
    Dim Name As String
    Dim i As Integer
    Dim colShapes As New Collection
    Set colShapes = New Collection
    Const RamkaLevo As Double = 20 / 25.4 '20 mm
    Const RamkaPravo As Double = 5 / 25.4 '5 mm

    'Удаляем существующие ячейки зон начиная с В... и с 2...
    
    'Ищем все блоки с именами больше 5 символов
    For Each vsoShape In shpSetkaKoord.Shapes
        
        If InStr(1, vsoShape.NameU, "HZone") > 0 Or InStr(1, vsoShape.NameU, "VZone") > 0 Then
            If Len(vsoShape.NameU) <> 5 Then colShapes.Add vsoShape
        End If
    Next
    
    'Удаляем все кроме первых блоков
    For i = 1 To colShapes.Count
        colShapes(i).Delete
    Next
    
    'Копируем начальные блоки и задаем их ширину

    PageScale = ActivePage.PageSheet.CellsU("DrawingScale") / ActivePage.PageSheet.CellsU("PageScale")
    OffsetFrame = ActiveDocument.DocumentSheet.CellsU("User.OffsetFrame")
    shpSetkaKoord.Shapes("HZone").CellsU("Width").Formula = ActiveDocument.DocumentSheet.CellsU("User.SA_Pole1") * PageScale - RamkaPravo - RamkaLevo + OffsetFrame
    shpSetkaKoord.Shapes("VZone").CellsU("Width").Formula = ActiveDocument.DocumentSheet.CellsU("User.SA_PoleA") * PageScale - OffsetFrame
    
    'Вставляем горизонтальные блоки
    Ostalos = shpSetkaKoord.CellsU("Width").ResultIU - shpSetkaKoord.Shapes("HZone").CellsU("Width").ResultIU
    Set LastShape = shpSetkaKoord.Shapes("HZone")
    ShirinaZony = ActiveDocument.DocumentSheet.CellsU("User.SA_PoleGor")
    Do While Ostalos > 0
        If Ostalos >= ShirinaZony * PageScale Then
            Set NewShape = shpSetkaKoord.Drop(LastShape, 0, 0)
            NewShape.CellsU("Width").Formula = ShirinaZony * PageScale
            Ostalos = Ostalos - NewShape.CellsU("Width").ResultIU
            NewShape.CellsU("PinX").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinX +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
            NewShape.CellsU("PinY").FormulaForceU = "GUARD(" + shpSetkaKoord.NameID + "!Height-Height*0.5)"
            Set LastShape = NewShape
        Else
            If Abs(Ostalos) < LastShape.CellsU("Height").ResultIU Then
                LastShape.CellsU("Width").Formula = LastShape.CellsU("Width").ResultIU + Abs(Ostalos)
            Else
                Set NewShape = shpSetkaKoord.Drop(LastShape, 0, 0)
                NewShape.CellsU("Width").Formula = Abs(Ostalos)
                NewShape.CellsU("PinX").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinX +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
                NewShape.CellsU("PinY").FormulaForceU = "GUARD(" + shpSetkaKoord.NameID + "!Height-Height*0.5)"
            End If
            Ostalos = 0
        End If
        DoEvents
    Loop
    
    'Вставляем вертикальные блоки
    Ostalos = shpSetkaKoord.CellsU("Height").ResultIU - shpSetkaKoord.Shapes("VZone").CellsU("Width").ResultIU
    Set LastShape = shpSetkaKoord.Shapes("VZone")
    LastShape.CellsU("TxtAngle").FormulaU = "IF(" + shpSetkaKoord.NameID + "!Scratch.C1=0, 0 deg, 270 deg)"
    ShirinaZony = ActiveDocument.DocumentSheet.CellsU("User.SA_PoleVert")
    Do While Ostalos > 0
        If Ostalos >= ShirinaZony * PageScale Then
            Set NewShape = shpSetkaKoord.Drop(LastShape, 0, 0)
            NewShape.CellsU("Width").Formula = ShirinaZony * PageScale
            Ostalos = Ostalos - NewShape.CellsU("Width").ResultIU
            NewShape.CellsU("PinY").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinY +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
            Set LastShape = NewShape
        Else
            If Abs(Ostalos) < LastShape.CellsU("Height").ResultIU Then
                LastShape.CellsU("Width").Formula = LastShape.CellsU("Width").ResultIU + Abs(Ostalos)
            Else
                Set NewShape = shpSetkaKoord.Drop(LastShape, 0, 0)
                NewShape.CellsU("Width").Formula = Abs(Ostalos)
                NewShape.CellsU("PinY").FormulaForceU = "GUARD(" + LastShape.NameID + "!PinY +" + LastShape.NameID + "!Width * 0.5 + width *0.5)"
            End If
            Ostalos = 0
        End If
        DoEvents
    Loop

    Set colShapes = Nothing
    
End Sub

Public Sub LockTitleBlock()
'------------------------------------------------------------------------------------------------------------
' Macros        : LockTitleBlock - Блокировка слоя рамки
'------------------------------------------------------------------------------------------------------------

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


Sub DrawingPageColor()
'------------------------------------------------------------------------------------------------------------
' Macros        : DrawingPageColor - Цвет листа как Splan 7 (15924991-кремовый)
'------------------------------------------------------------------------------------------------------------

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

Private Sub SetStyleGost()
'------------------------------------------------------------------------------------------------------------
' Macros        : SetStyleGost - Изменение стилей под Гост
'------------------------------------------------------------------------------------------------------------

    Dim vsoStyle As Visio.style
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

Private Sub Pole_Spec()
'------------------------------------------------------------------------------------------------------------
' Macros        : Pole_Spec - Массовая установка свойств в спецификации (поля 10pt до и 5/1pt  после текста во 2 и 9 столбце )
'------------------------------------------------------------------------------------------------------------

    For i = 1 To 30
    
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".2").CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "10 pt"
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".2").CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "10 pt"
        
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".9").CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "5 pt"
        Application.ActiveWindow.Page.Shapes("Спецификация").Shapes("row" & i).Shapes(i & ".9").CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "1 pt"
    Next i
End Sub

Private Sub Pole_VRCh()
'------------------------------------------------------------------------------------------------------------
' Macros        : Pole_VRCh - Установка полей текста ВРЧ 10pt
'------------------------------------------------------------------------------------------------------------

For i = 1 To 15
    Application.ActiveWindow.Page.Shapes("В Р Ч").Shapes(i).Shapes(3).CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "10 pt"
Next i

End Sub

Function AddSAPage(PageName As String) As Visio.Page
'------------------------------------------------------------------------------------------------------------
' Function        : AddSAPage - Добавляет страницу САПР-АСУ если ее нет, добавляет еще, если подобные уже есть
                  'В строке передается имя страницы, возвращаем что создали
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim colPages As Collection
    Dim Ramka As Visio.Master
    Dim shpRamka As Visio.Shape
    Dim Npage As Integer
    Dim MaxNumber As Double
    
    Set Ramka = Application.Documents.Item("SAPR_ASU_SHAPE.vss").Masters.Item("Рамка")  'ActiveDocument.Masters.Item("Рамка")
    Set colPages = New Collection
    
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            colPages.Add vsoPage
        End If
    Next
    
    If colPages.Count = 0 Then
        'Создаем первую страницу
        Set vsoPage = ActiveDocument.Pages.Add
        vsoPage.Name = PageName
        Set shpRamka = vsoPage.Drop(Ramka, 0, 0)
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(0,Prop.CHAPTER.Format)"


    Else
        'Ищем номер последней страницы
        For Each vsoPage In colPages
            Npage = CDbl(IIf(Mid(vsoPage.Name, Len(PageName) + 2) = "", 0, Mid(vsoPage.Name, Len(PageName) + 2)))
            If Npage > MaxNumber Then MaxNumber = Npage
        Next
        
        If Npage = 0 Then
            'Создаем вторую страницу
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.Name = PageName & ".2"
        Else
            'Создаем последующие страницы
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.Name = PageName & "." & CStr(Npage + 1)
        End If
        
        Set shpRamka = vsoPage.Drop(Ramka, 0, 0)
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"

        
    End If
    shpRamka.Cells("Prop.CNUM") = 0
    shpRamka.Cells("Prop.TNUM") = 0
    vsoPage.PageSheet.Cells("PageWidth").Formula = "420 MM"
    vsoPage.PageSheet.Cells("PageHeight").Formula = "297 MM"
    vsoPage.PageSheet.Cells("Paperkind").Formula = 9
    vsoPage.PageSheet.Cells("PrintPageOrientation").Formula = 1
    
    LockTitleBlock

    Set AddSAPage = vsoPage
    
End Function



