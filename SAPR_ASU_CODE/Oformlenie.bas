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
    OffsetFrame = ActiveDocument.DocumentSheet.CellsU("User.SA_FR_OffsetFrame")
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
    Set vsoLayer1 = Application.ActiveWindow.Page.Layers("SA_Рамка")
    
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
    Application.TemplatePaths = "D:\YandexDisk\VISIO\SAPR_ASU"
    Application.StencilPaths = "D:\YandexDisk\VISIO\SAPR_ASU"
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

    ПеределкаСтандартныхСтилей
    
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

Sub ПеределкаСтандартныхСтилей()
    Dim vsoStyle As Visio.style

    For i = 1 To Application.ActiveDocument.Styles.Count
        Set vsoStyle = Application.ActiveDocument.Styles.ItemFromID(i)
        If vsoStyle.NameU = "No Style" Or _
            vsoStyle.NameU = "Text Only" Or _
            vsoStyle.NameU = "None" Or _
            vsoStyle.NameU = "Normal" Or _
            vsoStyle.NameU = "Guide" _
        Then
            With vsoStyle
                .CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaU = 112
                .CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = 2
                .CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "11 pt"
                .CellsSRC(visSectionCharacter, 0, visCharacterDblUnderline).FormulaU = False
                .CellsSRC(visSectionCharacter, 0, visCharacterOverline).FormulaU = False
                .CellsSRC(visSectionCharacter, 0, visCharacterStrikethru).FormulaU = False
                .CellsSRC(visSectionCharacter, 0, 11).FormulaU = False
                .CellsSRC(visSectionCharacter, 0, visCharacterDoubleStrikethrough).FormulaU = False
                .CellsSRC(visSectionCharacter, 0, visCharacterRTLText).FormulaU = False
                .CellsSRC(visSectionCharacter, 0, visCharacterUseVertical).FormulaU = False
                .CellsSRC(visSectionObject, visRowText, visTxtBlkTopMargin).FormulaU = "0 pt"
                .CellsSRC(visSectionObject, visRowText, visTxtBlkBottomMargin).FormulaU = "0 pt"
                .CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.2 mm"
                .CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = 1
            End With
        End If
    Next
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
    Dim MaxNumber As Integer
    
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
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
        ActiveDocument.Masters.Item("Рамка").Delete
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(0,Prop.CHAPTER.Format)"

    Else
        'Ищем номер последней страницы
        MaxNumber = MaxPageNumber(colPages)

        If MaxNumber = 0 Then
            'Создаем вторую страницу
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.Name = PageName & ".2"
        Else
            'Создаем последующие страницы
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.Name = PageName & "." & CStr(MaxNumber + 1)
        End If
        
        Set shpRamka = vsoPage.Drop(Ramka, 0, 0)
        ActiveDocument.Masters.Item("Рамка").Delete
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"

        
    End If
    shpRamka.Cells("Prop.CNUM") = 0
    shpRamka.Cells("Prop.TNUM") = 0
    vsoPage.PageSheet.Cells("PageWidth").Formula = "420 MM"
    vsoPage.PageSheet.Cells("PageHeight").Formula = "297 MM"
    vsoPage.PageSheet.Cells("Paperkind").Formula = 8
    vsoPage.PageSheet.Cells("PrintPageOrientation").Formula = 2
    
    LockTitleBlock

    Set AddSAPage = vsoPage
    
End Function


Sub fff()
Dim eee As Page
Set eee = AddSAPage("PageName")
End Sub


Sub AddSAPageNext()
'------------------------------------------------------------------------------------------------------------
' Sub           : AddSAPageNext - Добавляет страницу САПР-АСУ за текущей, копируя ее свойства
                'Переименовывает страницы раздела идущие за вставляемой страницей
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim vsoPageNew As Visio.Page
    Dim colPagesAll As Collection
    Dim colPagesAfter As Collection
    Dim Ramka As Visio.Master
    Dim shpRamka As Visio.Shape
    Dim MaxNpageU As Integer
    Dim MaxNpage As Integer
    Dim NameActivePage As String
    Dim PageName As String
    Dim PageNumber As Integer
    Dim Index As Integer
    Dim ItemCol As Integer
    
    Set colPagesAfter = New Collection
    Set colPagesAll = New Collection
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    NameActivePage = ActivePage.Name
    Index = ActivePage.Index
    PageName = GetPageName(NameActivePage)
    PageNumber = GetPageNumber(NameActivePage)
    
    'Ищем страницы раздела больше текущей
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            colPagesAll.Add vsoPage
            If GetPageNumber(vsoPage.Name) > PageNumber Then
                colPagesAfter.Add vsoPage
            End If
        End If
    Next
    
    'Если вставляем страницу в середину раздела
    If colPagesAfter.Count > 0 Then
        'Сдвигаем = Переименовываем все листы ниже текущего : к номеру последнего прибавляем + 1
        While colPagesAfter.Count > 0
            ItemCol = FindPageMaxNumber(colPagesAfter)
            Set vsoPage = colPagesAfter.Item(ItemCol)
            colPagesAfter.Remove ItemCol
            vsoPage.Name = PageName & "." & CStr(GetPageNumber(vsoPage.Name) + 1)
        Wend
    End If
    
    'Находим максимальный номер страницы в NameU и Name
    MaxNpage = MaxPageNumber(colPagesAll)
    MaxNpageU = MaxPageNumberU(colPagesAll)
    'Создаем страницу раздела с максимальным номером
    Set vsoPageNew = ActiveDocument.Pages.Add
    vsoPageNew.Name = PageName & "." & CStr(IIf(MaxNpage > MaxNpageU, MaxNpage, MaxNpageU) + 1)
    'Переименовываем вставленный лист в нумерацию Name после текущего
    vsoPageNew.Name = PageName & "." & CStr(PageNumber + 1)
    'Положение новой страницы сразу за текущей
    vsoPageNew.Index = Index + 1
    Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
    ActiveDocument.Masters.Item("Рамка").Delete
    shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"
    shpRamka.Cells("Prop.CNUM") = 0
    shpRamka.Cells("Prop.TNUM") = 0
    vsoPageNew.PageSheet.Cells("PageWidth").Formula = "420 MM"
    vsoPageNew.PageSheet.Cells("PageHeight").Formula = "297 MM"
    vsoPageNew.PageSheet.Cells("Paperkind").Formula = 8
    vsoPageNew.PageSheet.Cells("PrintPageOrientation").Formula = 2
    
    LockTitleBlock

End Sub

Function GetPageName(NamePage As String) As String
    Dim mstrName() As String
    mstrName = Split(NamePage, ".")
    GetPageName = mstrName(0)
End Function

Function GetPageNumber(NamePage As String) As Integer
    Dim mstrName() As String
    mstrName = Split(NamePage, ".")
    If UBound(mstrName) > 0 Then GetPageNumber = CInt(mstrName(1)) Else GetPageNumber = 1
End Function

Function FindPageMaxNumber(colPages As Collection) As Integer
    Dim vsoPage As Visio.Page
    Dim vsoPageMax As Visio.Page
    Dim MaxNumber As Integer
    Dim Npage As Integer
    Dim i As Integer
    Dim ItemCol As Integer
    For i = 1 To colPages.Count
        Npage = GetPageNumber(colPages.Item(i).Name)
        If Npage > MaxNumber Then MaxNumber = Npage: ItemCol = i
    Next
    FindPageMaxNumber = ItemCol
End Function

Function MaxPageNumber(colPages As Collection) As Integer
    Dim vsoPage As Visio.Page
    Dim MaxNumber As Integer
    Dim Npage As Integer
    For Each vsoPage In colPages
        Npage = GetPageNumber(vsoPage.Name)
        If Npage > MaxNumber Then MaxNumber = Npage
    Next
    MaxPageNumber = MaxNumber
End Function

Function MaxPageNumberU(colPages As Collection) As Integer
    Dim vsoPage As Visio.Page
    Dim MaxNumber As Integer
    Dim Npage As Integer
    For Each vsoPage In colPages
        Npage = GetPageNumber(vsoPage.NameU)
        If Npage > MaxNumber Then MaxNumber = Npage
    Next
    MaxPageNumberU = MaxNumber
End Function