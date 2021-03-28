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


Sub SetSAStyle()
    SetVisioProp
    SetGridSnap
    SetDefStyleISOCPEUR11
    SetPanel
End Sub

Sub SetVisioProp()
'------------------------------------------------------------------------------------------------------------
' Macros        : SetVisioProp - Настройки Visio, Цвет листа как Splan 7 (15924991-кремовый)
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

Private Sub SetGridSnap()
'------------------------------------------------------------------------------------------------------------
' Macros        : SetGridSnap - Изменение сетки и силы привязки
'------------------------------------------------------------------------------------------------------------
    Dim vsoShape As Shape
    Dim vsoPage As Visio.Page
    
    'сетка 2,5 мм
    For Each vsoPage In Application.ActiveDocument.Pages
        Set vsoShape = vsoPage.PageSheet
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visXGridDensity).FormulaU = "0"
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visXGridSpacing).FormulaU = "2.5 mm"
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visYGridDensity).FormulaU = "0"
        vsoShape.CellsSRC(visSectionObject, visRowRulerGrid, visYGridSpacing).FormulaU = "2.5 mm"
    Next

    'Сила привязки к сетке в пикселях
    'Сервис -> Привязать и приклеить -> Дополнительно -> Сетка = 100
    Application.Settings.SnapStrengthGridX = 30
    Application.Settings.SnapStrengthGridY = 30
    
End Sub

Sub SetDefStyleISOCPEUR11()
'------------------------------------------------------------------------------------------------------------
' Macros        : SetDefStyleISOCPEUR11 - Изменение стандартные стили на ISOCPEUR 11pt
'------------------------------------------------------------------------------------------------------------
    Dim vsoStyle As Visio.style

    For i = 1 To Application.ActiveDocument.Styles.Count
        Set vsoStyle = Application.ActiveDocument.Styles.ItemFromID(i)
        If vsoStyle.NameU = "Normal" Then
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

Sub SetPanel()
    Application.CommandBars("Standard").Visible = True
    Application.CommandBars("Formatting").Visible = True
    Application.CommandBars("View").Visible = True
    Application.CommandBars("Action").Visible = True
    Application.CommandBars("Stop Recording").Visible = True
    Application.CommandBars("Snap & Glue").Visible = True
    Application.CommandBars("Developer").Visible = True
    Application.CommandBars("Drawing").Visible = True
    Application.CommandBars("Format Text").Visible = True
    Application.CommandBars("Format Shape").Visible = True
    Application.CommandBars("Reviewing").Visible = False
    Application.CommandBars("Web").Visible = False
    Application.CommandBars("Ink").Visible = False
    Application.CommandBars("Stencil").Visible = False
    Application.CommandBars("Picture").Visible = False
    Application.CommandBars("Layout & Routing").Visible = False
    Application.CommandBars("Data").Visible = False
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
    Dim MaxNpage As Integer
    
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
        MaxNumber = MaxMinPageNumber(colPages)

        If MaxNumber = 1 Then 'Создаем вторую страницу
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.Name = PageName & ".2"
            
        Else 'Создаем последующие страницы
            'Находим максимальный номер страницы в NameU и Name
            MaxNpage = MaxMinPageNumber(colPages, , , True)
            'Создаем страницу раздела с максимальным номером
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.Name = PageName & "." & CStr(MaxNpage + 1)
            'Переименовываем вставленный лист в нумерацию Name после текущего
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

Sub ShowSAPageRazdel()
    frmPageAddRazdel.Show
End Sub

Sub AddSAPageNext()
'------------------------------------------------------------------------------------------------------------
' Sub           : AddSAPageNext - Добавляет страницу САПР-АСУ за текущей, копируя ее свойства
                'Переименовывает страницы раздела идущие за вставляемой страницей
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim vsoPageNew As Visio.Page
    Dim vsoPageSource As Visio.Page
    Dim colPagesAll As Collection
    Dim colPagesAfter As Collection
    Dim Ramka As Visio.Master
    Dim Setka As Visio.Master
    Dim shpRamka As Visio.Shape
    Dim shpRamkaSource As Visio.Shape
    Dim MaxNpage As Integer
    Dim PageName As String
    Dim PageNumber As Integer
    Dim Index As Integer
    Dim ItemCol As Integer
    Dim NazvanieShemy As String
    Dim NazvanieFSA As String
    
    Set colPagesAfter = New Collection
    Set colPagesAll = New Collection
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    Set Setka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("SETKA KOORD")
    Set vsoPageSource = ActivePage
    Index = vsoPageSource.Index
    PageName = GetPageName(vsoPageSource.Name)
    PageNumber = GetPageNumber(vsoPageSource.Name)
    Set shpRamkaSource = vsoPageSource.Shapes("Рамка")

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
    'Сдвигаем = Переименовываем все листы ниже текущего : к номеру последнего прибавляем + 1
    While colPagesAfter.Count > 0
        ItemCol = FindPageMaxMinNumber(colPagesAfter)
        Set vsoPage = colPagesAfter.Item(ItemCol)
        colPagesAfter.Remove ItemCol
        vsoPage.Name = PageName & "." & CStr(GetPageNumber(vsoPage.Name) + 1) & IIf(GetPageDesc(vsoPage.Name) = "", "", "." & GetPageDesc(vsoPage.Name))
    Wend
    
    'Находим максимальный номер страницы в NameU и Name
    MaxNpage = MaxMinPageNumber(colPagesAll, , , True)
    'Создаем страницу раздела с максимальным номером
    Set vsoPageNew = ActiveDocument.Pages.Add
    vsoPageNew.Name = PageName & "." & CStr(MaxNpage + 1)
    'Переименовываем вставленный лист в нумерацию Name после текущего
    vsoPageNew.Name = PageName & "." & CStr(PageNumber + 1)
    'Положение новой страницы сразу за текущей
    vsoPageNew.Index = Index + 1
    Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
    ActiveDocument.Masters.Item("Рамка").Delete
    shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"
    shpRamka.Cells("Prop.Type").Formula = shpRamkaSource.Cells("Prop.Type").Formula
    shpRamka.Cells("Prop.CNUM").Formula = shpRamkaSource.Cells("Prop.CNUM").Formula
    shpRamka.Cells("Prop.TNUM").Formula = shpRamkaSource.Cells("Prop.TNUM").Formula
    vsoPageNew.PageSheet.Cells("PageWidth").Formula = vsoPageSource.PageSheet.Cells("PageWidth").Formula
    vsoPageNew.PageSheet.Cells("PageHeight").Formula = vsoPageSource.PageSheet.Cells("PageHeight").Formula
    vsoPageNew.PageSheet.Cells("Paperkind").Formula = vsoPageSource.PageSheet.Cells("Paperkind").Formula
    vsoPageNew.PageSheet.Cells("PrintPageOrientation").Formula = vsoPageSource.PageSheet.Cells("PrintPageOrientation").Formula
    If vsoPageSource.PageSheet.CellExists("Prop.SA_NazvanieShemy", 0) Then
        SetNazvanieShemy vsoPageNew.PageSheet
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").Formula
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShemy").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieShemy").Formula
        vsoPageNew.Drop Setka, 0, 0
    End If
    If vsoPageSource.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        SetNazvanieFSA vsoPageNew.PageSheet
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula
    End If
    
    LockTitleBlock

    ActiveWindow.DeselectAll
    
End Sub

Sub DelSAPage()
'------------------------------------------------------------------------------------------------------------
' Sub           : DelSAPage - Удаляет текущую страницу САПР-АСУ
                'Переименовывает страницы раздела идущие после удаленной страницы
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim colPagesAfter As Collection
    Dim NameActivePage As String
    Dim PageName As String
    Dim PageNumber As Integer
    Dim ItemCol As Integer
    
    If MsgBox("Удалить лист: " & ActivePage.Name, vbYesNo + vbCritical, "Удаление листа") = vbYes Then
    
        Set colPagesAfter = New Collection
        NameActivePage = ActivePage.Name
        PageName = GetPageName(NameActivePage)
        PageNumber = GetPageNumber(NameActivePage)

        ActiveWindow.DeselectAll
        On Error GoTo err
        ActiveWindow.SelectAll
        ActiveWindow.Selection.Delete
        
        DoEvents
err:
        ActivePage.Delete 0
        
        If PageNumber = 1 Then Exit Sub
        
        'Ищем страницы раздела больше текущей
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.Name Like PageName & "*" Then
                If GetPageNumber(vsoPage.Name) > PageNumber Then
                    colPagesAfter.Add vsoPage
                End If
            End If
        Next
        
        'Если удаляем страницу из середины раздела
        'Сдвигаем = Переименовываем все листы ниже текущего : у номера первого - 1
        While colPagesAfter.Count > 0
            ItemCol = FindPageMaxMinNumber(colPagesAfter, True)
            Set vsoPage = colPagesAfter.Item(ItemCol)
            colPagesAfter.Remove ItemCol
            vsoPage.Name = PageName & "." & CStr(GetPageNumber(vsoPage.Name) - 1) & IIf(GetPageDesc(vsoPage.Name) = "", "", "." & GetPageDesc(vsoPage.Name))
        Wend
        
    End If
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

Function GetPageDesc(NamePage As String) As String
    Dim mstrName() As String
    mstrName = Split(NamePage, ".")
    If UBound(mstrName) > 1 Then GetPageDesc = mstrName(2) Else GetPageDesc = ""
End Function

Function FindPageMaxMinNumber(colPages As Collection, Optional Min As Boolean) As Integer
    Dim vsoPage As Visio.Page
    Dim vsoPageMax As Visio.Page
    Dim MaxNumber As Integer
    Dim MinNumber As Integer
    Dim Npage As Integer
    Dim i As Integer
    Dim ItemCol As Integer
    MinNumber = 32767
    For i = 1 To colPages.Count
        Npage = GetPageNumber(colPages.Item(i).Name)
        If Min Then
            If Npage < MinNumber Then MinNumber = Npage: ItemCol = i
        Else
            If Npage > MaxNumber Then MaxNumber = Npage: ItemCol = i
        End If
    Next
    FindPageMaxMinNumber = ItemCol
End Function

Function MaxMinPageNumber(colPages As Collection, Optional Min As Boolean, Optional NameU As Boolean, Optional AllName As Boolean) As Integer
    Dim vsoPage As Visio.Page
    Dim MaxNumber As Integer
    Dim MinNumber As Integer
    Dim MaxNumberTemp As Integer
    Dim MinNumberTemp As Integer
    Dim Npage As Integer
    
    GoSub SubFind
    MaxMinPageNumber = IIf(Min, MinNumber, MaxNumber)

    If AllName Then
        NameU = Not NameU
        MaxNumberTemp = MaxNumber
        MinNumberTemp = MinNumber
        GoSub SubFind
        If Min Then
            MaxMinPageNumber = IIf(MinNumber < MinNumberTemp, MinNumber, MinNumberTemp)
        Else
            MaxMinPageNumber = IIf(MaxNumber > MaxNumberTemp, MaxNumber, MaxNumberTemp)
        End If
    End If
    Exit Function
    
SubFind:
    MinNumber = 32767
    For Each vsoPage In colPages
        Npage = GetPageNumber(IIf(NameU, vsoPage.NameU, vsoPage.Name))
        If Npage < MinNumber Then MinNumber = Npage
        If Npage > MaxNumber Then MaxNumber = Npage
    Next
    Return
End Function

Sub SetNazvanieShemy(vsoObject As Object) 'SetValueToSelSections
    Dim arrRowValue()
    Dim arrRowName()
    Dim SectionNumber As Long
    SectionNumber = visSectionProp 'Prop 243
    arrRowName = Array("SA_NazvanieShemy")
    arrRowValue = Array("""Название Схемы"":""Нумерация элементов идет в пределах одной схемы"":1:"""":INDEX(0,Prop.SA_NazvanieShemy.Format):"""":FALSE:FALSE:1049:0")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber
End Sub

Sub SetNazvanieFSA(vsoObject As Object) 'SetValueToSelSections
    Dim arrRowValue()
    Dim arrRowName()
    Dim SectionNumber As Long
    SectionNumber = visSectionProp 'Prop 243
    arrRowName = Array("SA_NazvanieFSA")
    arrRowValue = Array("""Название ФСА"":""Нумерация элементов идет в пределах одной ФСА"":1:"""":INDEX(0,Prop.SA_NazvanieFSA.Format):"""":FALSE:FALSE:1049:0")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber
End Sub