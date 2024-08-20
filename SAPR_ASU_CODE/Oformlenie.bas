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
        If vsoShape.name Like "SETKA KOORD*" Then UpdateZoneBlocks vsoShape
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
            If vsoShape.name Like "SETKA KOORD*" Then UpdateZoneBlocks vsoShape
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
    Dim name As String
    Dim i As Integer
    Dim colShapes As New Collection
    Set colShapes = New Collection
    Const RamkaLevo As Double = 20 / 25.4 '20 mm
    Const RamkaPravo As Double = 5 / 25.4 '5 mm
    Dim AppEventsEnabled As Boolean
    
    AppEventsEnabled = Application.EventsEnabled
    Application.EventsEnabled = False
    
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

    Application.EventsEnabled = AppEventsEnabled
    
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

Public Sub LockSelected()
'------------------------------------------------------------------------------------------------------------
' Macros        : LockSelected - Блокировка выделенных объектов
'------------------------------------------------------------------------------------------------------------
    Dim vsoLayer1 As Visio.Layer
    Dim vsoShape As Visio.Shape
    
    If Application.ActiveWindow.Selection.Count > 0 Then
        If MsgBox("Заблокировать выделененые объекты: " & Application.ActiveWindow.Selection.Count & "шт.?", vbExclamation + vbOKCancel, "САПР-АСУ: Блокировки выделенного объекта") = vbOK Then
            'Создаем и блокруем слой
            Set vsoLayer1 = Application.ActiveWindow.Page.Layers.Add("SA_LockedLayer")
            Set vsoLayer2 = Application.ActiveWindow.Page.Layers.Add("SA_LockedWire")
'            SetLayer Application.ActiveWindow.Selection(1), vsoLayer1
            For Each vsoShape In Application.ActiveWindow.Selection
                If ShapeSATypeIs(vsoShape, typeCxemaWire) Then
                    vsoLayer2.Add vsoShape, 0
                Else
                    vsoLayer1.Add vsoShape, 0
                End If
            Next
            vsoLayer1.CellsC(visLayerLock).FormulaU = "1"
            vsoLayer1.CellsC(visLayerColor).FormulaU = "19"
            vsoLayer1.CellsC(visLayerSnap).FormulaU = "0"
            vsoLayer1.CellsC(visLayerGlue).FormulaU = "0"
            
            vsoLayer2.CellsC(visLayerLock).FormulaU = "1"
            vsoLayer2.CellsC(visLayerColor).FormulaU = "19"
            vsoLayer2.CellsC(visLayerSnap).FormulaU = "0"
            vsoLayer2.CellsC(visLayerGlue).FormulaU = "1"
            ActiveWindow.DeselectAll
        Else
            Exit Sub
        End If
    Else
        'Форма разблокировки заблокированных шейпов
        Load frmMenuUnLockSALayer
        frmMenuUnLockSALayer.Show
    End If
End Sub


Public Sub SetLayer(vsoShape As Visio.Shape, Optional vsoLayer As Visio.Layer)
'------------------------------------------------------------------------------------------------------------
' Macros        : SetLayer - Устанавливает слой для фигуры(группы). Если слой не указан - очищает слой в фигуре(группе)
'------------------------------------------------------------------------------------------------------------
    Dim shpShape As Visio.Shape
    If vsoLayer Is Nothing Then
        vsoShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """"""
        For Each shpShape In vsoShape.Shapes
            If shpShape.Shapes.Count <> 0 Then SetLayer shpShape, vsoLayer
            shpShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """"""
        Next
    Else
        vsoShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & vsoLayer.Index & """"
        For Each shpShape In vsoShape.Shapes
            If shpShape.Shapes.Count <> 0 Then SetLayer shpShape, vsoLayer
            shpShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & vsoLayer.Index & """"
        Next
    End If
End Sub


Sub ShowSettingsProject()
    Load frmMenuSettingsProject
    frmMenuSettingsProject.Show
End Sub

Sub SaveProjectFileAs()
'------------------------------------------------------------------------------------------------------------
' Macros        : SaveProjectFileAs - Сохраняет копию файла с датой
'------------------------------------------------------------------------------------------------------------
    Dim sTime As String
    Dim sPath As String
    Dim sName As String
    Dim oWindow As Window
    Dim oDocument As Visio.Document
    Dim colWindows As Collection
    
    sPath = ActiveDocument.path
    sName = Replace(ActiveDocument.name, ".vsd", "")
    sTime = Format(Now(), "_yyyy.mm.dd_hh.mm.ss")
    If MsgBox("Сохранить копию проекта?" + vbNewLine + vbNewLine + sName, vbQuestion + vbOKCancel, "САПР-АСУ: SaveAs") = vbOK Then
        'Сохраняем наборы элементов
        For Each oDocument In Application.Documents
            If oDocument.Type = visTypeStencil Then
                On Error Resume Next
                oDocument.Save
            End If
        Next
        'Закрываем другие окна + ShapeSheet
        Set colWindows = New Collection
        For Each oWindow In Visio.Application.Windows
           If Not (oWindow.Type = visDrawing And oWindow.SubType = visPageWin) Then ' If oWindow.Type = visSheet Then
                colWindows.Add oWindow
            End If
        Next
        For Each oWindow In colWindows
            oWindow.Close
        Next
        Application.ActiveDocument.SaveAsEx sPath + sName + sTime + ".vsd", visSaveAsWS + visSaveAsListInMRU
        Application.ActiveDocument.SaveAsEx sPath + sName + ".vsd", visSaveAsWS + visSaveAsListInMRU
        MsgBox "Файл сохранен!" + vbNewLine + vbNewLine + sName + sTime + ".vsd", vbInformation + vbOKOnly, "САПР-АСУ: Info"
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
        If vsoPage.name Like PageName & ".*" Then
            colPages.Add vsoPage
        End If
    Next
    
'    If colPages.Count = 0 Then
'        'Создаем первую страницу
'        Set vsoPage = ActiveDocument.Pages.Add
'        vsoPage.name = PageName
'        Set shpRamka = vsoPage.Drop(Ramka, 0, 0)
''        ActiveDocument.Masters.Item("Рамка").Delete
'        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(0,Prop.CHAPTER.Format)"
'
'    Else

        'Ищем номер последней страницы
        MaxNumber = MaxMinPageNumber(colPages)

        If MaxNumber = 0 Then 'Создаем вторую страницу
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.name = PageName & ".2"
            
        Else 'Создаем последующие страницы
            'Находим максимальный номер страницы в NameU и Name
            MaxNpage = MaxMinPageNumber(colPages, , , True)
            'Создаем страницу раздела с максимальным номером
            Set vsoPage = ActiveDocument.Pages.Add
            vsoPage.name = PageName & "." & CStr(MaxNpage + 1)
            'Переименовываем вставленный лист в нумерацию Name после текущего
            vsoPage.name = PageName & "." & CStr(MaxNumber + 1)
        End If
        
        Set shpRamka = vsoPage.Drop(Ramka, 0, 0)
'        ActiveDocument.Masters.Item("Рамка").Delete
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"

'    End If
    shpRamka.Cells("Prop.CNUM") = 0
    shpRamka.Cells("Prop.TNUM") = 0
    vsoPage.PageSheet.Cells("PageWidth").Formula = "420 MM"
    vsoPage.PageSheet.Cells("PageHeight").Formula = "297 MM"
    vsoPage.PageSheet.Cells("Paperkind").Formula = 8
    vsoPage.PageSheet.Cells("PrintPageOrientation").Formula = 2
    
    SetRamkaProp shpRamka
    
    SetPageAction vsoPage
    
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
    Dim NazvanieShkafa As String
    Dim NazvanieFSA As String
    
    Set colPagesAfter = New Collection
    Set colPagesAll = New Collection
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    Set Setka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("SETKA KOORD")
    Set vsoPageSource = ActivePage
    Index = vsoPageSource.Index
    PageName = GetPageName(vsoPageSource.name)
    PageNumber = GetPageNumber(vsoPageSource.name)
    Set shpRamkaSource = vsoPageSource.Shapes("Рамка")

    'Ищем страницы раздела больше текущей
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            colPagesAll.Add vsoPage
            If GetPageNumber(vsoPage.name) > PageNumber Then
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
        vsoPage.name = PageName & "." & CStr(GetPageNumber(vsoPage.name) + 1) & IIf(GetPageDesc(vsoPage.name) = "", "", "." & GetPageDesc(vsoPage.name))
    Wend
    
    'Находим максимальный номер страницы в NameU и Name
    MaxNpage = MaxMinPageNumber(colPagesAll, , , True)
    'Создаем страницу раздела с максимальным номером
    Set vsoPageNew = ActiveDocument.Pages.Add
    vsoPageNew.name = PageName & "." & CStr(MaxNpage + 1)
    'Переименовываем вставленный лист в нумерацию Name после текущего
    vsoPageNew.name = PageName & "." & CStr(PageNumber + 1)
    'Положение новой страницы сразу за текущей
    vsoPageNew.Index = Index + 1
    Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
'    ActiveDocument.Masters.Item("Рамка").Delete
    shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"
    shpRamka.Cells("Prop.Type").Formula = shpRamkaSource.Cells("Prop.Type").Formula
    shpRamka.Cells("Prop.CNUM").Formula = shpRamkaSource.Cells("Prop.CNUM").Formula
    shpRamka.Cells("Prop.TNUM").Formula = shpRamkaSource.Cells("Prop.TNUM").Formula
    vsoPageNew.PageSheet.Cells("PageWidth").Formula = vsoPageSource.PageSheet.Cells("PageWidth").Formula
    vsoPageNew.PageSheet.Cells("PageHeight").Formula = vsoPageSource.PageSheet.Cells("PageHeight").Formula
    vsoPageNew.PageSheet.Cells("Paperkind").Formula = vsoPageSource.PageSheet.Cells("Paperkind").Formula
    vsoPageNew.PageSheet.Cells("PrintPageOrientation").Formula = vsoPageSource.PageSheet.Cells("PrintPageOrientation").Formula
    If vsoPageSource.name Like cListNameCxema & "*" Then
        If vsoPageSource.PageSheet.CellExists("Prop.SA_NazvanieShkafa", 0) Then
            SetNazvanieShkafa vsoPageNew.PageSheet
            vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShkafa.Format").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieShkafa.Format").Formula
            vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShkafa").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieShkafa").Formula
            vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieMesta").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieMesta").Formula
            vsoPageNew.Drop Setka, 0, 0
            UpdateNazvanieShkafa
        End If
    End If
    If vsoPageSource.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        SetNazvanieFSA vsoPageNew.PageSheet
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula
    End If
    If PageName = cListNameSpec Then ' "С" 'Спецификация оборудования, изделий и материалов
        shpRamka.Shapes("FORMA3").Shapes("Shifr").Cells("fields.value").FormulaU = "=TheDoc!User.SA_FR_Shifr & "".CO"""
        shpRamka.Cells("User.NomerLista").FormulaU = "=PAGENUMBER()+Sheet.1!Prop.CNUM + TheDoc!User.SA_FR_NListSpecifikac - PAGECOUNT()"
        shpRamka.Cells("User.ChisloListov").FormulaU = "=TheDoc!User.SA_FR_NListSpecifikac"
        ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").FormulaU = ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").Result(0) + 1
    End If
    
    SetRamkaProp shpRamka
    
    SetPageAction vsoPageNew
    
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
    
    If MsgBox("Удалить лист: " & ActivePage.name, vbYesNo + vbCritical, "САПР-АСУ: Удаление листа") = vbYes Then
    
        Set colPagesAfter = New Collection
        NameActivePage = ActivePage.name
        PageName = GetPageName(NameActivePage)
        PageNumber = GetPageNumber(NameActivePage)

        ActiveWindow.DeselectAll
        On Error GoTo err
        ActiveWindow.SelectAll
        ActiveWindow.Selection.Delete
        
        DoEvents
err:
        ActivePage.Delete 0

        If PageName = cListNameSpec Then ' "С" 'Спецификация оборудования, изделий и материалов
            ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").FormulaU = IIf(ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").Result(0) > 0, ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").Result(0) - 1, 0)
        End If

        If PageNumber = 1 Then Exit Sub
        
        'Ищем страницы раздела больше текущей
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.name Like PageName & "*" Then
                If GetPageNumber(vsoPage.name) > PageNumber Then
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
            vsoPage.name = PageName & "." & CStr(GetPageNumber(vsoPage.name) - 1) & IIf(GetPageDesc(vsoPage.name) = "", "", "." & GetPageDesc(vsoPage.name))
        Wend
        If NameActivePage Like cListNameCxema & "*" Then UpdateNazvanieShkafa
        If NameActivePage Like cListNameFSA & "*" Then UpdateNazvanieFSA
    End If
End Sub

Sub CopySAPage()
'------------------------------------------------------------------------------------------------------------
' Sub           : CopySAPage - Копирует страницу САПР-АСУ за текущей, копируя ее свойства и содержимое
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
    Dim NazvanieShkafa As String
    Dim NazvanieFSA As String
    
    Set colPagesAfter = New Collection
    Set colPagesAll = New Collection
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    Set Setka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("SETKA KOORD")
    Set vsoPageSource = ActivePage
    Index = vsoPageSource.Index
    PageName = GetPageName(vsoPageSource.name)
    PageNumber = GetPageNumber(vsoPageSource.name)
    Set shpRamkaSource = vsoPageSource.Shapes("Рамка")

    'Ищем страницы раздела больше текущей
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            colPagesAll.Add vsoPage
            If GetPageNumber(vsoPage.name) > PageNumber Then
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
        vsoPage.name = PageName & "." & CStr(GetPageNumber(vsoPage.name) + 1) & IIf(GetPageDesc(vsoPage.name) = "", "", "." & GetPageDesc(vsoPage.name))
    Wend
    
    'Находим максимальный номер страницы в NameU и Name
    MaxNpage = MaxMinPageNumber(colPagesAll, , , True)
    'Создаем страницу раздела с максимальным номером
    Set vsoPageNew = ActiveDocument.Pages.Add
    vsoPageNew.name = PageName & "." & CStr(MaxNpage + 1)
    'Переименовываем вставленный лист в нумерацию Name после текущего
    vsoPageNew.name = PageName & "." & CStr(PageNumber + 1)
    'Положение новой страницы сразу за текущей
    vsoPageNew.Index = Index + 1
'    Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
'    ActiveDocument.Masters.Item("Рамка").Delete
'    shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"
'    shpRamka.Cells("Prop.Type").Formula = shpRamkaSource.Cells("Prop.Type").Formula
'    shpRamka.Cells("Prop.CNUM").Formula = shpRamkaSource.Cells("Prop.CNUM").Formula
'    shpRamka.Cells("Prop.TNUM").Formula = shpRamkaSource.Cells("Prop.TNUM").Formula
    vsoPageNew.PageSheet.Cells("PageWidth").Formula = vsoPageSource.PageSheet.Cells("PageWidth").Formula
    vsoPageNew.PageSheet.Cells("PageHeight").Formula = vsoPageSource.PageSheet.Cells("PageHeight").Formula
    vsoPageNew.PageSheet.Cells("Paperkind").Formula = vsoPageSource.PageSheet.Cells("Paperkind").Formula
    vsoPageNew.PageSheet.Cells("PrintPageOrientation").Formula = vsoPageSource.PageSheet.Cells("PrintPageOrientation").Formula
    If vsoPageSource.PageSheet.CellExists("Prop.SA_NazvanieShkafa", 0) Then
        SetNazvanieShkafa vsoPageNew.PageSheet
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShkafa.Format").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieShkafa.Format").Formula
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShkafa").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieShkafa").Formula
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieMesta").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieMesta").Formula
        UpdateNazvanieShkafa
'        vsoPageNew.Drop Setka, 0, 0
    End If
    If vsoPageSource.PageSheet.CellExists("Prop.SA_NazvanieFSA", 0) Then
        SetNazvanieFSA vsoPageNew.PageSheet
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula
        vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula = vsoPageSource.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula
    End If
    
    Application.EventsEnabled = False
    SetPageAction vsoPageNew
    Application.ActiveWindow.Page = vsoPageSource
    ActiveWindow.DeselectAll
    ActiveWindow.Selection.Copy
    Application.ActiveWindow.Page = vsoPageNew
    ActiveWindow.Page.Paste
    LockTitleBlock
    ActiveWindow.Selection.Delete

    Application.ActiveWindow.Page = vsoPageSource
    ActiveWindow.DeselectAll
    ActiveWindow.SelectAll
    ActiveWindow.Selection.Copy visCopyPasteNoTranslate
    Application.ActiveWindow.Page = vsoPageNew
    ActivePage.Paste visCopyPasteNoTranslate
    ActiveWindow.DeselectAll
    Application.EventsEnabled = True
    ResetLocalShkafMesto ActivePage
End Sub

Sub SetPageAction(vsoPageNew As Visio.Page)
    Dim PageName As String
    
    PageName = GetPageName(vsoPageNew.name)
    Select Case PageName
        Case cListNameOD ' "ОД" 'Общие указания
        Case cListNameFSA ' "ФСА" 'Схема функциональная автоматизации
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить оборудование со схемы"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageFSAAddSensorsFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
            End With
        Case cListNamePlan ' "План" 'План расположения оборудования и приборов КИП
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить оборудование из ФСА"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PagePLANAddElementsFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """10"""
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Проложить кабели для всего оборудования"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""AddRouteCablesOnPlan"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "2633" '2645
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """20"""
            End With
        Case cListNameCxema ' "Схема" 'Схема электрическая принципиальная
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Обновить """"Шкафы/Места"""" + Перенумерация"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "CALLTHIS(""MISC.ResetLocalShkafMesto"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "688"
            End With
        Case cListNameVID ' "ВИД" 'Чертеж внешнего вида шкафа
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить элементы со схемы"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageVIDAddElementsFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
            End With
        Case cListNameSVP ' "СВП" 'Схема соединения внешних проводок
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить провода со схемы"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageSVPAddKabeliFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Удалить все листы СВП"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""svpDEL"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1088" '2645
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """30"""
            End With
        Case cListNameKJ  ' "КЖ" 'Кабельный журнал
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Создать кабельный журнал в Visio из Excel"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""KJ_Excel_2_Visio"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "7076" '6224
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """20"""
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Удалить все листы кабельного журнала"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""kjDEL"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1088" '2645
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """30"""
            End With
        Case cListNameSpec ' "С" 'Спецификация оборудования, изделий и материалов
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Создать спецификацию в Visio из Excel"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""SP_Excel_2_Visio"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "7076" '6224
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """20"""
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Удалить все листы спецификации"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""spDEL"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1088" '2645
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """30"""
            End With
        Case Else
    End Select
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
        Npage = GetPageNumber(colPages.Item(i).name)
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
        Npage = GetPageNumber(IIf(NameU, vsoPage.NameU, vsoPage.name))
        If Npage < MinNumber Then MinNumber = Npage
        If Npage > MaxNumber Then MaxNumber = Npage
    Next
    Return
End Function

Sub SetNazvanieShkafa(vsoObject As Object) 'SetValueToSelSections
    Dim arrRowValue()
    Dim arrRowName()
    Dim SectionNumber As Long
    SectionNumber = visSectionProp 'Prop 243
    arrRowName = Array("SA_NazvanieShkafa", "SA_NazvanieMesta")
    arrRowValue = Array("""Название Шкафа""|""Нумерация элементов идет в пределах одного шкафа""|4|""""|INDEX(0,Prop.SA_NazvanieShkafa.Format)|""""|FALSE|FALSE|1049|0", _
                        """Название Места""|""Название места расположения или название установки""|0|""""|""""|""""|FALSE|FALSE|1049|0")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber
End Sub

Sub SetNazvanieFSA(vsoObject As Object) 'SetValueToSelSections
    Dim arrRowValue()
    Dim arrRowName()
    Dim SectionNumber As Long
    SectionNumber = visSectionProp 'Prop 243
    arrRowName = Array("SA_NazvanieFSA")
    arrRowValue = Array("""Название ФСА""|""Нумерация элементов идет в пределах одной ФСА""|4|""""|INDEX(0,Prop.SA_NazvanieFSA.Format)|""""|FALSE|FALSE|1049|0")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber
End Sub

Sub UpdateNazvanieShkafa()
    Dim colNameCxema As Collection
    Dim PropPageSheet As String
    Dim i As Integer
    
    Set colNameCxema = GetColNazvanieShkafa
    For i = 1 To colNameCxema.Count
        PropPageSheet = PropPageSheet & colNameCxema.Item(i) & IIf(i = colNameCxema.Count, "", ";")
    Next
    NazvanieShkafaSetToAll PropPageSheet
End Sub

Function GetColNazvanieShkafa() As Collection
    Dim vsoPage As Visio.Page
    Dim vsoShape As Visio.Shape
    Dim colNameCxema As Collection
    Dim PageName As String
    
    Set colNameCxema = New Collection
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            On Error Resume Next
            If vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0) <> "" Then
                colNameCxema.Add vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0), vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
            End If
            err.Clear
            On Error GoTo 0
            For Each vsoShape In vsoPage.Shapes
                If ShapeSATypeIs(vsoShape, typeCxemaShkafMesto) Then
                    On Error Resume Next
                    If vsoShape.Cells("Prop.SA_NazvanieShkafa").ResultStr(0) <> "" Then
                        colNameCxema.Add vsoShape.Cells("Prop.SA_NazvanieShkafa").ResultStr(0), vsoShape.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
                    End If
                    err.Clear
                    On Error GoTo 0
                End If
            Next
        End If
    Next
    Set GetColNazvanieShkafa = colNameCxema
End Function

Sub NazvanieShkafaSetToAll(PropPageSheet As String)
    Dim vsoPage As Visio.Page
    Dim vsoShape As Visio.Shape
    Dim PageName As String
    Dim NazvanieShkafaValue As String
    Dim i As Integer
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            NazvanieShkafaValue = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
            vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa.Format").Formula = """" & PropPageSheet & """"
            vsoPage.PageSheet.Cells("Prop.SA_NazvanieShkafa").Formula = """" & NazvanieShkafaValue & """"
            For Each vsoShape In vsoPage.Shapes
                If ShapeSATypeIs(vsoShape, typeCxemaShkafMesto) Then
                    NazvanieShkafaValue = vsoShape.Cells("Prop.SA_NazvanieShkafa").ResultStr(0)
                    vsoShape.Cells("Prop.SA_NazvanieShkafa.Format").Formula = """" & PropPageSheet & """"
                    vsoShape.Cells("Prop.SA_NazvanieShkafa").Formula = """" & NazvanieShkafaValue & """"
                End If
            Next
        End If
    Next
End Sub

Sub UpdateNazvanieFSA()
    Dim colNameFSA As Collection
    Dim PropPageSheet As String
    Dim i As Integer
    
    Set colNameFSA = GetColNazvanieFSA
    For i = 1 To colNameFSA.Count
        PropPageSheet = PropPageSheet & colNameFSA.Item(i) & IIf(i = colNameFSA.Count, "", ";")
    Next
    NazvanieFSASetToAll PropPageSheet
End Sub

Function GetColNazvanieFSA() As Collection
    Dim vsoPage As Visio.Page
    Dim vsoShape As Visio.Shape
    Dim colNameFSA As Collection
    Dim PageName As String
    
    Set colNameFSA = New Collection
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            On Error Resume Next
            If vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0) <> "" Then
                colNameFSA.Add vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0), vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
            End If
            err.Clear
            On Error GoTo 0
        End If
    Next
    Set GetColNazvanieFSA = colNameFSA
End Function

Sub NazvanieFSASetToAll(PropPageSheet As String)
    Dim vsoPage As Visio.Page
    Dim vsoShape As Visio.Shape
    Dim PageName As String
    Dim NazvanieFSAValue As String
    Dim i As Integer
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.name Like PageName & "*" Then
            NazvanieFSAValue = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").ResultStr(0)
            vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = """" & PropPageSheet & """"
            vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA").Formula = """" & NazvanieFSAValue & """"
        End If
    Next
End Sub

Sub SetTheDocInAllFrame()
'------------------------------------------------------------------------------------------------------------
' Macros        : SetTheDocInAllFrame - Перезаписывает формулы с TheDoc!Var чтобы они обновлялись во всех рамках
'               Gennady Tumanov
'               VisioPort blog: Опасные ссылки на TheDoc в Visio
'               https://visioport.epizy.com/blog/34-thedocref.html
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim shpRamka As Visio.Shape

    For Each vsoPage In ActiveDocument.Pages    'Перебираем все листы в активном документе
        On Error Resume Next
        Set shpRamka = vsoPage.Shapes("Рамка")
        err.Clear
        On Error GoTo 0
        SetRamkaProp shpRamka
    Next
End Sub

Sub SetRamkaProp(shpRamka As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : SetRamkaProp - Перезаписывает формулы с TheDoc!Var чтобы они обновлялись в рамке
'               Gennady Tumanov
'               VisioPort blog: Опасные ссылки на TheDoc в Visio
'               https://visioport.epizy.com/blog/34-thedocref.html
'------------------------------------------------------------------------------------------------------------
    If Not shpRamka Is Nothing Then
        With shpRamka.Shapes("FORMA3")
            .Shapes("Razrabotal").Cells("Prop.date").Formula = .Shapes("Razrabotal").Cells("Prop.date").Formula
            .Shapes("Razrabotal").Cells("Prop.Row_2").Formula = .Shapes("Razrabotal").Cells("Prop.Row_2").Formula
            
            .Shapes("Proveril").Cells("Prop.date").Formula = .Shapes("Proveril").Cells("Prop.date").Formula
            .Shapes("Proveril").Cells("Prop.Row_2").Formula = .Shapes("Proveril").Cells("Prop.Row_2").Formula
            
            .Shapes("gip").Cells("Prop.date").Formula = .Shapes("gip").Cells("Prop.date").Formula
            .Shapes("gip").Cells("Prop.Row_2").Formula = .Shapes("gip").Cells("Prop.Row_2").Formula
            
            .Shapes("NachOtdela").Cells("Prop.date").Formula = .Shapes("NachOtdela").Cells("Prop.date").Formula
            .Shapes("NachOtdela").Cells("Prop.Row_2").Formula = .Shapes("NachOtdela").Cells("Prop.Row_2").Formula
            
            .Shapes("Utverdil").Cells("Prop.date").Formula = .Shapes("Utverdil").Cells("Prop.date").Formula
            .Shapes("Utverdil").Cells("Prop.Row_2").Formula = .Shapes("Utverdil").Cells("Prop.Row_2").Formula
            
            .Shapes("NKontr").Cells("Prop.date").Formula = .Shapes("NKontr").Cells("Prop.date").Formula
            .Shapes("NKontr").Cells("Prop.Row_2").Formula = .Shapes("NKontr").Cells("Prop.Row_2").Formula
        End With
    End If
End Sub

 Sub del_pages(ListName As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : del_pages - Удаляет листы проекта
'------------------------------------------------------------------------------------------------------------
    Dim dp As Page
    Dim colPage As Collection
    Set colPage = New Collection
    'Листы в колекцию
    For Each dp In ActiveDocument.Pages
        If dp.name Like ListName & ".*" Then
            colPage.Add dp
        End If
    Next
    'удаляем все страницы которые нашли выше
    For Each dp In colPage
        dp.Delete 1
    Next
    ActiveWindow.Page = ActiveDocument.Pages.Item(ListName)
    ActiveWindow.SelectAll
    ActiveWindow.Selection.Delete
End Sub

Sub SavePDF()
'------------------------------------------------------------------------------------------------------------
' Macros        : SavePDF - Сохраняет листы проекта в PDF. Все цвета - черные
'------------------------------------------------------------------------------------------------------------
    Dim str As String
    ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 1
    DoEvents
    If MsgBox("Сохранить в PDF?" + vbNewLine + vbNewLine + Replace(ActiveDocument.name, ".vsd", ""), vbQuestion + vbOKCancel, "САПР-АСУ: Save PDF") = vbOK Then
'        ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 1
        DoEvents
        str = Replace(ActiveDocument.name, ".vsd", "") & Format(Now(), "_yyyy.mm.dd_hh.mm.ss") & ".pdf"
        Application.ActiveDocument.ExportAsFixedFormat visFixedFormatPDF, ActiveDocument.path & str, visDocExIntentPrint, visPrintAll, 1, ActiveDocument.Pages.Count, True, True, True, True, False 'Первый true - все цвета чёрные
'        ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 0
        MsgBox "Файл сохранен в папке проекта!" + vbNewLine + vbNewLine + str, vbInformation + vbOKOnly, "САПР-АСУ: Info"
    End If
    ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 0
End Sub

Sub SavePDFColor()
'------------------------------------------------------------------------------------------------------------
' Macros        : SavePDFColor - Сохраняет листы проекта в PDF. Цвета - цветные
'------------------------------------------------------------------------------------------------------------
    Dim str As String
    ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 1
    DoEvents
    DoLockLayers 0
    DoEvents
    If MsgBox("Сохранить в PDF в цвете?" + vbNewLine + vbNewLine + Replace(ActiveDocument.name, ".vsd", ""), vbQuestion + vbOKCancel, "САПР-АСУ: Save PDF") = vbOK Then
'        ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 1
        DoEvents
        str = Replace(ActiveDocument.name, ".vsd", "") & Format(Now(), "_yyyy.mm.dd_hh.mm.ss") & ".pdf"
        Application.ActiveDocument.ExportAsFixedFormat visFixedFormatPDF, ActiveDocument.path & str, visDocExIntentPrint, visPrintAll, 1, ActiveDocument.Pages.Count, False, True, True, True, False 'Первый true - все цвета чёрные
'        ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 0
        MsgBox "Файл сохранен в папке проекта!" + vbNewLine + vbNewLine + str, vbInformation + vbOKOnly, "САПР-АСУ: Info"
    End If
    ActiveDocument.DocumentSheet.Cells("User.SA_NoColor").Formula = 0
    DoEvents
    DoLockLayers 1
    DoEvents
End Sub

Public Sub DoLockLayers(bLock As Boolean)
'------------------------------------------------------------------------------------------------------------
' Macros        : DoLockLayers - Блокировка=1/Разблокировка=0 слоёв
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Dim vsoLayer1 As Visio.Layer

    For Each vsoPage In ActiveDocument.Pages
        For Each vsoLayer1 In vsoPage.Layers
            If vsoLayer1.name = "SA_Рамка" Then
                GoSub LockSub
            End If
            If vsoLayer1.name = "SA_LockedWire" Then
                GoSub LockSub
            End If
            If vsoLayer1.name = "SA_LockedLayer" Then
                GoSub LockSub
            End If
        Next
    Next
    Exit Sub
    
LockSub:
        If bLock Then
            'Блокруем слой
            vsoLayer1.CellsC(visLayerLock).FormulaU = "1"
            vsoLayer1.CellsC(visLayerColor).FormulaU = "19"
            vsoLayer1.CellsC(visLayerSnap).FormulaU = "0"
            vsoLayer1.CellsC(visLayerGlue).FormulaU = "0"
        Else
            'Разблокруем слой
            vsoLayer1.CellsC(visLayerLock).FormulaU = "0"
            vsoLayer1.CellsC(visLayerColor).FormulaU = "255"
            vsoLayer1.CellsC(visLayerSnap).FormulaU = "0"
            vsoLayer1.CellsC(visLayerGlue).FormulaU = "0"
        End If
Return
End Sub

'-----------------------------Переделка таблицы спецификации под универсальную---------------------------------
Sub TuneTable_1()
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    For i = 1 To 30
        Set shpRow = ActivePage.Shapes("СП").Shapes("row" & i)
        shpRow.Shapes(i & "." & 1).Cells("Width").FormulaU = "=Sheet.65!Width"
        shpRow.Shapes(i & "." & 1).Cells("PinX").FormulaU = "=Sheet.65!PinX"
        shpRow.Shapes(i & "." & 2).Cells("Width").FormulaU = "=Sheet.57!Width"
        shpRow.Shapes(i & "." & 2).Cells("PinX").FormulaU = "=Sheet.57!PinX"
        shpRow.Shapes(i & "." & 3).Cells("Width").FormulaU = "=Sheet.64!Width"
        shpRow.Shapes(i & "." & 3).Cells("PinX").FormulaU = "=Sheet.64!PinX"
        shpRow.Shapes(i & "." & 4).Cells("Width").FormulaU = "=Sheet.62!Width"
        shpRow.Shapes(i & "." & 4).Cells("PinX").FormulaU = "=Sheet.62!PinX"
        shpRow.Shapes(i & "." & 5).Cells("Width").FormulaU = "=Sheet.61!Width"
        shpRow.Shapes(i & "." & 5).Cells("PinX").FormulaU = "=Sheet.61!PinX"
        shpRow.Shapes(i & "." & 6).Cells("Width").FormulaU = "=Sheet.60!Width"
        shpRow.Shapes(i & "." & 6).Cells("PinX").FormulaU = "=Sheet.60!PinX"
        shpRow.Shapes(i & "." & 7).Cells("Width").FormulaU = "=Sheet.63!Width"
        shpRow.Shapes(i & "." & 7).Cells("PinX").FormulaU = "=Sheet.63!PinX"
        shpRow.Shapes(i & "." & 8).Cells("Width").FormulaU = "=Sheet.59!Width"
        shpRow.Shapes(i & "." & 8).Cells("PinX").FormulaU = "=Sheet.59!PinX"
        shpRow.Shapes(i & "." & 9).Cells("Width").FormulaU = "=Sheet.58!Width"
        shpRow.Shapes(i & "." & 9).Cells("PinX").FormulaU = "=Sheet.58!PinX"
        shpRow.Shapes(i & "." & 10).Cells("Width").FormulaU = "=Sheet.367!Width"
        shpRow.Shapes(i & "." & 10).Cells("PinX").FormulaU = "=Sheet.367!PinX"
        For j = 1 To 10
            Set shpCel = shpRow.Shapes(i & "." & j)
            shpCel.Cells("PinY").FormulaU = shpRow.NameID & "!Height*0"
            shpCel.Cells("LocPinX").FormulaU = "=Width*0"
            shpCel.Cells("LocPinY").FormulaU = "=Height*0"
        Next
    Next
End Sub

Sub TuneTable_2()
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    For i = 1 To 30
        Set shpRow = ActivePage.Shapes("СП").Shapes("row" & i)
        shpRow.Cells("Height").FormulaForceU = Replace(shpRow.Cells("Height").FormulaU, "))", "," & shpRow.Shapes(i & ".10").NameID & "!User.Row_1))")
    Next
End Sub

Sub TuneTable_3()
    Dim shpCel As Visio.Shape
    For i = 1 To 10
        If i < 10 Then
            Set shpCel = ActivePage.Shapes("СП").Shapes("Head").Shapes("0" & i)
        Else
            Set shpCel = ActivePage.Shapes("СП").Shapes("Head").Shapes("10")
        End If
        With shpCel
            .AddSection visSectionFirstComponent
            .AddRow visSectionFirstComponent, visRowComponent, visTagComponent
            .AddRow visSectionFirstComponent, visRowVertex, visTagLineTo
            .AddRow visSectionFirstComponent, visRowVertex, visTagMoveTo
            .CellsSRC(visSectionFirstComponent, 0, 0).FormulaForceU = "TRUE"
            .CellsSRC(visSectionFirstComponent, 0, 1).FormulaForceU = "FALSE"
            .CellsSRC(visSectionFirstComponent, 0, 2).FormulaForceU = "FALSE"
            .CellsSRC(visSectionFirstComponent, 0, 3).FormulaForceU = "FALSE"
            .CellsSRC(visSectionFirstComponent, 1, 0).FormulaU = "Width*0"
            .CellsSRC(visSectionFirstComponent, 1, 1).FormulaU = "Height*0"
            .CellsSRC(visSectionFirstComponent, 2, 0).FormulaU = "Width*1"
            .CellsSRC(visSectionFirstComponent, 2, 1).FormulaU = "Height*0"
            .AddRow visSectionFirstComponent, 3, visTagLineTo
            .CellsSRC(visSectionFirstComponent, 3, 0).FormulaU = "Width*1"
            .CellsSRC(visSectionFirstComponent, 3, 1).FormulaU = "Height * 1"
            .AddRow visSectionFirstComponent, 4, visTagLineTo
            .CellsSRC(visSectionFirstComponent, 4, 0).FormulaU = "Width*0"
            .CellsSRC(visSectionFirstComponent, 4, 1).FormulaU = "Height*1"
            .AddRow visSectionFirstComponent, 5, visTagLineTo
            .CellsSRC(visSectionFirstComponent, 5, 0).FormulaU = "Width*0"
            .CellsSRC(visSectionFirstComponent, 5, 1).FormulaU = "Geometry1.Y1"
        End With
    Next
End Sub

Sub TuneTable_4()
    Dim shpCel As Visio.Shape
    For i = 1 To 10
        If i < 10 Then
            Set shpCel = ActivePage.Shapes("СП").Shapes("Head").Shapes("0" & i)
        Else
            Set shpCel = ActivePage.Shapes("СП").Shapes("Head").Shapes("10")
        End If
        shpCel.Cells("Width").FormulaU = "=Sheet.47!Width*Sheet.45!Prop.S" & i & "/Sheet.45!Prop.Width"
    Next
End Sub

Sub TuneTable_5() 'поля
    Dim shpRow As Visio.Shape
    For i = 1 To 30
        Set shpRow = ActivePage.Shapes("СП").Shapes("row" & i)
        shpRow.Shapes(i & "." & 2).CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "10 pt"
        shpRow.Shapes(i & "." & 2).CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "10 pt"
        shpRow.Shapes(i & "." & 4).CellsSRC(visSectionObject, visRowText, visTxtBlkLeftMargin).FormulaU = "5 pt"
        shpRow.Shapes(i & "." & 4).CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "1 pt"
    Next
End Sub
    
Sub TuneTable_6() 'очистка таблицы
    Dim shpRow As Visio.Shape
    Dim shpCel As Visio.Shape
    For i = 1 To 30
        Set shpRow = ActivePage.Shapes("СП").Shapes("row" & i)
        For j = 1 To 10
            Set shpCel = shpRow.Shapes(i & "." & j)
            shpCel.text = " "
        Next
    Next
End Sub

'-----------------------------------------------------------------------------------------------

'Преобразует тип данных артикула в избранном к типу текст
Sub ConvertArticulIzbrannoe()
    Dim lLastRow As Long
    Dim UserRange As Excel.Range
    Set oExcelAppIzbrannoe = CreateObject("Excel.Application")
    sSAPath = Visio.ActiveDocument.path
    Set wbExcelIzbrannoe = oExcelAppIzbrannoe.Workbooks.Open(sSAPath & DBNameIzbrannoeExcel)
    lLastRow = wbExcelIzbrannoe.Sheets(ExcelIzbrannoe).Cells(wbExcelIzbrannoe.Sheets(ExcelIzbrannoe).Rows.Count, 1).End(xlUp).Row
    Set UserRange = wbExcelIzbrannoe.Worksheets(ExcelIzbrannoe).Range("A2:A" & lLastRow)
    ExcelConvertToString UserRange
    lLastRow = wbExcelIzbrannoe.Sheets(ExcelNabory).Cells(wbExcelIzbrannoe.Sheets(ExcelNabory).Rows.Count, 1).End(xlUp).Row
    Set UserRange = wbExcelIzbrannoe.Worksheets(ExcelNabory).Range("A2:A" & lLastRow)
    ExcelConvertToString UserRange
    wbExcelIzbrannoe.Close savechanges:=True
    oExcelAppIzbrannoe.Quit
End Sub