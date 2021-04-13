'------------------------------------------------------------------------------------------------------------
' Module        : OD - Общие данные
' Author        : gtfox
' Date          : 2019.09.22
' Description   : odDELL - Удаляет общие данные
                ' odADD_A3 - Добавляет общие данные на листах А3, и если не хватает на последний А3 - добавляет А4
                ' odADD_A4 - Добавляет общие данные на листах только А4.
                ' OD_2_Visio.doc - Общие данные (текстовая часть проекта) - Содержит исходный текст, который будет порезан на листы и вставлен в чертеж Visio при помощи макроса.
                ' В результате его работы создается OD_2_Visio_Split.doc (в дальнейшем не используется + перезаписывается при каждом вызове макроса)
                ' На лист, с которого начинаются общие данные, кидаем фигуру ОД. Настраиваем верхнюю/нижнюю границы рамки текста (чтобы они двигались надо переместить фигуру влево - вправо из контекстного меню). Запускаем макрос odADD_А3 / odADD_А4
                ' Основная проблема текстовых данных в Visio – отсутствие автопереноса текста на новую страницу/новый шейп, а также нет возможности обращаться к отдельным строкам текста.
                ' Я решил возложить эту задачу на Word. Зная размеры шейпа ОД мы задаем поля в Word, лишний текст там переносится и мы копируем содержимое страницы в шейп, потом вставляем разрыв раздела и на следующей странице ставим новые поля для нового шейпа ОД.
' Link          : https://visio.getbb.ru/viewtopic.php?p=14130, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
    
    
Public Sub odADD_A4()
    OD_2_Visio (1)
End Sub

Public Sub odADD_A3()
    OD_2_Visio (0)
End Sub

    
Private Sub OD_2_Visio(A4 As Boolean)
    'нижнее поле в ворде для рамок в визио
    Const ramka5 = 1
    Const ramka15 = 2.5
    Const ramka55 = 6.5
    nA3 = 1
    
    Dim wa As Word.Application
    Dim vsoCharacters1 As Visio.Characters
    Dim oStartPage As Word.Range
    Dim oEndPage As Word.Range
    Dim nStartPageNum As Long
    Dim nPagesCount As Long
    Dim nEndPageNum As Long
    Dim sPath, sFile As String
    Dim objFSO As Object, objFile As Object
    Dim MastOD As Master
    Set MastOD = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item(cListNameOD)

    
    
    
    If Not Application.ActiveWindow.Selection.Count = 0 Then
    
        If InStr(1, Application.ActiveWindow.Selection.Item(1).Name, cListNameOD) > 0 Then
            
            Set vsoCharacters1 = Application.ActiveWindow.Selection.Item(1).Characters
            
            'есть файл
            sPath = Visio.ActiveDocument.path
            sFileName = "OD_2_Visio.doc"
            sFile = sPath & sFileName
            If Dir(sFile, 16) = "" Then
                MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "Ошибка"
                Exit Sub
            End If
            
            'подготавливаем копирование
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set objFile = objFSO.GetFile(sFile)
    
            'удаляем старый
            sFileName = "OD_2_Visio_Split.doc"
            sFile = sPath & sFileName
            If Len(Dir(sFile)) > 0 Then 'есть хотя бы один файл
                'On Error GoTo L1
                Kill sFile
            End If
            
            'копируем файл с новым именем
            objFile.Copy sFile
            
            'переименовываем новый
            'Name sPath & "ОД - копия.doc" As sFile
    
            Set wa = CreateObject("Word.Application")
            wa.Documents.Open (sFile)
            wa.Visible = True
            Set wad = wa.ActiveDocument
            
            'Заменяет разрывы страницы на разрывы раздела
            ReplacePageBreaks wa
    
            wa.Selection.WholeStory 'выделить все
     
            DoEvents
     
            With wa.Selection.Font
                .Name = "ISOCPEUR"
                .Size = 14
                .Bold = False
                .Italic = True
                .Underline = wdUnderlineNone
                .UnderlineColor = wdColorAutomatic
                .Strikethrough = False
                .DoubleStrikeThrough = False
                .Outline = False
                .Emboss = False
                .Shadow = False
                .Hidden = False
                .SmallCaps = False
                .AllCaps = False
                .Color = wdColorAutomatic
                .Engrave = False
                .Superscript = False
                .Subscript = False
                .Spacing = 0
                .Scaling = 100
                .Position = 0
                .Kerning = 0
                .Animation = wdAnimationNone
            End With
            
            DoEvents
            
            With wa.Selection.ParagraphFormat
                .LeftIndent = wa.CentimetersToPoints(0)
                .RightIndent = wa.CentimetersToPoints(0)
                .SpaceBefore = 5
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = wa.LinesToPoints(1) 'междустрочный интервал
                .Alignment = wdAlignParagraphJustify
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = wa.CentimetersToPoints(1)
                .OutlineLevel = wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = False
                .TextboxTightWrap = wdTightNone
            End With
            
            DoEvents
            
            With wa.Selection.PageSetup
                .LineNumbering.Active = False
                .Orientation = wdOrientLandscape
                .TopMargin = wa.CentimetersToPoints(1)
                .LeftMargin = wa.CentimetersToPoints(2.5)
                .RightMargin = wa.CentimetersToPoints(1)
                '.BottomMargin = wa.CentimetersToPoints(1) 'рамка 5
                .BottomMargin = wa.CentimetersToPoints(2.5) 'рамка 15
                '.BottomMargin = wa.CentimetersToPoints(6.5) 'рамка 55
                .Gutter = wa.CentimetersToPoints(0)
                .HeaderDistance = wa.CentimetersToPoints(0)
                .FooterDistance = wa.CentimetersToPoints(0)
                .PageWidth = wa.CentimetersToPoints(21)
                .PageHeight = wa.CentimetersToPoints(29.7)
                .FirstPageTray = wdPrinterDefaultBin
                .OtherPagesTray = wdPrinterDefaultBin
                .SectionStart = wdSectionNewPage
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .VerticalAlignment = wdAlignVerticalTop
                .SuppressEndnotes = False
                .MirrorMargins = False
                .TwoPagesOnOne = False
                .BookFoldPrinting = False
                .BookFoldRevPrinting = False
                .BookFoldPrintingSheets = 1
                .GutterPos = wdGutterPosLeft
            End With
            
            'табуляция по центру в визио
    '        Application.ActiveWindow.Selection.Item(1).CellsSRC(visSectionTab, 0, visTabStopCount).FormulaU = "1"
    '        Application.ActiveWindow.Selection.Item(1).CellsSRC(visSectionTab, 0, visTabPos).FormulaU = "Guard(92.5 mm)"
    '        Application.ActiveWindow.Selection.Item(1).CellsSRC(visSectionTab, 0, visTabAlign).FormulaU = "Guard(1)"
    '        Application.ActiveWindow.Selection.Item(1).CellsSRC(visSectionTab, 0, 3).FormulaU = "0"
            
            
            'табуляция по центру в ворде
            wa.Selection.ParagraphFormat.TabStops.Add Position:=wa.CentimetersToPoints(9.25), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces 'табуляция по центру
            
            
            hh = Application.ActiveWindow.Selection.Item(1).Cells("Height") ' высота первого куска текста в визио
            niznee_pole = 297 - 10 - hh * 25.4  'нижнее поле на странице в ворде в мм (для вставки разрыва)
            
    
            'верх сраницы 1
            wa.Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Name:="1"
            wa.Selection.PageSetup.BottomMargin = wa.CentimetersToPoints(niznee_pole / 10) 'ставим нижнее поле в см
            
            nStartPageNum = 1
            Set oStartPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum)
            nEndPageNum = 1
            'Конец последней страницы для выделения
            Set oEndPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum + nEndPageNum)  '.GoToNext(wdGoToPage)
            'Выделяем указанный диапазон документа
            wad.Range(oStartPage.Start, oEndPage.End).Select ' wad.Range(oStartPage.Start, IIf(nStartPageNum + nEndPageNum = nPagesCount + 1, wad.Range.End, oEndPage.End)).Select
            'копируем в буфер в ворде
            wa.Selection.Copy
            'вставляем из буфера в визио
            ActiveWindow.SelectedText.Paste
            'скрываем рамку текста
            ActivePage.Shapes.Item("ОД").Cells("Geometry1.NoLine").Formula = 1
            
            'переходим в начало 2-го листа ворда
            wa.Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Name:="2"
            wa.Selection.MoveEnd wdCharacter, -1 'шаг назад - конец предыдущей страницы
            wa.Selection.InsertBreak Type:=wdSectionBreakNextPage 'вставка разрыв раздела
            
            'ставим поле для рамки 15 чтобы перед первым проходом цикла for иметь "более/менее" реальное число листов
            niznee_pole = ramka15
            wa.Selection.PageSetup.BottomMargin = wa.CentimetersToPoints(niznee_pole) 'ставим нижнее поле в см
            
            nPagesCount = wad.Range.ComputeStatistics(wdStatisticPages) 'число листов ворда
            nPagesOst = nPagesCount - 1
            pNumberVisio = 1
            
            For CurPage = 2 To nPagesCount
                'переходим на верх текущего листа
                wa.Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Name:=CurPage
    
                If nPagesOst = 1 Or A4 Then 'последний лист или выбрано "все листы А4"
                
                    'нижнее поле в ворде для этого листа visio
                    niznee_pole = ramka15
                    wa.Selection.PageSetup.BottomMargin = wa.CentimetersToPoints(niznee_pole) 'ставим нижнее поле в см
                    'вставляем лист А4
                    Set aPage = AddNamedPageOD(cListNameOD & "." & pNumberVisio + 1)
                    If aPage Is Nothing Then
                        MsgBox "Лист " & cListNameOD & "." & CStr(pNumberVisio + 1) & " уже существует" & vbNewLine & "Сначала удалите существующие листы ОД", vbCritical, "Ошибка"
                        wad.Close SaveChanges:=False
                        wa.Quit
                        Set wa = Nothing
                        Exit Sub
                    End If
                    aPage.Index = 2 + pNumberVisio 'суем страницу за текущим листом ОД
                    pNumberVisio = pNumberVisio + 1
                    ActivePage.PageSheet.Cells("PageWidth").Formula = "210 MM"
                    ActivePage.PageSheet.Cells("PageHeight").Formula = "297 MM"
                    ActivePage.PageSheet.Cells("Paperkind").Formula = 9
                    ActivePage.PageSheet.Cells("PrintPageOrientation").Formula = 1
                    ActivePage.Drop MastOD, 6.889764, 8.661417
                    'скрываем рамку текста
                    ActiveWindow.Selection.Item(1).Cells("Geometry1.NoLine").Formula = 1
                    ActiveWindow.Selection.Item(1).Cells("Height").FormulaForceU = "(PinY-TheDoc!User.SA_FR_OffsetFrame-15 mm)/ThePage!PageScale*ThePage!DrawingScale"
                    'выделяем фигуру для последующей вставки текста
                    'shpOD.Paste '.Select 'либо если есть метод paste сразу
                    'выбрали диапазон текущего листа
                    nStartPageNum = CurPage
                    Set oStartPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum)
                    nEndPageNum = CurPage
                    'Конец последней страницы для выделения
                    Set oEndPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum + 1)  '.GoToNext(wdGoToPage)
                    'Выделяем указанный диапазон документа
                    wad.Range(oStartPage.Start, IIf(nStartPageNum = nPagesCount, wad.Range.End, oEndPage.End)).Select 'wad.Range(oStartPage.Start, oEndPage.End).Select '
                    'копируем в буфер в ворде
                    wa.Selection.Copy
                    
                    If Not nStartPageNum = nPagesCount Then
                        If oEndPage.Characters(1).Previous.Text <> ChrW(12) Then
                            oEndPage.InsertBreak Type:=wdSectionBreakNextPage 'вставка разрыв раздела
                        End If
                    End If

    
                    DoEvents
                    'shpOD.Paste
                    'вставляем из буфера в визио
                    ActiveWindow.SelectedText.Paste
                    
                    'оставшееся число страниц ворда
                    nPagesOst = nPagesCount - CurPage
    
                ElseIf nPagesOst >= 2 Then   'листов больше 2-х добавляем А3
                    
                    If nA3 = 1 Then ' левая половина А3
                    
                        'нижнее поле в ворде для этого листа visio
                        niznee_pole = ramka5
                        wa.Selection.PageSetup.BottomMargin = wa.CentimetersToPoints(niznee_pole) 'ставим нижнее поле в см
                        'вставляем лист А3
                        Set aPage = AddNamedPageOD(cListNameOD & "." & pNumberVisio + 1)
                        If aPage Is Nothing Then
                            MsgBox "Лист " & cListNameOD & "." & CStr(pNumberVisio + 1) & " уже существует" & vbNewLine & "Сначала удалите существующие листы ОД", vbCritical, "Ошибка"
                            wad.Close SaveChanges:=False
                            wa.Quit
                            Set wa = Nothing
                            Exit Sub
                        End If
                        aPage.Index = 2 + pNumberVisio 'суем страницу за текущим листом ОД
                        ActivePage.PageSheet.Cells("PageWidth").Formula = "420 MM"
                        ActivePage.PageSheet.Cells("PageHeight").Formula = "297 MM"
                        ActivePage.PageSheet.Cells("Paperkind").Formula = 8
                        ActivePage.PageSheet.Cells("PrintPageOrientation").Formula = 2
                        ActivePage.Drop MastOD, 6.889764, 8.661417
                        With ActiveWindow.Selection.Item(1) 'сдвигаем ОД влево
                            .Cells("Geometry1.NoLine").Formula = 1 'скрываем рамку текста
                            .Cells("PinX").FormulaForceU = "GUARD((25 mm-TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale)"
                            .Cells("PinY").FormulaForceU = "(ThePage!PageHeight-TheDoc!User.SA_FR_OffsetFrame)/ThePage!PageScale*ThePage!DrawingScale"
                            .Cells("Height").FormulaForceU = "ThePage!PageHeight-TheDoc!User.SA_FR_OffsetFrame*2"
                            .Cells("Actions.right.Invisible").Formula = 0
                            .Cells("Actions.left.Invisible").Formula = 1
                        End With

                        'выбрали диапазон текущего листа
                        nStartPageNum = CurPage
                        Set oStartPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum)
                        nEndPageNum = CurPage
                        'Конец последней страницы для выделения
                        Set oEndPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum + 1)  '.GoToNext(wdGoToPage)
                        'Выделяем указанный диапазон документа
                        wad.Range(oStartPage.Start, IIf(nStartPageNum = nPagesCount, wad.Range.End, oEndPage.End)).Select 'wad.Range(oStartPage.Start, oEndPage.End).Select '
                        'копируем в буфер в ворде
                        wa.Selection.Copy
                        
                        If Not nStartPageNum = nPagesCount Then
                            If oEndPage.Characters(1).Previous.Text <> ChrW(12) Then
                                oEndPage.InsertBreak Type:=wdSectionBreakNextPage 'вставка разрыв раздела
                            End If
                        End If

                        DoEvents
                        'вставляем из буфера в визио
                        ActiveWindow.SelectedText.Paste
                        nA3 = 2
                        
                    ElseIf nA3 = 2 Then ' правая половина А3
                        
                        'нижнее поле в ворде для этого листа visio
                        niznee_pole = ramka15
                        wa.Selection.PageSetup.BottomMargin = wa.CentimetersToPoints(niznee_pole) 'ставим нижнее поле в см
                        pNumberVisio = pNumberVisio + 1
                        ActivePage.Drop MastOD, 6.889764, 8.661417
                        'скрываем рамку текста
                        ActiveWindow.Selection.Item(1).Cells("Geometry1.NoLine").Formula = 1
                        ActiveWindow.Selection.Item(1).Cells("Height").FormulaForceU = "(PinY-TheDoc!User.SA_FR_OffsetFrame-15 mm)/ThePage!PageScale*ThePage!DrawingScale"
                        'выбрали диапазон текущего листа
                        nStartPageNum = CurPage
                        Set oStartPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum)
                        nEndPageNum = CurPage
                        'Конец последней страницы для выделения
                        Set oEndPage = wad.Range.GoTo(wdGoToPage, wdGoToAbsolute, nStartPageNum + 1)  '.GoToNext(wdGoToPage)
                        'Выделяем указанный диапазон документа
                        wad.Range(oStartPage.Start, IIf(nStartPageNum = nPagesCount, wad.Range.End, oEndPage.End)).Select 'wad.Range(oStartPage.Start, oEndPage.End).Select '
                        'копируем в буфер в ворде
                        wa.Selection.Copy
                        
                        If Not nStartPageNum = nPagesCount Then
                            If oEndPage.Characters(1).Previous.Text <> ChrW(12) Then
                                oEndPage.InsertBreak Type:=wdSectionBreakNextPage 'вставка разрыв раздела
                            End If
                        End If

                        DoEvents
                        'вставляем из буфера в визио
                        ActiveWindow.SelectedText.Paste
                        nA3 = 1
                        'оставшееся число страниц ворда
                        nPagesOst = nPagesCount - CurPage
                        
                        
                    End If
                    
                ElseIf nPagesOst <= 0 Then 'листов больше нет
                    
                End If
                
                nPagesCount = wad.Range.ComputeStatistics(wdStatisticPages) 'число листов ворда

                
            Next CurPage
            
            wad.Close SaveChanges:=True
            wa.Quit
            Set wa = Nothing
            
            Application.ActiveWindow.Page = Application.ActiveDocument.Pages.Item(cListNameOD)

            MsgBox "Текстовая часть ОД добавлена", vbInformation
            Exit Sub
                            
        End If

    End If
    
    MsgBox "Не выделен блок ОД", vbCritical, "Ошибка"
    
    Exit Sub
    
'        wa.Selection.Start = wa.Selection.Start - 1
'        wa.Selection.End = wa.Selection.Start
'        wa.Selection.HomeKey Unit:=wdStory 'верх докуменита
'        wa.Selection.GoToNext (wdGoToPage) 'начало следующей страницы
'        wa.Selection.MoveEnd wdCharacter, -1 'шаг назад - конец предыдущей страницы
'        wa.Selection.InsertBreak Type:=wdSectionBreakNextPage 'вставка разрыв раздела
'        nPagesCount = wad.Range.ComputeStatistics(wdStatisticPages) 'число листов ворда
'With wa.ActiveDocument
'Set Search = .Range(Start:=0, End:=100) 'это чтобы не ковырять весь документ
'Search.Select
'wa.Selection.Find.Execute FindText:="известный текст", Forward:=True
L1:
        MsgBox "Файл " & sFile & " занят и не может быть удален", vbCritical, "Ошибка"
End Sub


Function AddNamedPageOD(pName As String) As Visio.Page
    Dim aPage As Visio.Page
    Dim Ramka As Visio.Master
    Set aPage = ActiveDocument.Pages.Add
    On Error GoTo err
    aPage.Name = pName
    
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    Set sh = ActivePage.Drop(Ramka, 0, 0)
    ActiveDocument.Masters.Item("Рамка").Delete
    ActivePage.Shapes(1).Cells("user.n.value") = 6
    ActivePage.Shapes(1).Cells("Prop.cnum.value") = 0
    ActivePage.Shapes(1).Cells("Prop.tnum.value") = 0
    
    Set AddNamedPageOD = aPage
    Exit Function
err:
    aPage.Delete 1
    Set AddNamedPageOD = Nothing
End Function

Public Sub odDELL()
    Dim dp As Page
    Dim colPage As Collection
    Set colPage = New Collection
    'проходим все страницы и добавляем в коллекцию тока нужные (если удалять сразу тут же, то 3-я страница становится 2-й, а 2-ю for each уже пролистал :) сучара )
    For Each dp In ActiveDocument.Pages
        If InStr(1, dp.Name, cListNameOD & ".") > 0 Then
            colPage.Add dp
        End If
    Next
    'удаляем все страницы которые нашли выше
    For Each dp In colPage
        dp.Delete (1)
    Next
    Set colPage = Nothing
    Application.ActiveWindow.Page = Application.ActiveDocument.Pages.Item(cListNameOD)
    MsgBox "Листы ОД удалены", vbInformation
End Sub

Sub ReplacePageBreaks(wa As Word.Application)
    'Заменяет разрывы страницы на разрывы раздела
    'https://forumvba.ru/index.php?topic=335.0
   
    Dim rng As Word.Range, fnd As Word.Find
   
    '1. Отключение монитора.
'    Application.ScreenUpdating = False
   
    '2. Вставка после разрывов страниц разрывов разделов. Сразу нельзя заменить
        ' разрывы страниц на разрывы разделов, т.к. нет спецсимвола для разрыва
        ' раздела с текущей страницы - у всех разрывов разделов один ansi-символ 12.
       
    '1) Создание объектов для поиска.
    Set rng = wa.ActiveDocument.Range(0, 0)
    Set fnd = rng.Find
   
    '2) Настройка поиска.
    fnd.Text = "^m"
    fnd.MatchWildcards = False
    fnd.Wrap = wdFindStop
   
    '3) Вставка разрывов разделов.
    Do While fnd.Execute = True
        ' Вставка перед разрывом страницы знака абзаца, если его нет, т.к.
            ' это кажется правильнее, чем после текста сразу будет разрыв.
        If rng.Characters(1).Previous.Text <> Chr(13) Then
            rng.InsertBefore Text:=Chr(13)
            ' Знак абзаца будет добавлен в "rng", поэтому смещаем левый край вправо,
                ' чтобы разрыв раздела встал после знака абзаца.
            rng.MoveStart Unit:=wdCharacter, Count:=1
        End If
        ' Вставка перед разрывом страницы разрыва раздела. Разрыв вставляется
            ' именно перед разрывом страницы, а не после, как могло бы показаться.
        rng.InsertBreak Type:=wdSectionBreakNextPage
        ' После вставки разрыва раздела "rng" сделает коллапс в начало найденного разрыва страницы,
            ' поэтому нужно сместится вправо на один символ, чтобы выйти за пределы
            ' найденного разрыва страницы и приступить к поиску следующего разрыва страницы.
        rng.Move Unit:=wdCharacter, Count:=1
    Loop
   
    '3. Удаление разрывов страниц.
    '1) Удаление разрывов страниц в файлах формата "doc" (это "Word 2003").
        ' В старой версии для разрыва страницы не создавался отдельный абзац.
    If wa.ActiveDocument.SaveFormat = wdFormatDocument Then
        With wa.ActiveDocument.Range.Find
            .Text = "^m"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
        End With
    '2) Удаление разрывов страниц в файлах нового формата ("Word 2007+").
        ' В новых версиях разрыв страницы помещается в отдельный абзац. Если просто
            ' удалить разрыв страницы, то останется лишний знак абзаца. Поэтому нужно удалять не
            ' просто разрыв страницы, а разрыв страницы и знак абзаца.
        ' Применять такой поиск: .Text = "^m^p" к doc-формату нельзя, т.к.
            ' если после разрыва страницы есть пустой абзац, то пустой абзац будет удалён.
    Else
        With wa.ActiveDocument.Range.Find
            .Text = "^m^p"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
        End With
    End If
   
    '4. Включение монитора.
'    Application.ScreenUpdating = True
   
   
   
   
    '5. Сообщение.
'    MsgBox "Готово.", vbInformation


End Sub
