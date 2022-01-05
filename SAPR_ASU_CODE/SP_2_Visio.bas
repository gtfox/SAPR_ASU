'------------------------------------------------------------------------------------------------------------
' Module        : SP - Спецификация
' Author        : gtfox
' Date          : 2019.09.22
' Description   : spDEL - Удаляет листы спецификации
                ' spADD_Excel_Razbienie - Добавляет листы спецификации из Excel из листа SP_2_Visio (после разбиения на ячейки)
                ' spADD_Visio_Perenos - Добавляет листы спецификации из Excel из листа SP (перенос длинных строк делает визио)
                ' spEXP_2_XLS – Экспортирует спецификацию из Visio в Excel
                ' SP_2_Visio.xls -Спецификация
                ' Лист SP содержит исходную спецификацию, с многострочным текстом в одной ячейке.
                ' Лист SP_2_Visio создается автоматически и содержит лист SP с порезанными однострочными ячейками (в дальнейшем не используется + перезаписывается при каждом вызове макроса)
                ' Лист EXP_2_XLS содержит экспортированную из Visio спецификацию (если вы вносили изменения в спецификацию в самом Visio) Создается автоматически макросом + перезаписывается при каждом вызове макроса
                ' Основная проблема спецификации – длинные строки, которые надо разбивать/переносить.
                ' У Surrogate перенос делает Visio, а расчет высоты получившейся строки делает ShapeSheet. В 2007 версии VBA выполняется быстрее пересчета формул в ShapeSheet и макрос не получает высоту вовремя. Исправлено
                ' Я решил разбивать строки в Excel, и тогда в Visio не надо считать высоту через ShapeSheet.
                ' Деление многострочной ячейки на строки происходит на основе особенности реализации шейпа надпись в Excel. Задаем ширину прямоугольника и помещаем длинный текст. Он переносится, чтобы поместится в ширину. А особенностью является то, что мы можем обращаться отдельно к каждой получившейся строке в этом прямоугольнике через коллекции.
                ' Макрос написан на основе singleTextCellToRows https://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=77447&TITLE_SEO=77447-perenos-teksta-na-sleduyushchuyu-stroku-pri-zapolnenii-stolbtsa-po-shi&MID=841387#message841387
' Link          : https://visio.getbb.ru/viewtopic.php?p=14130, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе этого:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : speka2003 Спецификация
                ' Author    : Surrogate
                ' Date      : 07.11.2012
                ' Purpose   : Спецификация: перенос данных из Excel из Visio и обратно
                '           : Мастер для переноса данных из экселя в визио, для формирования спецификации
                ' Links     : https://visio.getbb.ru/viewtopic.php?f=15&t=234, https://visio.getbb.ru/download/file.php?id=106
                '------------------------------------------------------------------------------------------------------------

Option Base 1
'Option Explicit
Dim tabl(1 To 1000, 1 To 9) As Variant
Dim arr() As Variant
Dim str As Integer
Dim pNumber As Integer
Dim RowCountXls As Integer
Dim ColoumnCountXls As Integer
Dim xx As Integer
Dim yx As Integer
Dim pth As String

Public Sub spDEL()
    If MsgBox("Удалить листы спецификации?", vbQuestion + vbOKCancel, "Удалить спецификацию") = vbOK Then
        del_sp
        'MsgBox "Старая версия спецификации удалена", vbInformation
    End If
End Sub

'Public Sub spDEL_ADD()
'    del_sp
'    spADD
'End Sub

'Public Sub spADD_Excel_Razbienie()
'    xls_query "SP_2_Visio"
'    fill_table False
'    Application.ActiveWindow.Page = Application.ActiveDocument.Pages.Item(cListNameSpec)
'    MsgBox "Спецификация добавлена", vbInformation
'End Sub

Public Sub spADD_Visio_Perenos()
    xls_query "SP"
    fill_table
    Application.ActiveWindow.Page = Application.ActiveDocument.Pages.Item(cListNameSpec)
    MsgBox "Спецификация добавлена", vbInformation
End Sub

Private Sub xls_query(imya_lista As String)
    Dim oExcel As Excel.Application
    Dim sp As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim tr As Object
    Dim tc As Object
    Dim qx As Integer
    Dim qy As Integer
    Dim ffs As FileDialogFilters
    Dim sFileName As String
    Dim fd As FileDialog
    Dim sPath, sFile As String
    
    
    Set oExcel = CreateObject("Excel.Application")
    pth = Visio.ActiveDocument.path
'    oExcel.Visible = True ' для наглядности
    
    Set fd = oExcel.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = False
        .InitialFileName = pth
        Set ffs = .Filters
        With ffs
            .Clear
            .Add "Excel", "*.xls"
        End With
        oExcel.FileDialog(msoFileDialogOpen).Show
    End With
    sFileName = oExcel.FileDialog(msoFileDialogOpen).SelectedItems(1)

    

    sPath = pth
'    sFileName = "SP_2_Visio.xls"
    sFile = sFileName
    
'    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
'        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "Ошибка"
'        Exit Sub
'    End If
    
    Set sp = oExcel.Workbooks.Open(sFile)
    sp.Activate
    Dim UserRange As Excel.Range
    Dim Total As Excel.Range ' диапазон Full_list
    
    On Error Resume Next
    If oExcel.Worksheets(imya_lista) Is Nothing Then
        'действия, если листа нет
'        oExcel.run "'SP_2_Visio.xls'!Spec_2_Visio.Spec_2_Visio" 'создаем
    Else
        'действия, если лист есть
    End If
    
    'oExcel.GoTo Reference:=sp.Worksheets(1).Range("A2")
    'oExcel.ActiveCell.Select
    lLastRow = oExcel.Sheets(imya_lista).Cells(oExcel.Sheets(imya_lista).Rows.Count, 1).End(xlUp).Row
    Set UserRange = oExcel.Worksheets(imya_lista).Range("A3:I" & lLastRow) 'oExcel.InputBox _
    '(Prompt:="Выберите диапазон A3:Ix", _
    'Title:="Выбор диапазона", _
    'Type:=8)
    Set Total = UserRange
        For Each tr In Total.Rows
            RowCountXls = RowCountXls + 1
            ColoumnCountXls = 0
            For Each tc In Total.Rows.Columns
                ColoumnCountXls = ColoumnCountXls + 1
            Next tc
        Next tr
    ReDim arr(RowCountXls, ColoumnCountXls) As Variant
    For qx = 1 To RowCountXls
        For qy = 1 To ColoumnCountXls
            arr(qx, qy) = Total.Cells(qx, qy) ' заполнение массива arr
        Next qy
    Next qx
    sp.Close SaveChanges:=False
    oExcel.Application.Quit
    

End Sub

Private Sub fill_table()  ' заполнение спецификации

    Dim TheDocListovSpecifikac As Cell
    Dim NCell As Integer
    Dim NStrokiXls As Integer
    Dim NRow As Integer ' счетчик количества строк спецификации на странице
    Dim mastSpecifikacia As Master
    Dim pName As String
    Dim shpCell As Shape
    Dim shpSpecifikacia As Shape
    Dim shpRow As Shape
    Dim HMax As Integer
    Dim HTable As Integer

    Set TheDocListovSpecifikac = ActiveDocument.DocumentSheet.Cells("user.SA_FR_NListSpecifikac")
    TheDocListovSpecifikac.FormulaU = 1
    pNumber = 1
    NRow = 1
    pName = cListNameSpec & "."
    AddPageSpecifikac cListNameSpec
    Set mastSpecifikacia = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Спецификация")
    ActivePage.Drop mastSpecifikacia, 0, 0
    Set shpSpecifikacia = ActivePage.Shapes.Item("Спецификация")
    For NStrokiXls = 1 To RowCountXls
        Set shpRow = shpSpecifikacia.Shapes.Item("row" & NRow)
        For NCell = 1 To 9 'ColoumnCountXls
            Set shpCell = shpRow.Shapes.Item(NRow & "." & NCell)
            shpCell.Text = arr(NStrokiXls, NCell)
            If NCell = 2 Or NCell = 9 Then shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
            If NCell = 2 And arr(NStrokiXls, 1) = "" Then
                shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "1" 'По центру
                shpCell.CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = visItalic + visUnderLine 'Курсив+Подчеркивание
            End If
        Next NCell

        DoEvents

        If pNumber = 1 Then HMax = 198 Else HMax = 232
        HTable = shpSpecifikacia.Cells("User.V").Result("mm")
        
        If HTable > HMax Then 'Высота таблицы больше 232мм/198мм
            'Удаляем лишние строки
            While HTable > HMax
                For xNCell = 1 To 9 'ColoumnCountXls
                    Set shpCell = shpRow.Shapes.Item(NRow & "." & xNCell)
                    shpCell.Text = " "
'                    shpCell.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "0 mm"
                Next xNCell
                NStrokiXls = NStrokiXls - 1
                NRow = NRow - 1
                Set shpRow = shpSpecifikacia.Shapes.Item("row" & NRow)
                HTable = shpSpecifikacia.Cells("User.V").Result("mm")
            Wend
            'Добавляем лист
            GoSub SubAddPage

        ElseIf HTable = HMax And NStrokiXls <> RowCountXls Then 'Высота таблицы равна 232мм/198мм и это не полследняя строка
            'Добавляем лист
            GoSub SubAddPage
        End If
        
        NRow = NRow + 1
        If NRow > 30 Then NRow = 0

    Next NStrokiXls
    pNumber = 1
    RowCountXls = 0
    Exit Sub
    
SubAddPage:
    'Добавляем лист
    NRow = 0
    pNumber = pNumber + 1
    TheDocListovSpecifikac.Formula = pNumber
    AddPageSpecifikac pName & pNumber
    ActivePage.Drop mastSpecifikacia, 0, 0
    Set shpSpecifikacia = ActivePage.Shapes.Item("Спецификация")
    Return

 End Sub
 
Sub AddPageSpecifikac(pName As String)
    Dim aPage As Visio.Page
    Dim Mstr As Visio.Master
    Dim Ramka As Visio.Shape
    If GetSAPageExist(pName) Is Nothing Then
        Set aPage = ActiveDocument.Pages.Add
        aPage.Name = pName
        With aPage.PageSheet
            .Cells("PageWidth").Formula = "420 MM"
            .Cells("PageHeight").Formula = "297 MM"
            .Cells("Paperkind").Formula = 8
            .Cells("PrintPageOrientation").Formula = 2
            .AddSection visSectionAction
            .AddRow visSectionAction, visRowLast, visTagDefault
            .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Перечень оборудования со Схемы в Excel"""
            .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PagePLANAddElementsFrm"")"
            .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "263" '5897
            .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """10"""
            .AddRow visSectionAction, visRowLast, visTagDefault
            .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Создать спецификацию в Visio из Excel"""
            .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""spADD_Visio_Perenos"")"
            .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "7076" '6224
            .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """20"""
            .AddRow visSectionAction, visRowLast, visTagDefault
            .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Удалить все листы спецификации"""
            .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""spDEL"")"
            .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1088" '2645
            .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """30"""
        End With
        Set Mstr = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
        Set Ramka = ActivePage.Drop(Mstr, 0, 0)
        LockTitleBlock
        ActiveDocument.Masters.Item("Рамка").Delete
    Else
        ActiveWindow.Page = ActiveDocument.Pages(pName)
        ActiveWindow.SelectAll
        ActiveWindow.Selection.Delete
        Set Ramka = ActivePage.Shapes.Item("Рамка")
    End If
    Ramka.Shapes("FORMA3").Shapes("Shifr").Cells("fields.value").FormulaU = "=TheDoc!User.SA_FR_Shifr & "".CO"""
    Ramka.Cells("User.NomerLista").FormulaU = "=PAGENUMBER()+Sheet.1!Prop.CNUM + TheDoc!User.SA_FR_NListSpecifikac - PAGECOUNT()"
    Ramka.Cells("User.ChisloListov").FormulaU = "=TheDoc!User.SA_FR_NListSpecifikac"
'    Ramka.Cells("prop.type").Formula = """Спецификация оборудования, изделий и материалов"""
    If Len(pName) > 1 Then Ramka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"
    Ramka.Cells("Prop.cnum") = 0
    Ramka.Cells("Prop.tnum") = 0

End Sub

Private Sub del_sp()
    Dim dp As Page
    Dim colPage As Collection
    Set colPage = New Collection
    'Спецификацию в колекцию
    For Each dp In ActiveDocument.Pages
        If InStr(1, dp.Name, cListNameSpec & ".") > 0 Then
            colPage.Add dp
        End If
    Next
    'удаляем все страницы которые нашли выше
    For Each dp In colPage
        dp.Delete (1)
    Next
    On Error Resume Next
    ActiveDocument.Pages.Item(cListNameSpec).Delete (1)
    ActiveDocument.DocumentSheet.Cells("user.SA_FR_NListSpecifikac").Formula = 0
End Sub

Public Sub spEXP_2_XLS()
    Dim opn As Long
    Dim npName As String
    Dim pName As String
    Dim np As Page
    Dim pg As Page
    Dim N As Integer
    pName = cListNameSpec
    str = 1
    opn = ActiveDocument.Pages.Item(pName).Index
    Application.ActiveWindow.Page = ActiveDocument.Pages.Item(cListNameSpec)
    get_data
    For N = 2 To ActiveDocument.DocumentSheet.Cells("user.SA_FR_NListSpecifikac")
        pName = cListNameSpec & "." & N
        Application.ActiveWindow.Page = ActiveDocument.Pages.Item(pName)
        get_data
    Next
    Dim apx As Excel.Application
    Set apx = CreateObject("Excel.Application")
    Dim WB As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    
    Dim sPath, sFile As String
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "Ошибка"
        Exit Sub
    End If
    
    Set WB = apx.Workbooks.Open(sFile)
    

    
    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "Спецификация_" & un & ".xls"
    apx.Visible = True
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets("EXP_2_XLS").Delete
    apx.DisplayAlerts = True
    'добавляем новый
    apx.Sheets("SP").Copy After:=apx.Sheets(apx.Worksheets.Count)
    apx.Sheets("SP (2)").Name = "EXP_2_XLS"
    
    
    lLastRow = apx.Sheets("EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets("EXP_2_XLS").Activate
    apx.ActiveSheet.Rows("6:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A3:I5").ClearContents
    apx.ActiveSheet.Rows("5:" & str).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    WB.Activate
        
        For xx = 1 To str + 2
            For yx = 1 To 9
                WB.Sheets("EXP_2_XLS").Cells(xx + 2, yx) = tabl(xx, yx)
                'wb.Sheets("EXP_2_XLS").Range("A" & (xx + 2)).Select 'для наглядности
            Next yx
        Next xx
        
    apx.ActiveSheet.Range("A3:I" & apx.Sheets("EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row).WrapText = False
    apx.ActiveSheet.Range("A3:I" & apx.Sheets("EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид перед копированием
   

    WB.Close SaveChanges:=True
    apx.Quit
    MsgBox "Спецификация экспортирована в файл SP_2_Visio.xls на лист EXP_2_XLS", vbInformation
End Sub

Function get_data() '(pgName As Page)
    Dim r As Integer
    Dim target As Shape
    Dim c As Integer
    Dim main As Shape   ' шейп - основная группа
    Dim rw As Shape     ' шейп - строка
    Dim rn As String      ' имя шейпа-строки
    Set main = ActivePage.Shapes.Item("Спецификация")
    Dim SSS As Shapes  ' подмножество шейпов основной группы
    Dim tn As String ' имя целевого шейпа
    Set SSS = main.Shapes
    For r = 1 To 30
        rn = "row" & r
        Set rw = SSS.Item(rn)
        If rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight) = 0 Then GoTo out:
        For c = 1 To 9
            tn = r & "." & c
            Set target = rw.Shapes.Item(tn)
            tabl(str, c) = target.Text
        Next c
        str = str + 1
    Next r
out:
End Function





