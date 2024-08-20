'------------------------------------------------------------------------------------------------------------
' Module        : Specifikaciya - Спецификация
' Author        : gtfox
' Date          : 2019.09.22(Спецификация)/2022.05.11(Кабельный журнал)
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
Option Explicit
Dim tabl(1 To 1000, 1 To 9) As Variant
Public arr() As Variant
Dim str As Integer
Dim pNumber As Integer
Dim RowCountXls As Integer
Dim ColoumnCountXls As Integer
Dim xx As Integer
Dim yx As Integer
Dim pth As String
Public Excel_imya_lista As String
Public sp As Excel.Workbook
Public frmClose As Boolean

'------------------------------------------------------------------------------------------------------------
'--------------------------------------Спецификация + перечень элементов-------------------------------------
'------------------------------------------------------------------------------------------------------------
'---------ShowSpecifikaciya
'---------spDEL
'---------SP_Excel_2_Visio
'---------PE_Excel_2_Visio
'---------xls_query
'---------fill_table_SP
'---------fill_table_PE
'---------AddPageSpecifikac
'---------SP_EXP_2_XLS
'---------PE_EXP_2_XLS
'---------get_data
'---------SortNumInString
'---------ReplaceSequenceInString
'---------PozNameInString
'---------AddSostavNaboraIzBD

Sub ShowSpecifikaciya()
    frmMenuSpecifikaciya.Show
End Sub

Public Sub spDEL()
'------------------------------------------------------------------------------------------------------------
' Macros        : spDEL - Удаляет листы спецификации
'------------------------------------------------------------------------------------------------------------
    If MsgBox("Удалить листы спецификации?", vbQuestion + vbOKCancel, "САПР-АСУ: Удалить спецификацию") = vbOK Then
        del_pages cListNameSpec
        ActiveDocument.DocumentSheet.Cells("user.SA_FR_NListSpecifikac").Formula = 0
        'MsgBox "Старая версия спецификации удалена", vbInformation
    End If
End Sub

Public Sub SP_Excel_2_Visio()
'------------------------------------------------------------------------------------------------------------
' Macros        : SP_Excel_2_Visio - Создает спецификацию из Excel в Visio
'------------------------------------------------------------------------------------------------------------
    xls_query "A3:I"
    If frmClose Then Exit Sub
    fill_table_SP
    Application.ActiveWindow.Page = Application.ActiveDocument.Pages.Item(cListNameSpec)
    MsgBox "Спецификация добавлена", vbInformation, "САПР-АСУ: Info"
End Sub

Public Sub PE_Excel_2_Visio(PerechenElementov As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : PE_Excel_2_Visio - Создает перечень элементов из Excel в Visio
'------------------------------------------------------------------------------------------------------------
    xls_query "A3:I"
    If frmClose Then Exit Sub
    fill_table_PE PerechenElementov
End Sub

Private Sub xls_query(strRange As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : xls_query - Заполняет массив данными из Excel
'------------------------------------------------------------------------------------------------------------
    Dim oExcel As Excel.Application
'    Dim sp As Excel.Workbook
'    Dim sht As Excel.Worksheet
    Dim tr As Object
    Dim tc As Object
    Dim qx As Integer
    Dim qy As Integer
    Dim ffs As FileDialogFilters
    Dim sFileName As String
    Dim fd As FileDialog
    Dim sPath, sFile As String
    Dim Chois As Integer
    Dim lLastRow As Long
    
    
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
            .Add "Excel", "*.xls;*.xlsx"
        End With
        Chois = oExcel.FileDialog(msoFileDialogOpen).Show
    End With
    If Chois = 0 Then oExcel.Application.Quit: frmClose = True: Exit Sub
    sFileName = oExcel.FileDialog(msoFileDialogOpen).SelectedItems(1)
    
    

    sPath = pth
'    sFileName = "SP_2_Visio.xls"
    sFile = sFileName
    
'    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
'        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "Ошибка"
'        Exit Sub
'    End If
    
    Set sp = oExcel.Workbooks.Open(sFile)
    Load frmMenuVyborListaExcel
    frmMenuVyborListaExcel.run sp 'присваиваем Excel_imya_lista

    If frmClose Then oExcel.Application.Quit: Exit Sub

    sp.Activate
    Dim UserRange As Excel.Range
    Dim Total As Excel.Range ' диапазон Full_list
    
    On Error Resume Next
    If oExcel.Worksheets(Excel_imya_lista) Is Nothing Then
        'действия, если листа нет
'        oExcel.run "'SP_2_Visio.xls'!Spec_2_Visio.Spec_2_Visio" 'создаем
    Else
        'действия, если лист есть
    End If
    
    'oExcel.GoTo Reference:=sp.Worksheets(1).Range("A2")
    'oExcel.ActiveCell.Select
    lLastRow = oExcel.Sheets(Excel_imya_lista).Cells(oExcel.Sheets(Excel_imya_lista).Rows.Count, 1).End(xlUp).Row
    Set UserRange = oExcel.Worksheets(Excel_imya_lista).Range(strRange & lLastRow)
    
'    Set UserRange = oExcel.InputBox _
'    (Prompt:="Выберите диапазон A3:Ix", _
'    Title:="Выбор диапазона", _
'    Type:=8)
    
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
    sp.Close savechanges:=False
    oExcel.Application.Quit
    

End Sub

Private Sub fill_table_SP()
'------------------------------------------------------------------------------------------------------------
' Macros        : fill_table_SP - Заполняет листы спецификации данными из массива
'------------------------------------------------------------------------------------------------------------
    Dim ncell As Integer
    Dim NStrokiXls As Integer
    Dim NRow As Integer ' счетчик количества строк спецификации на странице
    Dim mastSpecifikacia As Master
    Dim pName As String
    Dim shpCell As Shape
    Dim shpSpecifikacia As Shape
    Dim shpRow As Shape
    Dim HMax As Integer
    Dim HTable As Integer
    Dim Ramka As Visio.Shape
    Dim xNCell As Integer
    Dim Index As Integer

    ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").FormulaU = 1
    NRow = 1
    Set Ramka = ActivePage.Shapes.Item("Рамка")
    Set mastSpecifikacia = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("СП")
    ActivePage.Drop mastSpecifikacia, 0, 0
    Set shpSpecifikacia = ActivePage.Shapes.Item("СП")
    For NStrokiXls = 1 To RowCountXls
        Set shpRow = shpSpecifikacia.Shapes.Item("row" & NRow)
        For ncell = 1 To 9 'ColoumnCountXls
            Set shpCell = shpRow.Shapes.Item(NRow & "." & ncell)
            shpCell.text = arr(NStrokiXls, ncell)
            If ncell = 2 Or ncell = 9 Then shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
            If ncell = 2 And arr(NStrokiXls, 1) = "" Then
                shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "1" 'По центру
                shpCell.CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = visItalic + visUnderLine 'Курсив+Подчеркивание
            End If
        Next ncell

        DoEvents

        If Ramka.Cells("User.N").Result(0) = 3 Then HMax = 198 Else HMax = 232
        HTable = shpSpecifikacia.Cells("User.V").Result("mm")
        
        If HTable > HMax Then 'Высота таблицы больше 232мм/198мм
            'Удаляем лишние строки
            While HTable > HMax
                For xNCell = 1 To 9 'ColoumnCountXls
                    Set shpCell = shpRow.Shapes.Item(NRow & "." & xNCell)
                    shpCell.text = " "
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
    RowCountXls = 0
    Exit Sub
    
SubAddPage:
    'Добавляем лист
    NRow = 0
    ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").FormulaU = ActiveDocument.DocumentSheet.Cells("User.SA_FR_NListSpecifikac").Result(0) + 1
    'Положение текущей страницы
    Index = ActivePage.Index
    'Создаем новую страницу КЖ
    ActiveWindow.Page = AddSAPage(cListNameSpec)
    'Положение новой страницы сразу за текущей
    ActivePage.Index = Index + 1
    Set Ramka = ActivePage.Shapes.Item("Рамка")
    Ramka.Shapes("FORMA3").Shapes("Shifr").Cells("fields.value").FormulaU = "=TheDoc!User.SA_FR_Shifr & "".CO"""
    Ramka.Cells("User.NomerLista").FormulaU = "=PAGENUMBER()+Sheet.1!Prop.CNUM + TheDoc!User.SA_FR_NListSpecifikac - PAGECOUNT()"
    Ramka.Cells("User.ChisloListov").FormulaU = "=TheDoc!User.SA_FR_NListSpecifikac"
    ActivePage.Drop mastSpecifikacia, 0, 0
    Set shpSpecifikacia = ActivePage.Shapes.Item("СП")
    Return

 End Sub
 
 
 Private Sub fill_table_PE(PerechenElementov As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : fill_table_PE - Заполняет листы перечня элементов данными из массива
'------------------------------------------------------------------------------------------------------------
    Dim ncell As Integer
    Dim NStrokiXls As Integer
    Dim NRow As Integer ' счетчик количества строк спецификации на странице
    Dim shpCell As Shape
    Dim shpRow As Shape

    NRow = 1
    If RowCountXls > 30 Then MsgBox "Строк на листе Excel больше, чем строк в таблице(30): " & RowCountXls & vbNewLine & vbNewLine & "Разбейте перечень на несколько таблиц", vbExclamation, "САПР-АСУ: Перечень элементов": RowCountXls = 30
    For NStrokiXls = 1 To RowCountXls
        Set shpRow = PerechenElementov.Shapes.Item("row" & NRow)
        For ncell = 1 To 4 'ColoumnCountXls
            Set shpCell = shpRow.Shapes.Item(NRow & "." & ncell)
            shpCell.text = arr(NStrokiXls, ncell)
            If ncell = 2 Or ncell = 4 Then shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
            If ncell = 2 And arr(NStrokiXls, 1) = "" Then
                shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "1" 'По центру
                shpCell.CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = visItalic + visUnderLine 'Курсив+Подчеркивание
            End If
        Next ncell
        NRow = NRow + 1
        If NRow > 30 Then NRow = 0
    Next NStrokiXls
    RowCountXls = 0
End Sub

 
Sub AddPageSpecifikac(pName As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : AddPageSpecifikac - Добавляет пустую страницу спецификации
'------------------------------------------------------------------------------------------------------------
    Dim aPage As Visio.Page
    Dim mStr As Visio.Master
    Dim Ramka As Visio.Shape
    If GetSAPageExist(pName) Is Nothing Then
        Set aPage = ActiveDocument.Pages.Add
        aPage.name = pName
        With aPage.PageSheet
            .Cells("PageWidth").Formula = "420 MM"
            .Cells("PageHeight").Formula = "297 MM"
            .Cells("Paperkind").Formula = 8
            .Cells("PrintPageOrientation").Formula = 2
        End With
        SetPageAction aPage
        Set mStr = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
        Set Ramka = ActivePage.Drop(mStr, 0, 0)
        LockTitleBlock
'        ActiveDocument.Masters.Item("Рамка").Delete
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
    If pName Like cListNameKJ & ".*" Then Ramka.Cells("Prop.CHAPTER").FormulaU = "INDEX(1,Prop.CHAPTER.Format)"
    Ramka.Cells("Prop.cnum") = 0
    Ramka.Cells("Prop.tnum") = 0

End Sub

Public Sub SP_EXP_2_XLS()
'------------------------------------------------------------------------------------------------------------
' Macros        : SP_EXP_2_XLS - Экспортирует данные из таблицы спецификации в Excel
'------------------------------------------------------------------------------------------------------------
    Dim opn As Long
    Dim npName As String
    Dim pName As String
    Dim np As Page
    Dim pg As Page
    Dim n As Integer
    pName = cListNameSpec
    str = 1
    opn = ActiveDocument.Pages.Item(pName).Index
    Application.ActiveWindow.Page = ActiveDocument.Pages.Item(cListNameSpec)
    get_data ActivePage.Shapes.Item("СП"), 9
    For n = 2 To ActiveDocument.DocumentSheet.Cells("user.SA_FR_NListSpecifikac")
        pName = cListNameSpec & "." & n
        Application.ActiveWindow.Page = ActiveDocument.Pages.Item(pName)
        get_data ActivePage.Shapes.Item("СП"), 9
    Next
    
    Dim apx As Excel.Application
    Set apx = CreateObject("Excel.Application")
    Dim wb As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    
    Dim sPath, sFile As String
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set wb = apx.Workbooks.Open(sFile)
    

    
    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "СП_" & un & ".xls"
    apx.Visible = True
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets("СП_EXP_2_XLS").Delete
    apx.DisplayAlerts = True
    'Отключаем On Error Resume Next
    err.Clear
    On Error GoTo 0
    'добавляем новый
    apx.Sheets("СП").Visible = True
    apx.Sheets("СП").Copy After:=apx.Sheets(apx.Worksheets.Count)
    apx.Sheets("СП").Visible = False
    apx.Sheets("СП (2)").name = "СП_EXP_2_XLS"
    
    
    lLastRow = apx.Sheets("СП_EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets("СП_EXP_2_XLS").Activate
    apx.ActiveSheet.Rows("6:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A3:I5").ClearContents
    apx.ActiveSheet.Rows("5:" & str).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    wb.Activate
        
        For xx = 1 To str + 2
            For yx = 1 To 9
                wb.Sheets("СП_EXP_2_XLS").Cells(xx + 2, yx) = tabl(xx, yx)
                'wb.Sheets("СП_EXP_2_XLS").Range("A" & (xx + 2)).Select 'для наглядности
            Next yx
        Next xx
        
    apx.ActiveSheet.Range("A3:I" & apx.Sheets("СП_EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row).WrapText = False
    apx.ActiveSheet.Range("A3:I" & apx.Sheets("СП_EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид перед копированием
    apx.ActiveSheet.Range("K1") = Format(Now(), "yyyy.mm.dd hh:mm:ss")
    apx.ActiveSheet.Range("K1").Select
    wb.Save
'    WB.Close SaveChanges:=True
'    apx.Quit
    MsgBox "Спецификация экспортирована в файл SP_2_Visio.xls на лист СП_EXP_2_XLS", vbInformation, "САПР-АСУ: Info"
End Sub

Public Sub PE_EXP_2_XLS(PerechenElementov As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : PE_EXP_2_XLS - Экспортирует данные из таблицы перечень элементов в Excel
'------------------------------------------------------------------------------------------------------------
    Dim opn As Long
    Dim npName As String
    Dim pName As String
    Dim NameListExcel As String
    Dim np As Page
    Dim pg As Page
    Dim n As Integer
    pName = PerechenElementov.ContainingPage.name
    NameListExcel = "ПЭ_" & pName & "_EXP_2_XLS"
    str = 1
    Erase tabl
    get_data PerechenElementov, 4
    
    Dim apx As Excel.Application
    Set apx = CreateObject("Excel.Application")
    Dim wb As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    
    Dim sPath, sFile As String
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set wb = apx.Workbooks.Open(sFile)

    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "СП_" & un & ".xls"
    apx.Visible = True
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets(NameListExcel).Delete
    apx.DisplayAlerts = True
    'Отключаем On Error Resume Next
    err.Clear
    On Error GoTo 0
    'добавляем новый
    apx.Sheets("СП").Visible = True
    apx.Sheets("СП").Copy After:=apx.Sheets(apx.Worksheets.Count)
    apx.Sheets("СП").Visible = False
    apx.Sheets("СП (2)").name = NameListExcel
    
    
    lLastRow = apx.Sheets(NameListExcel).Cells(apx.Rows.Count, 1).End(xlUp).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets(NameListExcel).Activate
    apx.ActiveSheet.Rows("6:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A3:I5").ClearContents
    apx.ActiveSheet.Rows("5:" & str).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    apx.ActiveSheet.Range("A1") = "Поз."
    apx.ActiveSheet.Range("B1") = "Наименование"
    apx.ActiveSheet.Range("C1") = "Кол."
    apx.ActiveSheet.Range("D1") = "Примечание"
    apx.ActiveSheet.Columns("E:I").Delete
   
    wb.Activate
        
        For xx = 1 To str + 2
            For yx = 1 To 4
                wb.Sheets(NameListExcel).Cells(xx + 2, yx) = tabl(xx, yx)
                'wb.Sheets(NameListExcel).Range("A" & (xx + 2)).Select 'для наглядности
            Next yx
        Next xx
        
    apx.ActiveSheet.Range("A1:I" & apx.Sheets(NameListExcel).Cells(apx.Rows.Count, 1).End(xlUp).Row).WrapText = False
    apx.ActiveSheet.Range("A3:I" & apx.Sheets(NameListExcel).Cells(apx.Rows.Count, 1).End(xlUp).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид перед копированием
    
    apx.ActiveSheet.Range("B3:B" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).HorizontalAlignment = xlLeft
    apx.ActiveSheet.Range("D3:D" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).HorizontalAlignment = xlLeft
    apx.ActiveSheet.Range("F1") = Format(Now(), "yyyy.mm.dd hh:mm:ss")
    apx.ActiveSheet.Range("A1:D" & apx.ActiveSheet.Cells(apx.Rows.Count, 1).End(xlDown).Row).Columns.AutoFit
    apx.ActiveSheet.Range("F1").Select
    
    
    
    wb.Save
'    WB.Close SaveChanges:=True
'    apx.Quit
'    MsgBox "Спецификация экспортирована в файл SP_2_Visio.xls на лист ПЭ_EXP_2_XLS", vbInformation
End Sub

Public Sub get_data(Tablica As Visio.Shape, kolcell As Integer)
'------------------------------------------------------------------------------------------------------------
' Macros        : get_data - Собирает данные из таблицы Visio в массив
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim c As Integer
    Dim rw As Shape     ' шейп - строка
    Dim cn As String ' имя целевого шейпа
    
    For i = 1 To 30
        Set rw = Tablica.Shapes.Item("row" & i)
        If rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight) = 0 Then GoTo out:
        For c = 1 To kolcell
            cn = i & "." & c
             tabl(str, c) = rw.Shapes.Item(cn).text
        Next
        str = str + 1
    Next
out:
End Sub

Public Function SortNumInString(strToSort As String) As String
'------------------------------------------------------------------------------------------------------------
' Function      : SortNumInString - "Сортировка вставками" чисел в строке, разделенных ";"
                'Строка чисел, разделенных ";", преобразуется в массив, сортируется,
                'и возвращается в виде склеенной строки
'------------------------------------------------------------------------------------------------------------
    Dim mNum() As String
    Dim NumTemp As Variant
    Dim i As Integer
    Dim j As Integer
    Dim UbNum As Long
    
    mNum = Split(strToSort, ";")
    UbNum = UBound(mNum)
    If UbNum > 0 Then
        strToSort = ""
        '--V--Сортировка
        For j = 1 To UbNum
            NumTemp = IIf(mNum(j) = "", "0", mNum(j))
            i = j
            While CInt(mNum(i - 1)) > CInt(NumTemp) '>:возрастание, <:убывание
                mNum(i) = mNum(i - 1)
                i = i - 1
                If i <= 0 Then GoTo ExitWhile
            Wend
ExitWhile:     mNum(i) = NumTemp
        Next
        '--Х--Сортировка
        For i = 0 To UbNum
            strToSort = strToSort & mNum(i) & ";"
        Next
        strToSort = Left(strToSort, Len(strToSort) - 1)
    End If
    SortNumInString = strToSort
End Function


Public Function ReplaceSequenceInString(strToReplace As String) As String
'------------------------------------------------------------------------------------------------------------
' Function      : ReplaceSequenceInString - Заменяет последовательно идущие чисела в строке на тире
                ' "1;2;3;4;5;9" заменяется на "1-;5;9"
                'и возвращается в виде склеенной строки
'------------------------------------------------------------------------------------------------------------
    Dim mNum() As String
    Dim NumTemp As Variant
    Dim i As Integer
    Dim j As Integer
    Dim NumStart As Integer
    Dim NumEnd As Integer
    Dim TempStart As Integer
    Dim nCount As Integer
    Dim UbNum As Long
    
    mNum = Split(strToReplace, ";")
    strToReplace = ""
    UbNum = UBound(mNum)
    For i = 0 To UbNum
        NumStart = CInt(IIf(mNum(i) = "", "0", mNum(i)))
        TempStart = NumStart
        For j = i To UbNum 'Сканируем диапазон
            If j = UbNum Then '--------достигли конца строки---------
                If TempStart - NumStart > 0 Then 'конец = диапазон
                    If TempStart - NumStart = 1 Then
                        strToReplace = strToReplace & NumStart & ";" & TempStart & ";" 'конец = диапазон - 2 цифры
                    Else
                        strToReplace = strToReplace & NumStart & "-;" & TempStart & ";" 'конец = диапазон - больше 2-х цифр
                    End If
                Else
                    strToReplace = strToReplace & TempStart & ";"  'конец = единичное число
                End If
                i = j
                Exit For
            End If
            NumEnd = CInt(mNum(j + 1))
            If NumEnd - TempStart = 1 Then 'идет последовательность
                TempStart = NumEnd
                nCount = nCount + 1
            Else '---------------Конец последовательности-------------------
                If nCount = 0 Then
                    strToReplace = strToReplace & TempStart & ";" 'нет последовательности
                ElseIf nCount = 1 Then
                    strToReplace = strToReplace & NumStart & ";" & TempStart & ";" 'диапазон - 2 цифры
                Else
                    strToReplace = strToReplace & NumStart & "-;" & TempStart & ";" 'диапазон - больше 2-х цифр
                End If
                nCount = 0
                i = j
                Exit For
            End If
        Next
    Next

    strToReplace = Left(strToReplace, Len(strToReplace) - 1)

    ReplaceSequenceInString = strToReplace
End Function

Public Function PozNameInString(strPozNumber As String, strPozName As String) As String
'------------------------------------------------------------------------------------------------------------
' Function      : PozNameInString - Добавляет ИМЕНА позиционных обозначений к НОМЕРАМ позиционных обозначений
                'Строка чисел, разделенных ";", преобразуется в массив, добавляются имена,
                'и возвращается в виде склеенной строки разделенной ","
'------------------------------------------------------------------------------------------------------------
    Dim mNum() As String
    Dim i As Integer
    Dim UbNum As Long
    
    mNum = Split(strPozNumber, ";")
    UbNum = UBound(mNum)
    If UbNum > -1 Then
        strPozNumber = ""
        For i = 0 To UbNum
            strPozNumber = strPozNumber & strPozName & mNum(i) & IIf(InStr(mNum(i), "-"), "", ",")
        Next
        strPozNumber = Left(strPozNumber, Len(strPozNumber) - 1)
    End If
    PozNameInString = strPozNumber
End Function


Sub ReplaceNaborToSostav(colStrokaSpecifInner As Collection)
'------------------------------------------------------------------------------------------------------------
' Sub      : ReplaceNaborToSostav - Заменяет строку спецификации с набором на состав набора из БД
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim NElemNabora As Double
    
    For i = 1 To colStrokaSpecifInner.Count
        If colStrokaSpecifInner(i).ArtikulDB Like "Набор_*" Then
            NElemNabora = AddSostavNaboraIzExcelBD(colStrokaSpecifInner, colStrokaSpecifInner(i).KolVo, colStrokaSpecifInner(i).ArtikulDB, i)
            ReplaceNaborToSostav colStrokaSpecifInner
        End If
    Next
End Sub

Public Function AddSostavNaboraIzExcelBD(colStrokaSpecif As Collection, KolVo As Integer, strNabor As String, iIndex As Integer) As Double
'------------------------------------------------------------------------------------------------------------
' Function      : AddSostavNaboraIzExcelBD - Добавляет состав набора из БД к списку позиций спецификации
                'Возвращает число добавленных строк
'------------------------------------------------------------------------------------------------------------
    Dim oConn As New ADODB.Connection
    Dim oRecordSet As New ADODB.Recordset
    Dim SQLQuery As String
    Dim clsStrokaSpecif As classStrokaSpecifikacii
    Dim strColKey As String
    Dim nCount As Long
    Dim oldnCount As Double
    Dim SymName As String
    Dim PozOboznach As String
    
    SymName = colStrokaSpecif(iIndex).SymName
    PozOboznach = colStrokaSpecif(iIndex).PozOboznach
    colStrokaSpecif.Remove iIndex
    
    nCount = colStrokaSpecif.Count
    oldnCount = nCount
    
    SQLQuery = "SELECT * FROM [" & ExcelNabory & "$] WHERE Набор='" & strNabor & "';"
    
    
    oConn.Mode = adModeReadWrite
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sSAPath & DBNameIzbrannoeExcel & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    oRecordSet.CursorType = adOpenStatic
    oRecordSet.Open SQLQuery, oConn
    
    With oRecordSet
        If .RecordCount > 0 Then
            If .EOF Then .Close: Exit Function
            Do Until .EOF

                Set clsStrokaSpecif = New classStrokaSpecifikacii
                clsStrokaSpecif.SymName = SymName
                clsStrokaSpecif.SAType = ""
                clsStrokaSpecif.NazvanieDB = IIf(IsNull(.Fields(2 - 1).Value), "", .Fields(2 - 1).Value) 'Название
                clsStrokaSpecif.ArtikulDB = IIf(IsNull(.Fields(1 - 1).Value), "", .Fields(1 - 1).Value) 'Артикул
                clsStrokaSpecif.ProizvoditelDB = IIf(IsNull(.Fields(5 - 1).Value), "", .Fields(5 - 1).Value) 'Производитель
                clsStrokaSpecif.CenaDB = IIf(IsNull(.Fields(3 - 1).Value), "", .Fields(3 - 1).Value) 'Цена
                clsStrokaSpecif.EdDB = IIf(IsNull(.Fields(4 - 1).Value), "", .Fields(4 - 1).Value) 'Единица
                clsStrokaSpecif.KolVo = IIf(IsNull(.Fields(6 - 1).Value), "", .Fields(6 - 1).Value) * KolVo 'Количество
                clsStrokaSpecif.PozOboznach = PozOboznach
                clsStrokaSpecif.KodPoziciiDB = ""
                strColKey = SymName & PozOboznach & ";;" & IIf(IsNull(.Fields(1 - 1).Value), "", .Fields(1 - 1).Value) 'Артикул
                        
                On Error Resume Next
                colStrokaSpecif.Add clsStrokaSpecif, strColKey
                err.Clear
                On Error GoTo 0
                If colStrokaSpecif.Count = nCount Then 'Если кол-во не увеличелось, значит уже есть такой элемент
                    MsgBox "В наборе присутствуют позиции с одинаковым артикулом: " & SymName & PozOboznach & " : " & IIf(IsNull(.Fields(1 - 1).Value), "", .Fields(1 - 1).Value), vbExclamation, "САПР-АСУ: Добавление набора в состав спецификации"
                Else
                    nCount = colStrokaSpecif.Count
                End If

                .MoveNext
            Loop
        End If
    End With

    AddSostavNaboraIzExcelBD = nCount - oldnCount 'oRecordSet.RecordCount
    
    oRecordSet.Close
    oConn.Close
    Set oRecordSet = Nothing
    Set oConn = Nothing

End Function

'Access

'Public Function AddSostavNaboraIzBD(colStrokaSpecif As Collection, KolVo As Integer, IzbPozCod As String, iIndex As Integer) As Double
''------------------------------------------------------------------------------------------------------------
'' Function      : AddSostavNaboraIzBD - Добавляет состав набора из БД к списку позиций спецификации
'                'Возвращает число добавленных строк
''------------------------------------------------------------------------------------------------------------
'    Dim i As Double
'    Dim iold As Double
'    Dim rst As DAO.Recordset
'    Dim RecordCount As Double
'    Dim SQLQuery As String
'    Dim clsStrokaSpecif As classStrokaSpecifikacii
'    Dim strColKey As String
'    Dim nCount As Long
'
'    nCount = colStrokaSpecif.Count
'    SQLQuery = "SELECT Наборы.КодПозиции, Наборы.ИзбрПозицииКод, Наборы.Артикул, Наборы.Название, Наборы.Цена, Наборы.Количество, Наборы.ПроизводительКод, Производители.Производитель, Наборы.ЕдиницыКод, Единицы.Единица " & _
'                "FROM Единицы INNER JOIN (Производители INNER JOIN Наборы ON Производители.КодПроизводителя = Наборы.ПроизводительКод) ON Единицы.КодЕдиницы = Наборы.ЕдиницыКод " & _
'                "WHERE Наборы.ИзбрПозицииКод=" & IzbPozCod & ";"
'    Set rst = GetRecordSet(DBNameIzbrannoeAccess, SQLQuery)
'    If rst.RecordCount > 0 Then
'        rst.MoveLast
'        RecordCount = rst.RecordCount
'        i = 0
'        iold = 1000
'        With rst
'            If .EOF Then Exit Function
'            .MoveFirst
'            Do Until .EOF
'                Set clsStrokaSpecif = New classStrokaSpecifikacii
'                clsStrokaSpecif.SymName = colStrokaSpecif(iIndex).SymName
'                clsStrokaSpecif.SAType = ""
'                clsStrokaSpecif.NazvanieDB = .Fields("Название").Value
'                clsStrokaSpecif.ArtikulDB = .Fields("Артикул").Value
'                clsStrokaSpecif.ProizvoditelDB = .Fields("Производитель").Value
'                clsStrokaSpecif.CenaDB = .Fields("Цена").Value
'                clsStrokaSpecif.EdDB = .Fields("Единица").Value
'                clsStrokaSpecif.KolVo = .Fields("Количество").Value * KolVo
'                clsStrokaSpecif.PozOboznach = colStrokaSpecif(iIndex).PozOboznach
'                clsStrokaSpecif.KodPoziciiDB = ""
'                strColKey = ";;" & .Fields("Артикул").Value
'
'                On Error Resume Next
'                colStrokaSpecif.Add clsStrokaSpecif, strColKey
'                If colStrokaSpecif.Count = nCount Then 'Если кол-во не увеличелось, значит уже есть такой элемент
'                    MsgBox "В наборе присутствуют позиции с одинаковым артикулом: " & .Fields("Артикул").Value, vbExclamation, "САПР-АСУ: Добавление набора в состав спецификации"
'                Else
'                    nCount = colStrokaSpecif.Count
'                End If
'
'                .MoveNext
'            Loop
'        End With
'        AddSostavNaboraIzBD = RecordCount
'    End If
'    Set rst = Nothing
'End Function

'------------------------------------------------------------------------------------------------------------
'----------------------------------------------Кабельный журнал----------------------------------------------
'------------------------------------------------------------------------------------------------------------
'---------KJ_Excel_2_Visio
'---------kjDEL
'---------fill_table_KJ
'---------AddPageKJ
'---------KJ_EXP_2_XLS
'---------GetTrassa
'---------KJColToArray

Public Sub KJ_Excel_2_Visio()
'------------------------------------------------------------------------------------------------------------
' Macros        : KJ_Excel_2_Visio - Создает кабельный журнал из Excel в Visio
'------------------------------------------------------------------------------------------------------------
    Dim vsoPage As Visio.Page
    Set vsoPage = ActivePage
    xls_query "A4:G"
    If frmClose Then Exit Sub
    fill_table_KJ
    Application.ActiveWindow.Page = vsoPage 'Application.ActiveDocument.Pages.Item(cListNameKJ)
    MsgBox "Кабельный журнал добавлен", vbInformation, "САПР-АСУ: Info"
End Sub

Public Sub kjDEL()
'------------------------------------------------------------------------------------------------------------
' Macros        : kjDEL - Удаляет листы кабельного журнала
'------------------------------------------------------------------------------------------------------------
    If MsgBox("Удалить листы кабельного журнала?", vbQuestion + vbOKCancel, "САПР-АСУ: Удалить кабельный журнал") = vbOK Then
        del_pages cListNameKJ
        'MsgBox "Старая версия спецификации удалена", vbInformation
    End If
End Sub

Public Sub fill_table_KJ()
'------------------------------------------------------------------------------------------------------------
' Macros        : fill_table_KJ - Заполняет листы кабельного журнала данными из массива
'------------------------------------------------------------------------------------------------------------
    Dim TheDocListovSpecifikac As Cell
    Dim ncell As Integer
    Dim NStrokiXls As Integer
    Dim NRow As Integer ' счетчик количества строк кабельного журнала на странице
    Dim mastKJ As Master
    Dim pName As String
    Dim shpCell As Shape
    Dim shpKJ As Shape
    Dim shpRow As Shape
    Dim HMax As Integer
    Dim HTable As Integer
    Dim mStr() As String
    Dim Ramka As Visio.Shape
    Dim Index As Integer
    Dim xNCell As Integer

    Set Ramka = ActivePage.Shapes.Item("Рамка")
    NRow = 1

    Set mastKJ = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("КЖ")
    ActivePage.Drop mastKJ, 0, 0
    Set shpKJ = ActivePage.Shapes.Item("КЖ")
    For NStrokiXls = 1 To UBound(arr, 1)
        Set shpRow = shpKJ.Shapes.Item("row" & NRow)
        For ncell = 1 To 7 'ColoumnCountXls
            Set shpCell = shpRow.Shapes.Item(NRow & "." & ncell)
            If ncell = 7 Then
                shpCell.text = Round(arr(NStrokiXls, ncell), 1)
            ElseIf ncell = 5 Then
                shpCell.Cells("Char.FontScale").Formula = "70%"
                shpCell.text = arr(NStrokiXls, ncell)
            Else
                shpCell.text = arr(NStrokiXls, ncell)
            End If
'            If ncell = 2 Or ncell = 9 Then shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
'            If ncell = 2 And arr(NStrokiXls, 1) = "" Then
'                shpCell.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "1" 'По центру
'                shpCell.CellsSRC(visSectionCharacter, 0, visCharacterStyle).FormulaU = visItalic + visUnderLine 'Курсив+Подчеркивание
'            End If
        Next ncell

        DoEvents

        If Ramka.Cells("User.N").Result(0) = 3 Then HMax = 198 Else HMax = 232
        HTable = shpKJ.Cells("User.V").Result("mm")
        
        If HTable > HMax Then 'Высота таблицы больше 232мм/198мм
            'Удаляем лишние строки
            While HTable > HMax
                For xNCell = 1 To 7 'ColoumnCountXls
                    Set shpCell = shpRow.Shapes.Item(NRow & "." & xNCell)
                    shpCell.text = " "
'                    shpCell.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "0 mm"
                Next xNCell
                NStrokiXls = NStrokiXls - 1
                NRow = NRow - 1
                Set shpRow = shpKJ.Shapes.Item("row" & NRow)
                HTable = shpKJ.Cells("User.V").Result("mm")
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
    RowCountXls = 0
    Exit Sub
    
SubAddPage: 'Добавляем лист
    NRow = 0
    'Положение текущей страницы
    Index = ActivePage.Index
    'Создаем новую страницу КЖ
    ActiveWindow.Page = AddSAPage(cListNameKJ)
    'Положение новой страницы сразу за текущей
    ActivePage.Index = Index + 1
    Set Ramka = ActivePage.Shapes.Item("Рамка")
    ActivePage.Drop mastKJ, 0, 0
    Set shpKJ = ActivePage.Shapes.Item("КЖ")
    Return

 End Sub

Public Sub KJ_EXP_2_XLS()
'------------------------------------------------------------------------------------------------------------
' Macros        : KJ_EXP_2_XLS - Экспортирует данные из таблицы кабельного журнала в Excel
'------------------------------------------------------------------------------------------------------------
    Dim opn As Long
    Dim npName As String
    Dim pName As String
    Dim np As Page
    Dim pg As Page
    Dim n As Integer
    Dim m As Integer
    pName = cListNameKJ
    str = 1
    opn = ActiveDocument.Pages.Item(pName).Index
    Application.ActiveWindow.Page = ActiveDocument.Pages.Item(cListNameKJ)
    get_data ActivePage.Shapes.Item("КЖ"), 7
    'находим все листы кабельного журнала
    For Each pg In ActiveDocument.Pages
        If pg.name Like cListNameKJ & ".*" Then
            m = m + 1
        End If
    Next
    For n = 2 To m
        pName = cListNameKJ & "." & n
        Application.ActiveWindow.Page = ActiveDocument.Pages.Item(pName)
        get_data ActivePage.Shapes.Item("КЖ"), 7
    Next
    
    Dim apx As Excel.Application
    Set apx = CreateObject("Excel.Application")
    Dim wb As Excel.Workbook
    Dim sht As Excel.Sheets
    Dim en As String
    Dim un As String
    Dim lLastRow As Long
    Dim nstr As Long
    
    Dim sPath, sFile, sFileName As String
    sPath = Visio.ActiveDocument.path
    sFileName = "SP_2_Visio.xls"
    sFile = sPath & sFileName
    
    
    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "САПР-АСУ: Ошибка"
        Exit Sub
    End If
    
    Set wb = apx.Workbooks.Open(sFile)

    'Set wb = apx.Workbooks.Add
    'un = Format(Now(), "yyyy_mm_dd")
    'pth = Visio.ActiveDocument.Path
    'en = pth & "СП_" & un & ".xls"
    apx.Visible = True
    'удаляем старый лист
    apx.DisplayAlerts = False
    On Error Resume Next
    apx.Sheets("КЖ_EXP_2_XLS").Delete
    apx.DisplayAlerts = True
    'Отключаем On Error Resume Next
    err.Clear
    On Error GoTo 0
    'добавляем новый
    apx.Sheets("КЖ").Visible = True
    apx.Sheets("КЖ").Copy After:=apx.Sheets(apx.Worksheets.Count)
    apx.Sheets("КЖ").Visible = False
    apx.Sheets("КЖ (2)").name = "КЖ_EXP_2_XLS"
    
    
    lLastRow = apx.Sheets("КЖ_EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlDown).Row
    apx.Application.CutCopyMode = False
    apx.Worksheets("КЖ_EXP_2_XLS").Activate
    apx.ActiveSheet.Rows("7:" & lLastRow).Delete Shift:=xlUp
    apx.ActiveSheet.Range("A4:J6").ClearContents
'    str = UBound(tabl, 1)
    If str < 5 Then nstr = 5 Else nstr = str
    apx.ActiveSheet.Rows("5:" & nstr).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    wb.Activate
        
        For xx = 1 To str
            For yx = 1 To 7
                If yx = 7 Then
                    wb.Sheets("КЖ_EXP_2_XLS").Cells(xx + 3, yx) = CSng(tabl(xx, yx))
                ElseIf yx = 2 Or yx = 3 Then
                    wb.Sheets("КЖ_EXP_2_XLS").Cells(xx + 3, yx) = " " & tabl(xx, yx)
                Else
                    wb.Sheets("КЖ_EXP_2_XLS").Cells(xx + 3, yx) = tabl(xx, yx)
                End If
                'wb.Sheets("КЖ_EXP_2_XLS").Range("A" & (xx + 2)).Select 'для наглядности
            Next yx
        Next xx
        
    apx.ActiveSheet.Range("A4:I" & apx.Sheets("КЖ_EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row).WrapText = False
    apx.ActiveSheet.Range("A4:I" & apx.Sheets("КЖ_EXP_2_XLS").Cells(apx.Rows.Count, 1).End(xlUp).Row).RowHeight = 20 'Если ячейки, в которых были многострочные тексты, были растянуты по высоте, то мы их приводим в нормальный вид перед копированием
    apx.ActiveSheet.Range("K3") = Format(Now(), "yyyy.mm.dd hh:mm:ss")
    apx.ActiveSheet.Range("K3").Select
    wb.Save
'    WB.Close SaveChanges:=True
'    apx.Quit
    MsgBox "Кабельный журнал экспортирован в файл SP_2_Visio.xls на лист КЖ_EXP_2_XLS", vbInformation, "САПР-АСУ: Info"
End Sub



Public Sub KJColToArray(colCollection As Collection)
'------------------------------------------------------------------------------------------------------------
' Macros        : KJColToArray - Заполняет массив данными из колекции
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    ReDim arr(colCollection.Count, 7) As Variant
    For i = 1 To colCollection.Count
        arr(i, 1) = colCollection(i).Oboznach '1 Обозначение кабеля, провода
        arr(i, 2) = colCollection(i).Nachalo '2 Трасса - Начало
        arr(i, 3) = colCollection(i).Konec '3 Трасса - Конец
        arr(i, 4) = colCollection(i).Trassa '4 Участок трассы кабеля, провода
        arr(i, 5) = colCollection(i).Marka '5 Кабель, провод - по проекту - Марка
        arr(i, 6) = colCollection(i).Sechenie '6 Кабель, провод - по проекту - Кол., число и сечение жил
        arr(i, 7) = colCollection(i).Dlina '7 Кабель, провод - по проекту - Длина, м.
    Next
End Sub