'------------------------------------------------------------------------------------------------------------
' Module        : DB_Excel - База данных прайс листов и избранного на основе Excel
' Author        : gtfox
' Date          : 2023.01.30
' Description   : База данных прайс листов, избранного и их обеспечение на основе Excel
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit
Public sSAPath As String
Public oExcelAppPrice As Excel.Application
Public oExcelAppIzbrannoe As Excel.Application
Public wbExcelPrice As Excel.Workbook
Public wshPrice As Excel.Worksheet
Public wbExcelIzbrannoe As Excel.Workbook
Public wshIzbrannoe As Excel.Worksheet
Public wshNabory As Excel.Worksheet
Public wshNastrojkiPrajsov As Excel.Worksheet
Public wshExcelEdinicyIzmereniya As Excel.Worksheet
Public mProizvoditel() As classProizvoditelBD
Public PriceSettings As classProizvoditelBD
Public IzbrannoeSettings As classProizvoditelBD
Public Const DBNameIzbrannoeExcel As String = "SAPR_ASU_Izbrannoe.xls" 'Имя файла избронного
Public Const ExcelNastrojkiPrajsov As String = "НастройкиПрайсов" 'Имя листа настроек производителей
Public Const ExcelIzbrannoe As String = "Избранное" 'Имя листа Избранное
Public Const ExcelNabory As String = "Наборы" 'Имя листа Наборы
Public Const ExcelEdinicyIzmereniya As String = "ЕдиницыИзмерения" 'Имя листа Единицы Измерения
Public Const ExcelTemp As String = "temp" 'Имя листа для временных данных
Public MaxColumn As Double
Public MinColumn As Double
Public RangePrice As Excel.Range
Public SA_nRows As Double
Public bBlock As Boolean
Public colProcessHandle As Collection

#If VBA7 Then
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

'Активация формы выбора элементов схемы из БД.
Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
    Set colProcessHandle = New Collection
    GetAllExcelProcess
    sSAPath = Visio.ActiveDocument.path
'    Load frmDBPriceAccess
'    frmDBPriceAccess.run vsoShape 'Передали его в форму
    Load frmDBPriceExcel
    frmDBPriceExcel.run vsoShape 'Передали его в форму
End Sub

Sub InitPricelExceDB()
'------------------------------------------------------------------------------------------------------------
' Macros        : InitPricelExceDB - Инициализирует переменные для доступа к Excel на форме прайсов
'------------------------------------------------------------------------------------------------------------

End Sub

Sub InitIzbrannoeExcelDB()
'------------------------------------------------------------------------------------------------------------
' Macros        : InitIzbrannoeExcelDB - Инициализирует переменные для доступа к Excel на форме избранного
'------------------------------------------------------------------------------------------------------------
    Set oExcelAppIzbrannoe = CreateObject("Excel.Application")
    oExcelAppIzbrannoe.WindowState = xlMinimized
    oExcelAppIzbrannoe.Visible = True
    Set wbExcelIzbrannoe = oExcelAppIzbrannoe.Workbooks.Open(sSAPath & DBNameIzbrannoeExcel)
    Set wshIzbrannoe = wbExcelIzbrannoe.Worksheets(ExcelIzbrannoe)
    Set wshNabory = wbExcelIzbrannoe.Worksheets(ExcelNabory)
    Set wshNastrojkiPrajsov = wbExcelIzbrannoe.Worksheets(ExcelNastrojkiPrajsov)
    Set wshExcelEdinicyIzmereniya = wbExcelIzbrannoe.Worksheets(ExcelEdinicyIzmereniya)
    Set IzbrannoeSettings = New classProizvoditelBD
        IzbrannoeSettings.StolbArtikul = 1
        IzbrannoeSettings.StolbNazvanie = 2
        IzbrannoeSettings.StolbCena = 3
        IzbrannoeSettings.StolbEd = 4
        IzbrannoeSettings.StolbKategoriya = 6
        IzbrannoeSettings.StolbGruppa = 7
        IzbrannoeSettings.StolbPodgruppa = 8
    
    FillExcel_mProizvoditel
    
End Sub

Sub FillExcel_mProizvoditel()
    Dim UserRange As Excel.Range
    Dim lLastRow As Long
    Dim i As Integer

    lLastRow = wshNastrojkiPrajsov.Cells(wshNastrojkiPrajsov.Rows.Count, 1).End(xlUp).Row
    Set UserRange = wshNastrojkiPrajsov.Range("A2:J" & lLastRow)
    
    'Заполняем массив mProizvoditel Производители из Excel как базы данных САПР-АСУ
    ReDim mProizvoditel(lLastRow - 2)
    For i = 1 To lLastRow - 1
        Set mProizvoditel(i - 1) = New classProizvoditelBD
        mProizvoditel(i - 1).Proizvoditel = UserRange.Cells(i, 1)
        mProizvoditel(i - 1).FileName = UserRange.Cells(i, 2)
        mProizvoditel(i - 1).NameListExcel = UserRange.Cells(i, 3)
        mProizvoditel(i - 1).StolbArtikul = UserRange.Cells(i, 4)
        mProizvoditel(i - 1).StolbNazvanie = UserRange.Cells(i, 5)
        mProizvoditel(i - 1).StolbCena = UserRange.Cells(i, 6)
        mProizvoditel(i - 1).StolbEd = UserRange.Cells(i, 7)
        mProizvoditel(i - 1).StolbKategoriya = UserRange.Cells(i, 8)
        mProizvoditel(i - 1).StolbGruppa = UserRange.Cells(i, 9)
        mProizvoditel(i - 1).StolbPodgruppa = UserRange.Cells(i, 10)
    Next
End Sub


Sub WizardAddPriceExcel(sProizvoditel As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : WizardAddPriceExcel - Мастер добавления прайс-листа Excel в виде базы данных САПР-АСУ
'------------------------------------------------------------------------------------------------------------
    Dim sFilePathName As String
    Dim fdFileDialog As FileDialog
    Dim fdFilters As FileDialogFilters
    Dim Chois As Integer
    Dim i As Integer
    Dim mRange() As String
    Dim mDialogString() As String
    Dim sDialogString As String
    Dim mVendorData(0 To 9) As Variant
    Dim lLastRow As Long
    Dim UserRange As Excel.Range
    Dim FindRange As Excel.Range

    InitExcelDB
  
    'Проверяем, что такого производителя нет в списке
    lLastRow = wshNastrojkiPrajsov.Cells(wshNastrojkiPrajsov.Rows.Count, 1).End(xlUp).Row
    Set UserRange = wshNastrojkiPrajsov.Range("A2:A" & lLastRow)

    Set FindRange = UserRange.Find(sProizvoditel, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not FindRange Is Nothing Then
        MsgBox "Такой производитель уже есть в списке: " & sProizvoditel, vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
        ExcelAppQuit oExcelAppIzbrannoe
        Exit Sub
    End If
    
    'Открываем прайс
    Set oExcelAppPrice = CreateObject("Excel.Application")
    Set fdFileDialog = oExcelAppPrice.FileDialog(msoFileDialogOpen)
    With fdFileDialog
        .AllowMultiSelect = False
        .InitialFileName = sSAPath
        Set fdFilters = .Filters
        With fdFilters
            .Clear
            .Add "Excel", "*.xls"
        End With
        Chois = oExcelAppPrice.FileDialog(msoFileDialogOpen).Show
    End With
    If Chois = 0 Then ExcelAppQuit oExcelAppIzbrannoe: ExcelAppQuit oExcelAppPrice:  frmClose = True: Exit Sub
    sFilePathName = oExcelAppPrice.FileDialog(msoFileDialogOpen).SelectedItems(1)
    
    If InStr(sFilePathName, sSAPath) = 1 Then 'файл в той же папке, что и проект (но может быть и глубже)
        sRelativeFileName = Replace(sFilePathName, sSAPath, "") 'относительный путь
    Else
        sRelativeFileName = sFilePathName 'абсолютный путь
    End If

    Set wbExcelPrice = oExcelAppPrice.Workbooks.Open(sFilePathName)
    Load frmVyborListaExcel
    frmVyborListaExcel.run wbExcelPrice 'присваиваем Excel_imya_lista

    If frmClose Then ExcelAppQuit oExcelAppIzbrannoe: ExcelAppQuit oExcelAppPrice: Exit Sub
    Set wshPrice = wbExcelPrice.Worksheets(Excel_imya_lista)
    oExcelAppPrice.Visible = True
    wshPrice.Activate
    
    'Строка Производителя на листе НастройкиПрайсов в файле SAPR_ASU_Izbrannoe.xls
    mVendorData(0) = sProizvoditel 'Производитель
    mVendorData(1) = sRelativeFileName 'ИмяФайлаБазы
    mVendorData(2) = Excel_imya_lista 'ИмяЛиста
    
    '0-6
    sDialogString = "Выберите ячейку в столбце ""Артикул""." & vbCrLf & "Будет выполнено преобразование Артикула в текст;" & _
                    "Выберите ячейку в столбце ""Название"";" & _
                    "Выберите ячейку в столбце ""Цена"";" & _
                    "Выберите ячейку в столбце ""Единица"";" & _
                    "Выберите ячейку в столбце ""Категория"";" & _
                    "Выберите ячейку в столбце ""Группа"";" & _
                    "Выберите ячейку в столбце ""Подгруппа"""

    mDialogString = Split(sDialogString, ";")

    For i = 0 To 6
        On Error GoTo err1
        Set UserRange = oExcelAppPrice.InputBox _
        (Prompt:=mDialogString(i), _
        Title:="Выбор ячейки", _
        Type:=8)
        err.Clear
        On Error GoTo 0
        mRange = Split(UserRange.Address, ":")
        If UBound(mRange) = 0 Then 'выбрана одна ячейка
'            mRange = Split(UserRange.Address, "$") 'буква столбца mRange(1) 'СтолбецАртикул/СтолбецНазвание/СтолбецЦена/СтолбецЕдиницы/СтолбецКатегория/СтолбецГруппа/СтолбецПодгруппа
            mVendorData(i + 3) = UserRange.Column 'СтолбецАртикул/СтолбецНазвание/СтолбецЦена/СтолбецЕдиницы/СтолбецКатегория/СтолбецГруппа/СтолбецПодгруппа
            'Преобразование Артикула в тип Текст
            
            If i = 0 Then
                oExcelAppPrice.WindowState = xlMinimized
                oExcelAppPrice.ScreenUpdating = False
                If MsgBox("Преобразовать ""Артикул"" к типу ТЕКСТ?" & vbCrLf & vbCrLf & "Если ""Артикул"" в Excel сохранён как ЧИСЛО то возможны проблемы с поиском" & vbCrLf & vbCrLf & "Дождитесь окончания процесса...", vbYesNo + vbInformation, "САПР-АСУ: Преобразовать в ТЕКСТ?") = vbYes Then
                    wshPrice.Range("A1").AutoFilter Field:=1
                    ExcelConvertToString wshPrice.Range(wshPrice.AutoFilter.Range.Columns(UserRange.Column).Address) 'напрямую передаваяя Columns не работало...
                End If
                oExcelAppPrice.ScreenUpdating = True
                oExcelAppPrice.WindowState = xlMaximized
            End If
        Else 'выбран диапазон
            oExcelAppPrice.WindowState = xlMinimized
            MsgBox "Был выбран диапазон ячеек!" & vbCrLf & vbCrLf & "Необходимо выбрать одну ячейку", vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
            i = i - 1
            oExcelAppPrice.WindowState = xlMaximized
        End If
    Next

    wbExcelPrice.Close savechanges:=True
    
    'Запись данных в лист НастройкиПрайсов
    
    wshNastrojkiPrajsov.Activate
    lLastRow = wshNastrojkiPrajsov.Cells(wshNastrojkiPrajsov.Rows.Count, 1).End(xlUp).Row
    For i = 1 To 10
        wshNastrojkiPrajsov.Cells(lLastRow + 1, i) = mVendorData(i - 1)
    Next
    oExcelAppIzbrannoe.Visible = True
    wbExcelIzbrannoe.Save
    Exit Sub
err1:
    ExcelAppQuit oExcelAppIzbrannoe
    ExcelAppQuit oExcelAppPrice
End Sub

'Заполняет lstvTable данными из БД в виде Excel через ADODB
 Function Fill_lstvTable(FileName As String, wshWorkSheet As Excel.Worksheet, lstvTable As ListView, PoizvoditelSettings As classProizvoditelBD, Optional ByVal TableType As Integer = 0) As String
    'TableType=1 - Избранное
    'TableType=2 - Набор
    Dim oConn As ADODB.Connection
    Dim oRecordSet As ADODB.Recordset
    Dim oExcelApp As Excel.Application
    Dim wshTemp As Excel.Worksheet
    Dim RangeSource As Excel.Range
    Dim sAddress As String
    Dim i As Double
    Dim j As Double
    Dim itmx As ListItem
    
    Set oConn = New ADODB.Connection
    Set oRecordSet = New ADODB.Recordset
    wshWorkSheet.Range("A1").AutoFilter Field:=1
    Set RangeSource = wshWorkSheet.AutoFilter.Range
    Set wshTemp = GetSheetExcel(wshWorkSheet.Parent, ExcelTemp)
    wshTemp.Cells.ClearContents
    RangeSource.Copy wshTemp.Cells(1, 1)
    wshTemp.Cells(1, 1).AutoFilter Field:=1
    sAddress = Replace(wshTemp.AutoFilter.Range.Address, "$", "")
    Set oExcelApp = wshWorkSheet.Parent.Parent
'    wshWorkSheet.Parent.Close SaveChanges:=True
'    oExcelApp.Quit
    
    ';Mode=Read
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & IIf(FileName Like "*:*", FileName, sSAPath & FileName) & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    oRecordSet.Open "SELECT * FROM [" & ExcelTemp & "$" & sAddress & "]", oConn
'    oRecordSet.Open "SELECT * FROM [" & ExcelTemp & "$" & sAddress & "] WHERE ", oConn
    
    lstvTable.ListItems.Clear
    With oRecordSet
        If .EOF Then .Close: Exit Function
        Do Until .EOF
            If i < SA_nRows Then
                Set itmx = lstvTable.ListItems.Add(, , IIf(IsNull(.Fields(PoizvoditelSettings.StolbArtikul - 1).Value), "", .Fields(PoizvoditelSettings.StolbArtikul - 1).Value)) 'Артикул
                itmx.SubItems(1) = IIf(IsNull(.Fields(PoizvoditelSettings.StolbNazvanie - 1).Value), "", .Fields(PoizvoditelSettings.StolbNazvanie - 1).Value) 'Название
                itmx.SubItems(2) = IIf(IsNull(.Fields(PoizvoditelSettings.StolbCena - 1).Value), "", .Fields(PoizvoditelSettings.StolbCena - 1).Value) 'Цена
                itmx.SubItems(3) = IIf(IsNull(.Fields(PoizvoditelSettings.StolbEd - 1).Value), "", .Fields(PoizvoditelSettings.StolbEd - 1).Value) 'Единица
                If TableType = 1 Then
                    itmx.SubItems(4) = IIf(IsNull(.Fields(5 - 1).Value), "", .Fields(5 - 1).Value) 'Производитель
                ElseIf TableType = 2 Then
                    itmx.SubItems(4) = IIf(IsNull(.Fields(5 - 1).Value), "", .Fields(5 - 1).Value) 'Производитель
                    itmx.SubItems(5) = IIf(IsNull(.Fields(6 - 1).Value), "", .Fields(6 - 1).Value) 'Количество
                End If
        
                'красим наборы
                If TableType = 1 Then
                    If IIf(IsNull(.Fields(PoizvoditelSettings.StolbArtikul - 1).Value), "", .Fields(PoizvoditelSettings.StolbArtikul - 1).Value) Like "Набор_*" Then
                        itmx.ForeColor = NaboryColor
                       'itmx.Bold = True
                        For j = 1 To itmx.ListSubItems.Count
                           'itmx.ListSubItems(j).Bold = True
                            itmx.ListSubItems(j).ForeColor = NaboryColor
                        Next
                    End If
                End If
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
    Fill_lstvTable = IIf(TableType = 2, i, IIf(i <= SA_nRows, i, i & ".  Показано: " & SA_nRows))
    oRecordSet.Close
    oConn.Close
    Set oRecordSet = Nothing
    Set oConn = Nothing
    Set wshTemp = Nothing
End Function

Public Sub RuleFilterCmbx(wshWorkSheet As Excel.Worksheet, RangeToFilter As Excel.Range, UserForm As MSForms.UserForm, PoizvoditelSettings As classProizvoditelBD, Ncmbx As Integer)
    Dim fltrMode As Integer
    
    '-------------------ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
    Select Case Ncmbx
        Case 1
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya, Criteria1:=UserForm.cmbxKategoriya 'Категория
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa 'Группа
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbPodgruppa 'Подгруппа
            UpdateCmbxFilters wshWorkSheet, UserForm.cmbxGruppa, PoizvoditelSettings.StolbGruppa
            UpdateCmbxFilters wshWorkSheet, UserForm.cmbxPodgruppa, PoizvoditelSettings.StolbPodgruppa
        Case 2
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa, Criteria1:=UserForm.cmbxGruppa 'Группа
            If UserForm.cmbxKategoriya.ListIndex = -1 Then
                RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya
                UpdateCmbxFilters wshWorkSheet, UserForm.cmbxKategoriya, PoizvoditelSettings.StolbKategoriya
            Else
                RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya, Criteria1:=UserForm.cmbxKategoriya 'Категория
            End If
            UpdateCmbxFilters wshWorkSheet, UserForm.cmbxPodgruppa, PoizvoditelSettings.StolbPodgruppa
        Case 3
            '-------------------ФИЛЬТРАЦИЯ Подгруппы при разных (Категория || Группа)------------------------------------------------
            '*    К   Гр
            '0    0   0
            '1    0   1
            '2    1   0
            '3    1   1
            
            fltrMode = IIf(UserForm.cmbxKategoriya.ListIndex = -1, 0, 2) + IIf(UserForm.cmbxGruppa.ListIndex = -1, 0, 1)
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbPodgruppa, Criteria1:=UserForm.cmbxPodgruppa 'Подгруппа
            Select Case fltrMode
                Case 0
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya 'Категория
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa 'Группа
                    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxKategoriya, PoizvoditelSettings.StolbKategoriya
                    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxGruppa, PoizvoditelSettings.StolbGruppa
                Case 1
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya 'Категория
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa, Criteria1:=UserForm.cmbxGruppa 'Группа
                    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxKategoriya, PoizvoditelSettings.StolbKategoriya
                Case 2
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya, Criteria1:=UserForm.cmbxKategoriya 'Категория
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa 'Группа
                    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxGruppa, PoizvoditelSettings.StolbGruppa
                Case 3
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya, Criteria1:=UserForm.cmbxKategoriya 'Категория
                    RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa, Criteria1:=UserForm.cmbxGruppa 'Группа
                Case Else
            End Select
            '-------------------/ФИЛЬТРАЦИЯ Подгруппы при разных (Категория || Группа)------------------------------------------------
        Case Else
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbKategoriya 'Категория
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbGruppa 'Группа
            RangeToFilter.AutoFilter Field:=PoizvoditelSettings.StolbPodgruppa 'Подгруппа
            UpdateAllCmbxFilters wshWorkSheet, UserForm, PoizvoditelSettings
    End Select
    '-------------------/ФИЛЬТРАЦИЯ С ПРИОРИТЕТОМ (По иерархии: Категория->Группа->Подгруппа)------------------------------------------------
End Sub

 Sub UpdateCmbxFilters(wshWorkSheet As Excel.Worksheet, cmbxComboBox As ComboBox, nColumn As Long)
    'nColumn = classProizvoditelBD.StolbKategoriya - Категория
    'nColumn = classProizvoditelBD.StolbGruppa - Группа
    'nColumn = classProizvoditelBD.StolbPodgruppa - Подгруппа
    Dim wshTemp As Excel.Worksheet
    Dim UserRange As Excel.Range
    Dim lLastRow As Long
    Dim i As Integer
    Dim sCmbx As String
    
    bBlock = True
    sCmbx = cmbxComboBox
    Set wshTemp = GetSheetExcel(wshWorkSheet.Parent, ExcelTemp)
    wshTemp.Cells.ClearContents
    lLastRow = wshWorkSheet.Cells(wshWorkSheet.Rows.Count, 1).End(xlUp).Row
    If lLastRow > 1 Then
        wshWorkSheet.Range(wshWorkSheet.Cells(2, nColumn), wshWorkSheet.Cells(lLastRow, nColumn)).Copy wshTemp.Cells(1, 1)
        Set UserRange = wshTemp.Range(wshTemp.Cells(1, 1), wshTemp.Cells(lLastRow - 1, 1))
        UserRange.RemoveDuplicates Columns:=1, Header:=xlNo
        lLastRow = wshTemp.Cells(wshTemp.Rows.Count, 1).End(xlUp).Row
        If lLastRow > 0 Then
            cmbxComboBox.Clear
            For i = 1 To lLastRow
                cmbxComboBox.AddItem wshTemp.Cells(i, 1)
            Next
        End If
    Else
        cmbxComboBox.Clear
    End If
    For i = 0 To cmbxComboBox.ListCount - 1
        If cmbxComboBox.List(i, 0) = sCmbx Then cmbxComboBox.ListIndex = i
    Next
    bBlock = False
    Set wshTemp = Nothing
End Sub

Public Sub UpdateAllCmbxFilters(wshWorkSheet As Excel.Worksheet, UserForm As MSForms.UserForm, PoizvoditelSettings As classProizvoditelBD)
    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxKategoriya, PoizvoditelSettings.StolbKategoriya
    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxGruppa, PoizvoditelSettings.StolbGruppa
    UpdateCmbxFilters wshWorkSheet, UserForm.cmbxPodgruppa, PoizvoditelSettings.StolbPodgruppa
End Sub

'Очистка фильтров
Public Sub ClearFilter(wshWorkSheet As Excel.Worksheet)
    wshWorkSheet.Range("A1").AutoFilter
    wshWorkSheet.Range("A1").AutoFilter Field:=1
End Sub

Public Sub ADODB_Excel_Connect(oConn As ADODB.Connection, FileName As String)
    On Error Resume Next
    oConn.Close
    err.Clear
    On Error GoTo 0
    Set oConn = Nothing
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Mode=Read;Data Source=" & IIf(FileName Like "*:*", FileName, sSAPath & FileName) & ";Extended Properties=""Excel 12.0;HDR=YES"";"
End Sub

Public Sub ADODB_Excel_RecordSet(oRecordSet As ADODB.Recordset, oConn As ADODB.Connection, SheetName As String, Table_SourceAddress As String)
    On Error Resume Next
    oRecordSet.Close
    err.Clear
    On Error GoTo 0
    Set oRecordSet = Nothing
    Set oRecordSet = New ADODB.Recordset
    oRecordSet.CursorType = adOpenStatic
    oRecordSet.Open "SELECT * FROM [" & SheetName & "$" & Table_SourceAddress & "]", oConn
End Sub

Public Sub FillExcel_cmbxProizvoditel(cmbx As ComboBox, Optional ByVal Price As Boolean = False)
'------------------------------------------------------------------------------------------------------------
' Macros        : FillExcel_cmbxProizvoditel - Заполняет ComboBox Производители из массива mProizvoditel
'------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    cmbx.Clear
    For i = 0 To UBound(mProizvoditel)
        If mProizvoditel(i).FileName = "" And Price Then
            'для формы Прайс пропускаем производителя, если у него нету файла
        Else
            cmbx.AddItem mProizvoditel(i).Proizvoditel
        End If
    Next
End Sub

Public Sub FillCmbxEdinicy(cmbxComboBox As ComboBox)
'------------------------------------------------------------------------------------------------------------
' Macros        : FillCmbxEdinicy - Заполняет ComboBox Единицы измерения из листа ЕдиницыИзмерения SAPR_ASU_Izbrannoe.xls
'------------------------------------------------------------------------------------------------------------
    Dim UserRange As Excel.Range
    Dim lLastRow As Long
    Dim i As Integer
    
    lLastRow = wshExcelEdinicyIzmereniya.Cells(wshExcelEdinicyIzmereniya.Rows.Count, 1).End(xlUp).Row
    Set UserRange = wshExcelEdinicyIzmereniya.Range("A2:A" & lLastRow)
    cmbxComboBox.Clear
    For i = 1 To lLastRow - 1
        cmbxComboBox.AddItem UserRange.Cells(i, 1)
    Next
End Sub

Public Sub ExcelConvertToString(ConvertRange As Excel.Range)
'------------------------------------------------------------------------------------------------------------
' Macros        : ExcelConvertToString - Преобразует диапазон ячеек Excel в текстовый тип данных для работы фильтра (стандартное преобразование в текст не работает)
'------------------------------------------------------------------------------------------------------------
    Dim text$
    Dim rCell As Excel.Range
    For Each rCell In ConvertRange
        text = WorksheetFunction.text(rCell.Value, rCell.NumberFormat)
        rCell.NumberFormat = "@"
        rCell.Value = text
    Next
End Sub

Public Sub FindArticulInBrowser(Artikul As String, NomerMagazina As Integer)
'------------------------------------------------------------------------------------------------------------
' Macros        : FindArticulInBrowser - Открывает браузер с поиском артикула товара в нужном магазине
'------------------------------------------------------------------------------------------------------------
    If Artikul = "" Then Exit Sub
    Select Case NomerMagazina
        Case 0 'ЭТМ
            CreateObject("WScript.Shell").run "https://www.etm.ru/catalog/?searchValue=" & Artikul
        Case 1 'АВС Электро
            CreateObject("WScript.Shell").run "https://avselectro.ru/search/index.php?q=" & Artikul
        Case Else
    End Select
End Sub

Public Function GetSheetExcel(wbExcel As Excel.Workbook, mySheetName As String) As Excel.Worksheet
    Dim mySheetNameTest As String
    
    On Error Resume Next
    mySheetNameTest = wbExcel.Worksheets(mySheetName).name
    If err.Number = 0 Then
        Set GetSheetExcel = wbExcel.Worksheets(mySheetName)
    Else
        err.Clear
        On Error GoTo 0
        wbExcel.Worksheets.Add.name = mySheetName
        Set GetSheetExcel = wbExcel.Worksheets(mySheetName)
    End If
End Function

Public Sub ExcelAppQuit(oExcelApp As Excel.Application)
    Dim wbWbExcel As Excel.Workbook
    On Error Resume Next
    For Each wbWbExcel In oExcelApp.Workbooks
        wbWbExcel.Close savechanges:=False
    Next
    Set wbExcelIzbrannoe = Nothing
    Set wbExcelPrice = Nothing
    oExcelApp.Application.Quit
    Set oExcelApp = Nothing
End Sub

'Собираем процессы Excel открытые не нами
Sub GetAllExcelProcess()
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        If Process.Caption Like "EXCEL.EXE" Then
            colProcessHandle.Add Process.Handle, Process.Handle
        End If
    Next
End Sub

'Убиваем процессы Excel открытые нами
Sub KillSAExcelProcess()
    Dim Process As Object
    Dim nCount As Double
    For Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
        If Process.Caption Like "EXCEL.EXE" Then
            nCount = colProcessHandle.Count
            On Error Resume Next
            colProcessHandle.Add Process.Handle, Process.Handle
            If colProcessHandle.Count > nCount Then 'Если кол-во увеличелось, значит че-то всунулось - его надо прибить
                colProcessHandle.Remove Process.Handle
                Process.Terminate
            End If
        End If
    Next
End Sub