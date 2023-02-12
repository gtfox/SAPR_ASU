'------------------------------------------------------------------------------------------------------------
' Module        : DB_Excel - База данных прайс листов и избранного на основе Excel
' Author        : gtfox
' Date          : 2023.01.30
' Description   : База данных прайс листов, избранного и их обеспечение на основе Excel
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit
Public sSAPath As String
Public oExcelApp As Excel.Application
Public wbExcelPrice As Excel.Workbook
Public wshPrice As Excel.Worksheet
Public wbExcelIzbrannoe As Excel.Workbook
Public wshIzbrannoe As Excel.Worksheet
Public wshNabory As Excel.Worksheet
Public wshNastrojkiPrajsov As Excel.Worksheet
Public wshExcelEdinicyIzmereniya As Excel.Worksheet
Public wshTemp As Excel.Worksheet
Public mProizvoditel() As classProizvoditelBD
Public CurentPrice As classProizvoditelBD
Public Const DBNameIzbrannoeExcel As String = "SAPR_ASU_Izbrannoe.xls" 'Имя файла избронного
Public Const ExcelNastrojkiPrajsov As String = "НастройкиПрайсов" 'Имя листа настроек производителей
Public Const ExcelIzbrannoe As String = "Избранное" 'Имя листа Избранное
Public Const ExcelNabory As String = "Наборы" 'Имя листа Наборы
Public Const ExcelEdinicyIzmereniya As String = "ЕдиницыИзмерения" 'Имя листа Единицы Измерения
Public Const ExcelTemp As String = "temp" 'Имя листа для временных данных
Public MaxColumn As Double
Public MinColumn As Double
Public RangePrice As Excel.Range


#If VBA7 Then
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

'Активация формы выбора элементов схемы из БД. Расположено в модуле DB_Access
'Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
''    Load frmDBPriceAccess
''    frmDBPriceAccess.run vsoShape 'Передали его в форму
'    Load frmDBPriceExcel
'    frmDBPriceExcel.run vsoShape 'Передали его в форму
'End Sub


Sub InitExcelDB()
'------------------------------------------------------------------------------------------------------------
' Macros        : InitExcelDB - Инициализирует переменные для доступа к Excel
'------------------------------------------------------------------------------------------------------------
    sSAPath = Visio.ActiveDocument.path
    Set oExcelApp = CreateObject("Excel.Application")
    Set wbExcelIzbrannoe = oExcelApp.Workbooks.Open(sSAPath & DBNameIzbrannoeExcel)
    Set wshIzbrannoe = wbExcelIzbrannoe.Worksheets(ExcelIzbrannoe)
    Set wshNabory = wbExcelIzbrannoe.Worksheets(ExcelNabory)
    Set wshNastrojkiPrajsov = wbExcelIzbrannoe.Worksheets(ExcelNastrojkiPrajsov)
    Set wshExcelEdinicyIzmereniya = wbExcelIzbrannoe.Worksheets(ExcelEdinicyIzmereniya)
    Set wshTemp = wbExcelIzbrannoe.Worksheets(ExcelTemp)
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
    Dim mVendorData(0 To 11) As Variant
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
        ExcelAppExit
        Exit Sub
    End If
    
    'Открываем прайс
    Set fdFileDialog = oExcelApp.FileDialog(msoFileDialogOpen)
    With fdFileDialog
        .AllowMultiSelect = False
        .InitialFileName = sSAPath
        Set fdFilters = .Filters
        With fdFilters
            .Clear
            .Add "Excel", "*.xls"
            .Add "Excel", "*.xlsx"
        End With
        Chois = oExcelApp.FileDialog(msoFileDialogOpen).Show
    End With
    If Chois = 0 Then oExcelApp.Quit: frmClose = True: Exit Sub
    sFilePathName = oExcelApp.FileDialog(msoFileDialogOpen).SelectedItems(1)
    
    If InStr(sFilePathName, sSAPath) = 1 Then 'файл в той же папке, что и проект (но может быть и глубже)
        sRelativeFileName = Replace(sFilePathName, sSAPath, "") 'относительный путь
    Else
        sRelativeFileName = sFilePathName 'абсолютный путь
    End If

    Set wbExcelPrice = oExcelApp.Workbooks.Open(sFilePathName)
    Load frmVyborListaExcel
    frmVyborListaExcel.run wbExcelPrice 'присваиваем Excel_imya_lista

    If frmClose Then ExcelAppExit: Exit Sub
    Set wshPrice = wbExcelPrice.Worksheets(Excel_imya_lista)
    oExcelApp.Visible = True
    wshPrice.Activate
    
    'Строка Производителя на листе НастройкиПрайсов в файле SAPR_ASU_Izbrannoe.xls
    mVendorData(0) = sProizvoditel 'Производитель
    mVendorData(1) = sRelativeFileName 'ИмяФайлаБазы
    mVendorData(2) = Excel_imya_lista 'ИмяЛиста
    
    '0-8
    sDialogString = "Выберите начальную ячейку данных прайса (Ctrl+Home);" & _
                    "Выберите конечную ячейку данных прайса (Ctrl+End);" & _
                    "Выберите ячейку в столбце ""Артикул""." & vbCrLf & "Будет выполнено преобразование Артикула в текст" & vbCrLf & "Дождитесь окончания процесса...;" & _
                    "Выберите ячейку в столбце ""Название"";" & _
                    "Выберите ячейку в столбце ""Цена"";" & _
                    "Выберите ячейку в столбце ""Единица"";" & _
                    "Выберите ячейку в столбце ""Категория"";" & _
                    "Выберите ячейку в столбце ""Группа"";" & _
                    "Выберите ячейку в столбце ""Подгруппа"""

    mDialogString = Split(sDialogString, ";")

    For i = 0 To 8
        Set UserRange = oExcelApp.InputBox _
        (Prompt:=mDialogString(i), _
        Title:="Выбор ячейки", _
        Type:=8)
    
        mRange = Split(UserRange.Address, ":")
        If UBound(mRange) = 0 Then 'выбрана одна ячейка
            If i < 2 Then
    '            mRange = Split(UserRange.Address, "$")'номер строки mRange(2) 'Строка начало/СтрокаКонец
                mVendorData(i + 3) = UserRange.Row 'Строка начало/СтрокаКонец
            Else
    '            mRange = Split(UserRange.Address, "$") 'буква столбца mRange(1) 'СтолбецАртикул/СтолбецНазвание/СтолбецЦена/СтолбецЕдиницы/СтолбецКатегория/СтолбецГруппа/СтолбецПодгруппа
                mVendorData(i + 3) = UserRange.Column 'СтолбецАртикул/СтолбецНазвание/СтолбецЦена/СтолбецЕдиницы/СтолбецКатегория/СтолбецГруппа/СтолбецПодгруппа
                'Преобразование Артикула в тип Текст
                oExcelApp.ScreenUpdating = False
                If i = 2 Then
                    ExcelConvertToString wshPrice.Range(wshPrice.Cells(mVendorData(3), mVendorData(5)), wshPrice.Cells(mVendorData(4), mVendorData(5)))
                End If
                oExcelApp.ScreenUpdating = True
            End If
        Else 'выбран диапазон
            oExcelApp.WindowState = xlMinimized
            MsgBox "Был выбран диапазон ячеек!" & vbCrLf & vbCrLf & "Необходимо выбрать одну ячейку", vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
            i = i - 1
            oExcelApp.WindowState = xlMaximized
        End If
    Next

    wbExcelPrice.Close SaveChanges:=True
    
    'Запись данных в лист НастройкиПрайсов
    
    wshNastrojkiPrajsov.Activate
    lLastRow = wshNastrojkiPrajsov.Cells(wshNastrojkiPrajsov.Rows.Count, 1).End(xlUp).Row
    For i = 1 To 12
        wshNastrojkiPrajsov.Cells(lLastRow + 1, i) = mVendorData(i - 1)
    Next
    oExcelApp.Visible = True
    wbExcelIzbrannoe.Save

End Sub


Public Sub FillExcel_mProizvoditel()
'------------------------------------------------------------------------------------------------------------
' Macros        : FillExcel_mProizvoditel - Заполняет массив mProizvoditel Производители из Excel как базы данных САПР-АСУ
'------------------------------------------------------------------------------------------------------------
    Dim UserRange As Excel.Range
    Dim i As Integer

    lLastRow = wshNastrojkiPrajsov.Cells(wshNastrojkiPrajsov.Rows.Count, 1).End(xlUp).Row
    Set UserRange = wshNastrojkiPrajsov.Range("A2:L" & lLastRow)
    
    ReDim mProizvoditel(lLastRow - 2)
    For i = 1 To lLastRow - 1
        Set mProizvoditel(i - 1) = New classProizvoditelBD
        mProizvoditel(i - 1).Proizvoditel = UserRange.Cells(i, 1)
        mProizvoditel(i - 1).FileName = UserRange.Cells(i, 2)
        mProizvoditel(i - 1).NameListExcel = UserRange.Cells(i, 3)
        mProizvoditel(i - 1).FirstRow = UserRange.Cells(i, 4)
        mProizvoditel(i - 1).LastRow = UserRange.Cells(i, 5)
        mProizvoditel(i - 1).Artikul = UserRange.Cells(i, 6)
        mProizvoditel(i - 1).Nazvanie = UserRange.Cells(i, 7)
        mProizvoditel(i - 1).Cena = UserRange.Cells(i, 8)
        mProizvoditel(i - 1).Ed = UserRange.Cells(i, 9)
        mProizvoditel(i - 1).Kategoriya = UserRange.Cells(i, 10)
        mProizvoditel(i - 1).Gruppa = UserRange.Cells(i, 11)
        mProizvoditel(i - 1).Podgruppa = UserRange.Cells(i, 12)
    Next
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
    
'    wbExcelIzbrannoe.Close SaveChanges:=False
'    oExcelApp.Quit
    oExcelApp.Visible = True
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

Public Sub ExcelAppExit()
    If Not wbExcelIzbrannoe Is Nothing Then wbExcelIzbrannoe.Close SaveChanges:=False
    Set wbExcelIzbrannoe = Nothing
    If Not wbExcelPrice Is Nothing Then wbExcelPrice.Close SaveChanges:=False
    Set wbExcelPrice = Nothing
    oExcelApp.Application.Quit
    Set oExcelApp = Nothing
End Sub