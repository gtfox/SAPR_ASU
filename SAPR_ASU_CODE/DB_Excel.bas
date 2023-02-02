'------------------------------------------------------------------------------------------------------------
' Module        : DB_Excel - База данных прайс листов и избранного на основе Excel
' Author        : gtfox
' Date          : 2023.01.30
' Description   : База данных прайс листов, избранного и их обеспечение на основе Excel
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit

Public Const DBNameIzbrannoeExcel As String = "SAPR_ASU_Izbrannoe.xls" 'Имя файла избронного
Public Const ExcelNastrojkiPrajsov As String = "НастройкиПрайсов" 'Имя листа настроек производителей

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


Sub WizardAddPriceExcel(sNameVendor As String)
'------------------------------------------------------------------------------------------------------------
' Macros        : WizardAddPriceExcel - Мастер добавления прайс-листа Excel в виде базы данных САПР-АСУ
'------------------------------------------------------------------------------------------------------------
    Dim oExcelApp As Excel.Application
    Dim wbExcel As Excel.Workbook
    Dim sPath As String
    Dim sFilePathName As String
    Dim fdFileDialog As FileDialog
    Dim fdfFilters As FileDialogFilters
    
    Dim Chois As Integer
    Dim i As Integer
    Dim mRange() As String
    Dim mDialogString() As String
    Dim sDialogString As String
    Dim mVendorData(0 To 11) As String

    Set oExcelApp = CreateObject("Excel.Application")
    sPath = Visio.ActiveDocument.path
    
    Set fdFileDialog = oExcelApp.FileDialog(msoFileDialogOpen)
    With fdFileDialog
        .AllowMultiSelect = False
        .InitialFileName = sPath
        Set fdfFilters = .Filters
        With fdfFilters
            .Clear
            .Add "Excel", "*.xls"
            .Add "Excel", "*.xlsx"
        End With
        Chois = oExcelApp.FileDialog(msoFileDialogOpen).Show
    End With
    If Chois = 0 Then oExcelApp.Application.Quit: frmClose = True: Exit Sub
    sFilePathName = oExcelApp.FileDialog(msoFileDialogOpen).SelectedItems(1)
    
    If InStr(sFilePathName, sPath) = 1 Then 'файл в той же папке, что и проект (но может быть и глубже)
        sRelativeFileName = Replace(sFilePathName, sPath, "") 'относительный путь
    Else
        sRelativeFileName = sFilePathName 'абсолютный путь
    End If

    Set wbExcel = oExcelApp.Workbooks.Open(sFilePathName)
    Load frmVyborListaExcel
    frmVyborListaExcel.Show 'присваиваем Excel_imya_lista
    If frmClose Then oExcelApp.Application.Quit: Exit Sub

    oExcelApp.Visible = True
    wbExcel.Activate
    
    'Строка Производителя на листе НастройкиПрайсов в файле SAPR_ASU_Izbrannoe.xls
    mVendorData(0) = sNameVendor 'Производитель
    mVendorData(1) = sRelativeFileName 'ИмяФайлаБазы
    mVendorData(2) = Excel_imya_lista 'ИмяЛиста
    
    '0-8
    sDialogString = "Выберите начальную ячейку данных прайса (Ctrl+Home);" & _
                    "Выберите конечную ячейку данных прайса (Ctrl+End);" & _
                    "Выберите ячейку в столбце ""Артикул"";" & _
                    "Выберите ячейку в столбце ""Название"";" & _
                    "Выберите ячейку в столбце ""Цена"";" & _
                    "Выберите ячейку в столбце ""Единица"";" & _
                    "Выберите ячейку в столбце ""Категория"";" & _
                    "Выберите ячейку в столбце ""Группа"";" & _
                    "Выберите ячейку в столбце ""Подгруппа"""

    mDialogString = Split(sDialogString, ";")

    For i = 0 To 1
        Set UserRange = oExcelApp.InputBox _
        (Prompt:=mDialogString(i), _
        Title:="Выбор ячейки", _
        Type:=8)
    
        mRange = Split(UserRange.Address, ":")
        If UBound(mRange) = 0 Then 'выбрана 1 ячейка
            mRange = Split(UserRange.Address, "$") 'номер строки
            mVendorData(i + 3) = mRange(2) 'Строка начало/СтрокаКонец
        Else 'выбран диапазон
            oExcelApp.WindowState = xlMinimized
            MsgBox "Был выбран диапазон ячеек!" & vbCrLf & vbCrLf & "Необходимо выбрать одну ячейку", vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
            i = i - 1
            oExcelApp.WindowState = xlMaximized
        End If
    Next

    For i = 2 To 8
        Set UserRange = oExcelApp.InputBox _
        (Prompt:=mDialogString(i), _
        Title:="Выбор ячейки", _
        Type:=8)
    
        mRange = Split(UserRange.Address, ":")
        If UBound(mRange) = 0 Then 'выбрана 1 ячейка
            mRange = Split(UserRange.Address, "$") 'буква столбца
            mVendorData(i + 3) = mRange(1) 'СтолбецАртикул/СтолбецНазвание/СтолбецЦена/СтолбецЕдиницы/СтолбецКатегория/СтолбецГруппа/СтолбецПодгруппа
        Else 'выбран диапазон
            oExcelApp.WindowState = xlMinimized
            MsgBox "Был выбран диапазон ячеек!" & vbCrLf & vbCrLf & "Необходимо выбрать одну ячейку", vbExclamation + vbOKOnly, "САПР-АСУ: Предупреждение"
            i = i - 1
            oExcelApp.WindowState = xlMaximized
        End If
    Next

    wbExcel.Close SaveChanges:=False
    oExcelApp.Application.Quit
    
    Set oExcelApp = CreateObject("Excel.Application")
    Set wbExcel = oExcelApp.Workbooks.Open(sPath & DBNameIzbrannoeExcel)
    wbExcel.Worksheets(ExcelNastrojkiPrajsov).Activate
    lLastRow = oExcelApp.Sheets(ExcelNastrojkiPrajsov).Cells(oExcelApp.Sheets(ExcelNastrojkiPrajsov).Rows.Count, 1).End(xlUp).Row
    For i = 1 To 12
        wbExcel.Sheets(ExcelNastrojkiPrajsov).Cells(lLastRow + 1, i) = mVendorData(i - 1)
    Next
    oExcelApp.Visible = True
    wbExcel.Save

End Sub


Public Sub FillExcel_cmbxProizvoditel(DBName As String, SQLQuery As String, cmbx As ComboBox, Optional ByVal Price As Boolean = False)
'------------------------------------------------------------------------------------------------------------
' Macros        : FillExcel_cmbxProizvoditel - Заполняет ComboBox Производители из Excel как базы данных САПР-АСУ
'------------------------------------------------------------------------------------------------------------
    Dim oExcelApp As Excel.Application
    Dim wbExcel As Excel.Workbook
    Dim sPath As String
    Dim UserRange As Excel.Range
    Dim i As Integer
    Dim j As Integer
    Dim mProizvoditel() As classProizvoditelBD

    Set oExcelApp = CreateObject("Excel.Application")
    sPath = Visio.ActiveDocument.path
    Set wbExcel = oExcelApp.Workbooks.Open(sPath & DBNameIzbrannoeExcel)
    
    lLastRow = oExcelApp.Sheets(ExcelNastrojkiPrajsov).Cells(oExcelApp.Sheets(ExcelNastrojkiPrajsov).Rows.Count, 1).End(xlUp).Row
    Set UserRange = oExcelApp.Worksheets(ExcelNastrojkiPrajsov).Range("A2:L" & lLastRow)
    
    ReDim mProizvoditel(lLastRow - 2)
    For i = 1 To lLastRow - 1
            Set mProizvoditel(i - 1) = New classProizvoditelBD
            mProizvoditel(i - 1).NameVendor = UserRange.Cells(i, 1)
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
    
    j = 0
    For i = 0 To UBound(mProizvoditel)
        If mProizvoditel(i).FileName = "" And Price Then
            'пропускаем производителя, если у него нету файла
        Else
            cmbx.AddItem mProizvoditel(i).NameVendor
            cmbx.List(j, 1) = "" & i + 1
            j = j + 1
        End If
    Next
    
    wbExcel.Close SaveChanges:=False
    oExcelApp.Application.Quit
    
End Sub