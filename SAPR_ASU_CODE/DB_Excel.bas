'------------------------------------------------------------------------------------------------------------
' Module        : DB_Excel - База данных прайс листов и избранного на основе Excel
' Author        : gtfox
' Date          : 2023.01.30
' Description   : База данных прайс листов, избранного и их обеспечение на основе Excel
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit

Public Const DBNameIzbrannoeExcel As String = "SAPR_ASU_Izbrannoe.xls" 'Имя файла избронного

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


End Sub

Private Sub xls_query1()
'------------------------------------------------------------------------------------------------------------
' Macros        : xls_query - Заполняет массив данными из Excel
'------------------------------------------------------------------------------------------------------------
Dim strRange As String
strRange = "A2:H"
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
    Dim ttt() As String
    Dim mDialogString() As String
    Dim sDialogString As String
    
    sDialogString = "Выберите начальную ячейку данных прайса;" & _
                    "Выберите конечную ячейку данных прайса;" & _
                    "Выберите ячейку в столбце ""Артикул"";" & _
                    "Выберите ячейку в столбце ""Название"";" & _
                    "Выберите ячейку в столбце ""Цена"";" & _
                    "Выберите ячейку в столбце ""Категория"";" & _
                    "Выберите ячейку в столбце ""Группа"";" & _
                    "Выберите ячейку в столбце ""Подгруппа"";" & _
                    "Выберите ячейку в столбце ""Единица"""
    mDialogString = Split(sDialogString, ";")
    
    Set oExcel = CreateObject("Excel.Application")
    pth = Visio.ActiveDocument.path
'    oExcel.Visible = True ' для наглядности
'    oExcel.WindowState = xlMinimized
    
    Set fd = oExcel.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = False
        .InitialFileName = pth
        Set ffs = .Filters
        With ffs
            .Clear
            .Add "Excel", "*.xls"
            .Add "Excel", "*.xlsx"
        End With
        Chois = oExcel.FileDialog(msoFileDialogOpen).Show
    End With
    If Chois = 0 Then oExcel.Application.Quit: frmClose = True: Exit Sub
    sFileName = oExcel.FileDialog(msoFileDialogOpen).SelectedItems(1)
    
    If InStr(sFileName, pth) = 1 Then 'файл в той же папке, что и проект
        sRelativeFileName = Replace(sFileName, pth, "") 'относительный путь
    Else
        sFileName = sFileName 'абсолютный путь
    End If
    

    sPath = pth
'    sFileName = "SP_2_Visio.xls"
    sFile = sFileName
    
'    If Dir(sFile, 16) = "" Then 'есть хотя бы один файл
'        MsgBox "Файл " & sFileName & " не найден в папке: " & sPath, vbCritical, "Ошибка"
'        Exit Sub
'    End If
    
    Set sp = oExcel.Workbooks.Open(sFile)
    Load frmVyborListaExcel
    frmVyborListaExcel.Show 'присваиваем Excel_imya_lista
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
    
'    oExcel.WindowState = xlMaximized
    oExcel.Visible = True ' для наглядности
'    oExcel.WindowState = xlMinimized
    
    Set UserRange = oExcel.InputBox _
    (Prompt:="Выберите диапазон A3:Ix", _
    Title:="Выбор диапазона", _
    Type:=8)

    oExcel.WindowState = xlMinimized
    
    
    
    

    ttt = Split(UserRange.Address, ":")
    If UBound(ttt) = 0 Then 'выбрана 1 ячейка
        ttt = Split(UserRange.Address, "$") 'буква столбца
        BukvaStolbca = ttt(1)
    Else 'выбран диапазон
        rrr = Split(ttt(0), "$") 'Строка начало
        StrokaNachalo = rrr(2)
        BukvaStolbca = rrr(1)
        rrr = Split(ttt(1), "$") 'Строка конец
        StrokaKonec = rrr(2)
    End If
    
    ttt = Split(UserRange.Address, "$")
    
    rrr = ttt(1)
    

    sp.Close SaveChanges:=False
    oExcel.Application.Quit
    

End Sub

'    Set Total = UserRange
'        For Each tr In Total.Rows
'            RowCountXls = RowCountXls + 1
'            ColoumnCountXls = 0
'            For Each tc In Total.Rows.Columns
'                ColoumnCountXls = ColoumnCountXls + 1
'            Next tc
'        Next tr
'    ReDim arr(RowCountXls, ColoumnCountXls) As Variant
'    For qx = 1 To RowCountXls
'        For qy = 1 To ColoumnCountXls
'            arr(qx, qy) = Total.Cells(qx, qy) ' заполнение массива arr
'        Next qy
'    Next qx