'------------------------------------------------------------------------------------------------------------
' Module        : DB_Access - База данных прайс листов и избранного на основе Access
' Author        : gtfox
' Date          : 2021.02.22
' Description   : База данных прайс листов, избранного и их обеспечение на основе Access
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

'Option Explicit

Public Const frmMinWdth As Integer = 417 'Минимальна ширина формы
Public Const DBNameIzbrannoeAccess As String = "SAPR_ASU_Izbrannoe.accdb" 'Имя файла избронного

Public Const NaboryColor   As Long = &HBD0429 'синий

#If VBA7 Then
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

'Активация формы выбора элементов схемы из БД
Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
'    Load frmDBPriceAccess
'    frmDBPriceAccess.run vsoShape 'Передали его в форму
    Load frmDBPriceExcel
    frmDBPriceExcel.run vsoShape 'Передали его в форму
End Sub

'Function returns DBEngine for current Office Engine Type (DAO.DBEngine.60 or DAO.DBEngine.120)
Public Function GetDBEngine() As Object
    Dim engine As Object
        On Error GoTo EX
        Set GetDBEngine = DBEngine
    Exit Function
EX:
    Set GetDBEngine = CreateObject("DAO.DBEngine.120")
End Function

'Получаем Recordset по запросу
Public Function GetRecordSet(DBName As String, SQLQuery As String, Optional QueryDefName As String = "") As DAO.Recordset
    Dim dbs As DAO.Database
    Set dbs = GetDBEngine.OpenDatabase(ThisDocument.path & DBName)
    If QueryDefName <> "" Then
        On Error Resume Next
        dbs.QueryDefs.Delete QueryDefName
    End If
    Set GetRecordSet = dbs.CreateQueryDef(QueryDefName, SQLQuery).OpenRecordset(dbOpenDynaset)
    Set dbs = Nothing
End Function

'Выполняем SQL запрос
Public Sub ExecuteSQL(DBName As String, SQLQuery As String, Optional QueryDefName As String = "")
    Dim dbs As DAO.Database
    Set dbs = GetDBEngine.OpenDatabase(ThisDocument.path & DBName)
    dbs.Execute SQLQuery
    Set dbs = Nothing
End Sub

'Заполняет ComboBox запросами из БД
Public Sub Fill_ComboBox(DBName As String, SQLQuery As String, cmbx As ComboBox)
    Dim rst As DAO.Recordset
    Dim i As Integer
    Set rst = GetRecordSet(DBName, SQLQuery)
    cmbx.Clear
    cmbx.ColumnCount = 1
    i = 0
    With rst
    If .EOF Then Exit Sub
        .MoveFirst
        Do Until .EOF
            cmbx.AddItem .Fields(1).Value
            cmbx.List(i, 1) = "" & .Fields(0).Value
            i = i + 1
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub

'Заполняет ComboBox Производители запросами из БД
Public Sub Fill_cmbxProizvoditel(DBName As String, SQLQuery As String, cmbx As ComboBox, Optional ByVal Price As Boolean = False)
    Dim rst As DAO.Recordset
    Dim i As Integer
    Set rst = GetRecordSet(DBName, SQLQuery)
    cmbx.Clear
    cmbx.ColumnCount = 1
    i = 0
    With rst
    If .EOF Then Exit Sub
        .MoveFirst
        .MoveNext 'Пропускаем первый элемент
        Do Until .EOF
            If "" & .Fields(0).Value = "" And Price Then

            Else
                cmbx.AddItem .Fields(1).Value
                cmbx.List(i, 1) = "" & .Fields(0).Value
                cmbx.List(i, 2) = .Fields(2).Value
                i = i + 1
            End If
            .MoveNext
        Loop
    End With
    Set rst = Nothing
End Sub

'Заполняет lstvTable запросами из БД
Public Function Fill_lstvTable(DBName As String, SQLQuery As String, QueryDefName As String, lstvTable As ListView, Optional ByVal TableType As Integer = 0) As Double
    'TableType=1 - Избранное
    'TableType=2 - Набор
    Dim i As Double
    Dim iold As Double
    Dim j As Double
    Dim itmx As ListItem
    Dim rst As DAO.Recordset
    Dim RecordCount As Double

    Set rst = GetRecordSet(DBName, SQLQuery, QueryDefName)
    lstvTable.ListItems.Clear
    If rst.RecordCount > 0 Then
        rst.MoveLast
        RecordCount = rst.RecordCount
    '    frmDBPriceAccess.lblResult.Caption = "Найдено записей: " & RecordCount
    '    frmDBPriceAccess.ProgressBar.Visible = True
    '    frmDBPriceAccess.ProgressBar.Max = RecordCount
'        lstvTable.ListItems.Clear
        i = 0
        iold = 1000
        With rst
            If .EOF Then Exit Function
            .MoveFirst
            Do Until .EOF
                Set itmx = lstvTable.ListItems.Add(, """" & .Fields("КодПозиции").Value & "/" & .Fields("ПроизводительКод").Value & "/" & .Fields("ЕдиницыКод").Value & """", .Fields("Артикул").Value)
                itmx.SubItems(1) = .Fields("Название").Value
                itmx.SubItems(2) = .Fields("Цена").Value
                itmx.SubItems(3) = .Fields("Единица").Value
                If TableType = 1 Then
                    itmx.SubItems(4) = .Fields("Производитель").Value
                    itmx.SubItems(5) = "    "
                ElseIf TableType = 2 Then
                    itmx.SubItems(4) = .Fields("Производитель").Value
                    itmx.SubItems(5) = .Fields("Количество").Value
                    itmx.SubItems(6) = "    "
                End If
    
                'красим наборы
                If TableType = 1 Then  'and .Fields("Артикул").Value like "Набор_*" then
                    If .Fields("ПодгруппыКод").Value = 2 Then
                        itmx.ForeColor = NaboryColor
        '                    itmx.Bold = True
                        For j = 1 To itmx.ListSubItems.Count
        '                        itmx.ListSubItems(j).Bold = True
                            itmx.ListSubItems(j).ForeColor = NaboryColor
                        Next
                    End If
                End If
    '            i = i + 1
    '            If iold < i Then
    '                iold = iold + 1000
    '                frmDBPriceAccess.ProgressBar.Value = i
    '            End If
                .MoveNext
            Loop
        End With
        Fill_lstvTable = RecordCount 'i
    '    frmDBPriceAccess.ProgressBar.Visible = False
    End If
    Set rst = Nothing

End Function

'Заполняет lstvTableNabor запросами из БД
Public Function Fill_lstvTableNabor(DBName As String, IzbPozCod As String, lstvTable As ListView) As String
    Dim SQLQuery As String

    SQLQuery = "SELECT Наборы.КодПозиции, Наборы.ИзбрПозицииКод, Наборы.Артикул, Наборы.Название, Наборы.Цена, Наборы.Количество, Наборы.ПроизводительКод, Производители.Производитель, Наборы.ЕдиницыКод, Единицы.Единица " & _
                "FROM Единицы INNER JOIN (Производители INNER JOIN Наборы ON Производители.КодПроизводителя = Наборы.ПроизводительКод) ON Единицы.КодЕдиницы = Наборы.ЕдиницыКод " & _
                "WHERE Наборы.ИзбрПозицииКод=" & IzbPozCod & ";"
    Fill_lstvTableNabor = Fill_lstvTable(DBName, SQLQuery, "", lstvTable, 2)
End Function

'Считаем цену набора
Public Function CalcCenaNabora(lstvTable As ListView) As Double
    Dim i As Integer
    Dim Sum As Double
    For i = 1 To lstvTable.ListItems.Count
        Sum = Sum + CDbl(lstvTable.ListItems(i).SubItems(2)) * CInt(lstvTable.ListItems(i).SubItems(5))
    Next
    CalcCenaNabora = Sum
End Function

'Открываем форму с информацией о товаре из выбранного магазина
Public Sub MagazinInfo(Artikul As String, NomerMagazina As Integer)
    Dim mstrTempFile() As String
    Dim strTempFile As String
    Dim strImgURL As String
    Dim mstrTovar As Variant
    Dim link As String
    Dim linkCatalog As String

    If Artikul = "" Then Exit Sub
    
    Select Case NomerMagazina
        Case 0 'ЭТМ
            frmDBMagazinInfo.linkFind = "https://www.etm.ru/catalog/?searchValue=" & Artikul
            mstrTovar = ParseHTML_ETM(Artikul)
        Case 1 'АВС Электро
            frmDBMagazinInfo.linkFind = "https://avselectro.ru/search/index.php?q=" & Artikul
            mstrTovar = ParseHTML_AVS(Artikul)
        Case Else
            frmDBMagazinInfo.linkFind = "https://www.etm.ru/catalog/?searchValue=" & Artikul
            mstrTovar = ParseHTML_ETM(Artikul)
    End Select
    
    linkCatalog = mstrTovar(0)
    strImgURL = mstrTovar(4)
    mstrTempFile = Split(strImgURL, "/")
    If UBound(mstrTempFile) = -1 Then GoTo err
    strTempFile = ThisDocument.path & mstrTempFile(UBound(mstrTempFile))
    lngRC = URLDownloadToFile(0, strImgURL, strTempFile, 0, 0)
    If Right(strImgURL, 3) = "png" Then
        strTempFile = ConvertToJPG(strTempFile)
    End If
    On Error Resume Next
    frmDBMagazinInfo.imgKartinka.Picture = LoadPicture(strTempFile)
    Kill strTempFile
    frmDBMagazinInfo.lblNazvanie = mstrTovar(1)
    frmDBMagazinInfo.txtCena = mstrTovar(2)
    frmDBMagazinInfo.txtCenaRozn = mstrTovar(3)
    frmDBMagazinInfo.linkCatalog = mstrTovar(0)
err:
    frmDBMagazinInfo.run

End Sub

'Получаем товар со страницы поиска товара на сайте ETM.ru
Public Function ParseHTML_ETM(Artikul As String) As String()
    Dim HtmlFile As Object
    Dim Elemet As Object ', Elemet2 As Object
    Dim mstrTovar(4) As String
    Dim rUrl As String
    Dim done As Integer

    rUrl = "https://www.etm.ru/catalog/?searchValue=" & Artikul
    
    Set HtmlFile = CreateObject("HtmlFile")

    With HtmlFile
'        AddIntoTXTfile ThisDocument.path & "temp.html", GetHtml(rUrl)
        .Body.innerhtml = GetHtml(rUrl)
        For Each Elemet In .getElementsByTagName("a")
            If Elemet.ClassName = "nameofgood" Then
                mstrTovar(0) = Replace(Elemet.GetAttribute("href"), "about:", "https://www.etm.ru") 'Каталог
                mstrTovar(1) = Elemet.innertext 'Название
                Exit For
            End If
        Next
        done = 0
        For Each Elemet In .getElementsByTagName("div")
            Select Case Elemet.ClassName
                Case "catalog-col-right sale"
                    mstrTovar(2) = Elemet.getElementsByTagName("span")(0).innertext 'Цена
                    mstrTovar(3) = Elemet.getElementsByTagName("span")(3).innertext 'Цена розница
                    done = done + 1
                Case "catalog-col-img"
                    mstrTovar(4) = "https:" & Elemet.getElementsByTagName("img")(0).GetAttribute("data-originalSrc") 'Картинка
                    done = done + 1
                Case Else
            End Select
            If done = 2 Then Exit For
        Next
    End With
    ParseHTML_ETM = mstrTovar
End Function

'Получаем товар со страницы поиска товара на сайте avselectro.ru
Public Function ParseHTML_AVS(Artikul As String) As String()
    Dim HtmlFile As Object
    Dim Elemet As Object
    Dim mstrTovar(4) As String
    Dim rUrl As String

    rUrl = "https://avselectro.ru/search/index.php?q=" & Artikul
    
    Set HtmlFile = CreateObject("HtmlFile")

    With HtmlFile
'        AddIntoTXTfile ThisDocument.path & "temp.html", GetHtml(rUrl)
        .Body.innerhtml = GetHtml(rUrl)
        For Each Elemet In .getElementsByTagName("div")
            Select Case Elemet.ClassName
                Case "info__title"
                    mstrTovar(0) = Replace(Elemet.getElementsByTagName("a")(0).GetAttribute("href"), "about:", "https://avselectro.ru") 'Каталог
                    mstrTovar(1) = Elemet.getElementsByTagName("span")(0).innertext 'Название
                    Exit For
                Case Else
            End Select
        Next
        done = 0
        For Each Elemet In .getElementsByTagName("span")
            Select Case Elemet.ClassName
                Case "m-price"
                    mstrTovar(2) = Elemet.innertext 'Цена
                    done = done + 1
                Case "crossed-out"
                    mstrTovar(3) = Elemet.innertext 'Цена розница
                    done = done + 1
                Case Else
            End Select
            If done = 2 Then Exit For
        Next
        For Each Elemet In .getElementsByTagName("a")
            If Elemet.ClassName = "lightzoom" Then
                mstrTovar(4) = Replace(Elemet.GetAttribute("href"), "about:", "https://avselectro.ru") 'Картинка
                Exit For
            End If
        Next
    End With
    ParseHTML_AVS = mstrTovar
End Function

'Получает страницу сайта в строку
Public Function GetHtml(ByVal URL As String) As String
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", URL, False
        .send
        Do: DoEvents: Loop Until .ReadyState = 4
        GetHtml = .responsetext
    End With
End Function

'Конвертирует картринку PNG в JPG при помощи Excel
Public Function ConvertToJPG(ImgPNG As String) As String
    Dim pic As Object
    Dim oExcel As Excel.Application
    Dim WB As Excel.Workbook
    Dim strTempXls As String
    
    strTempXls = ThisDocument.path & "temp.xls"
    Set oExcel = CreateObject("Excel.Application")
    If Dir(strTempXls, 16) = "" Then
        Set WB = oExcel.Workbooks.Add
        WB.SaveAs filename:=strTempXls
    Else
        Set WB = oExcel.Workbooks.Open(strTempXls)
    End If
    WB.Activate
    Set pic = WB.ActiveSheet.Pictures.Insert(ImgPNG)
    pic.Width = 300
    pic.Height = 300
    pic.Copy

    With WB.Worksheets(1).ChartObjects.Add(0, 0, pic.Width, pic.Height).Chart
        .Paste
        .Export Left(ImgPNG, Len(ImgPNG) - 3) & "jpg", "jpg"
    End With
    
    WB.Close SaveChanges:=False
    oExcel.Application.Quit
    
    Kill ImgPNG
    Kill strTempXls
    
    ConvertToJPG = Left(ImgPNG, Len(ImgPNG) - 3) & "jpg"

End Function