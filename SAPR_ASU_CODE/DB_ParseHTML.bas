'------------------------------------------------------------------------------------------------------------
' Module        : DB_ParseHTML - Парсинг сайтов магазинов для получения цен и картинок товара
' Author        : gtfox
' Date          : 2021.02.22
' Description   : Разбор разметки HTML сайтов магазинов.
'               : Не используется, т.к. отсутсвует имитация поведения человека и сайты банят эти парсеры + меняется разметка сайтов и надо переписывать алгоритмы.
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

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
    Dim wb As Excel.Workbook
    Dim strTempXls As String
    
    strTempXls = ThisDocument.path & "temp.xls"
    Set oExcel = CreateObject("Excel.Application")
    If Dir(strTempXls, 16) = "" Then
        Set wb = oExcel.Workbooks.Add
        wb.SaveAs FileName:=strTempXls
    Else
        Set wb = oExcel.Workbooks.Open(strTempXls)
    End If
    wb.Activate
    Set pic = wb.ActiveSheet.Pictures.Insert(ImgPNG)
    pic.Width = 300
    pic.Height = 300
    pic.Copy

    With wb.Worksheets(1).ChartObjects.Add(0, 0, pic.Width, pic.Height).Chart
        .Paste
        .Export Left(ImgPNG, Len(ImgPNG) - 3) & "jpg", "jpg"
    End With
    
    wb.Close savechanges:=False
    oExcel.Application.Quit
    
    Kill ImgPNG
    Kill strTempXls
    
    ConvertToJPG = Left(ImgPNG, Len(ImgPNG) - 3) & "jpg"

End Function