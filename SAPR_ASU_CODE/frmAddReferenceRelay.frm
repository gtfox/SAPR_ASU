


'------------------------------------------------------------------------------------------------------------
' Module        : frmAddReferenceRelay - Форма создания связей (перекрестных ссылок) элементов схемы
' Author        : gtfox на основе Shishok::Form_Find
' Date          : 2020.05.19
' Description   : Дочерний элемент выбирает себе родителя через форму
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе этого:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : Form_Find поиск и выделение
                ' Author    : Shishok
                ' Date      : 11.06.2018
                ' Purpose   : Поиск и выделение шейпов по критерию(текст). Для Windows 7 x 32 или типа того
                ' Links     : https://github.com/Shishok/, https://yadi.sk/d/qbpj9WI9d2eqF
                '------------------------------------------------------------------------------------------------------------

'Option Explicit
'Option Base 1

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#End If
Private Const LVM_FIRST As Long = &H1000   ' 4096
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)   ' 4126
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Dim shpChild As Visio.Shape 'шейп из модуля CrossReference
Dim shpParent As Visio.Shape 'шейп выбанный в форме lstvParent. нужен для создания связи
Dim colShapes As Collection
Dim colPages As Collection
Dim FindType As Integer 'Кто запустил создание связи (родитль/дочерний)
Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи

Sub run(vsoShape As Visio.Shape) 'Приняли шейп из модуля CrossReference
    Set shpChild = vsoShape 'И определили его в форме frmAddReference
    
    FindType = ShapeSAType(shpChild)
    
    Fill_lstvPages
    
    Fill_ShapeCollection ActivePage
    
    Select Case FindType
        Case typeNO, typeNC 'Если макрос активировался дочерним - значит искали родителей
            lstvParent.ColumnHeaders.Add , , "Элементы" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Item(1).Width = lstvParent.Width - 18
        Case typeCoil, typeParent 'Если макрос активировался родителем - значит искали дочерних
            lstvParent.ColumnHeaders.Add , , "Контакт" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Add , , "Связь" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Add , , "Адрес" ' добавить ColumnHeaders
            lstvParent.ColumnHeaders.Add , , "Страница" ' добавить ColumnHeaders
            lstvChild.Visible = False
            lstvParent.Width = 170
            lblResult.Left = 200
            btnClose.Left = 200
            Me.Width = 286
    End Select
    
    Fill_lstvParent

    Call lblHeaders_Click

    lblResult.Caption = "Найдено фигур: " & colShapes.Count
    
    ReSize

    frmAddReferenceRelay.Show
End Sub

Sub Fill_ShapeCollection(vsoPage As Visio.Page) 'Заполняем список с родительскими элементами
    'Dim vsoPage As Visio.Page
    Dim vsoShape As Visio.Shape
    
    If chkAllPages Then
        For Each vsoPage In ActiveDocument.Pages ' перебор страниц документа и шейпов
            For Each vsoShape In vsoPage.Shapes
                SelectType vsoShape, vsoPage
            Next
        Next
    Else
        For Each vsoShape In vsoPage.Shapes ' перебор шейпов на выбранной странице
            SelectType vsoShape, vsoPage
        Next
    End If
    
End Sub

Private Sub SelectType(vsoShape As Visio.Shape, vsoPage As Visio.Page) ' Выбор по типу

    If vsoShape.CellExistsU("User.SAType", 0) Then 'отсеиваем посторонние шейпы не имеющие поле ТИП
        Select Case FindType 'Определяемся в соответствии с типом вызвавшего макрос шейпа
            Case typeNO, typeNC 'Если макрос активировался дочерним - значит искали родителей
                Select Case ShapeSAType(vsoShape)
                    Case typeCoil, typeParent

                        SelectText vsoShape, vsoPage
                End Select
            Case typeCoil, typeParent 'Если макрос активировался родителем - значит искали дочерних
                Select Case ShapeSAType(vsoShape)
                    Case typeNO, typeNC

                        SelectText vsoShape, vsoPage
                End Select
        End Select
    End If
   
End Sub

Sub SelectText(vsoShape As Visio.Shape, vsoPage As Visio.Page) ' Выбор - по тексту
    Dim shtxt As String, txt As String
    
    shtxt = Switch(chkCase = True, vsoShape.Characters.Text, chkCase = False, LCase(vsoShape.Characters.Text))
    txt = Switch(chkCase = True, txtShapeText.Text, chkCase = False, LCase(txtShapeText.Text))
    
    If shtxt Like txt Then ' проверка текста шейпа
        Call AddToCol(vsoShape, vsoPage)
    End If
    
End Sub

Private Sub AddToCol(vsoShape As Visio.Shape, vsoPage As Visio.Page)  ' добавление элементов в коллекции
    On Error GoTo ExitLine
        colShapes.Add vsoShape.ID ' коллекция ID шейпов
        colPages.Add vsoPage.ID ' коллекция ID страниц
ExitLine:
End Sub


Private Sub btnFindAll_Click() ' процедура поиска по кнопке

    FindShapes
    
End Sub

Private Sub FindShapes() ' процедура поиска
    Set colShapes = New Collection
    Set colPages = New Collection

    Fill_ShapeCollection ActiveDocument.Pages(lblCurPage.Caption)
    
    If chkAllPages.Value Then
        lblCurPageALL.Visible = True
        lblCurPage.Visible = False
    Else
        lblCurPageALL.Visible = False
        lblCurPage.Visible = True
    End If
    
    If colShapes.Count > 0 Then
        Fill_lstvParent
    Else
        lstvParent.ListItems.Clear
        lstvChild.ListItems.Clear
    End If

    lblResult.Caption = "Найдено фигур: " & colShapes.Count
    
    ReSize
    
    Call lblHeaders_Click
    
End Sub



Private Sub ReSize() ' изменение высоты формы. Зависит от количества элементов в listbox
    Dim H As Single
    
    H = lstvPages.ListItems.Count
    If H < lstvParent.ListItems.Count Then H = lstvParent.ListItems.Count
    If H < lstvChild.ListItems.Count Then H = lstvChild.ListItems.Count
    
    H = H * 12 + 12
    If H < 48 Then H = 48
    If H > 328 Then H = 328
    
    Me.Height = lstvPages.Top + H + 26
    
    lstvPages.Height = H
    lstvParent.Height = H
    lstvChild.Height = H


    
End Sub

Private Sub chkAllPages_Click()

    FindShapes
    
End Sub

Private Sub lstvChild_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim vsoShape As Visio.Shape
    Dim ShapeID As String
    Dim PageID As String
    Dim mstrShPgID() As String

    mstrShPgID = Split(Item.Key, "/")
    PageID = mstrShPgID(0)   ' ID страницы
    ShapeID = mstrShPgID(1)   ' ID шейпа

    With ActiveWindow
        .Page = ActiveDocument.Pages.ItemFromID(PageID) ' активация нужной страницы
        Set vsoShape = ActivePage.Shapes.ItemFromID(ShapeID)
        .Select vsoShape, visDeselectAll + visSelect     ' выделение шейпа
        '.CenterViewOnShape ActivePage.Shapes(shName) , visCenterViewSelectShape '2010+
        .SetViewRect vsoShape.Cells("PinX") - pinWidth / 2, vsoShape.Cells("PinY") + pinHeight / 2, pinWidth, pinHeight
        '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    
End Sub

Private Sub lstvChild_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvChild
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
End Sub

Private Sub lstvPages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvPages
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
End Sub

Private Sub lstvParent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvParent
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
End Sub

Sub Fill_lstvChild(vsoShape As Visio.Shape) ' заполнение списка дочерних элементов схемы (контактов)
    Dim i As Integer
    Dim itmx As ListItem
    Dim mstrAdrChild() As String
    Dim shpInfoChild As Visio.Shape
    
    lstvChild.ListItems.Clear
    
    If vsoShape.CellExistsU("Scratch.A1", 0) Then
        For i = 1 To vsoShape.Section(visSectionScratch).Count
            If vsoShape.CellsU("Scratch.A" & i).ResultStr(0) <> "" Then
                'Разбиваем HyperLink на имя страницы и имя шейпа
                mstrAdrChild = Split(vsoShape.CellsU("Scratch.A" & i).ResultStr(0), "/")
                Set shpInfoChild = ActiveDocument.Pages.ItemU(mstrAdrChild(0)).Shapes(mstrAdrChild(1))
                Set itmx = lstvChild.ListItems.Add(, shpInfoChild.ContainingPage.ID & "/" & shpInfoChild.ID, _
                shpInfoChild.Characters.Text + " " + IIf(ShapeSATypeIs(shpInfoChild, typeNO), "NO", "NC") _
                + " " + shpInfoChild.CellsU("User.Location").ResultStr(0)) '
            End If
        Next
    End If
    
End Sub

Sub Fill_lstvParent() ' заполнение списка родительских элементов схемы
    Dim i As Integer
    Dim itmx As ListItem
    
    lstvParent.ListItems.Clear
    
    Select Case FindType
        Case typeNO, typeNC 'Если макрос активировался дочерним - значит искали родителей
            For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
                With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
                Set itmx = lstvParent.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), .Characters.Text) '.Cells("TheText").ResultStr("")
              End With
            Next i
        Case typeCoil, typeParent 'Если макрос активировался родителем - значит искали дочерних
            For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
                With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
                    Set itmx = lstvParent.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), .Characters.Text) '.Cells("TheText").ResultStr("")
                    itmx.SubItems(1) = IIf(.Cells("User.LocationParent").ResultStr(0) = "0,0000", "", .Cells("User.LocationParent").ResultStr(0))
                    itmx.SubItems(2) = .Cells("User.Location").ResultStr(0)
                    itmx.SubItems(3) = .ContainingPage.Name
                End With
            Next i
    End Select

End Sub

Private Sub Fill_lstvPages()   ' заполнение списка страниц
    Dim i As Integer
    Dim itmx As ListItem
    Dim vsoPage As Visio.Page
    
    lstvPages.ListItems.Clear
    
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.PageSheet.CellExistsU("Prop.SA_NazvanieShemy", 0) Then
            Set itmx = lstvPages.ListItems.Add(, vsoPage.ID & "/", vsoPage.Name)
        End If
    Next
    
End Sub

Private Sub lstvPages_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    chkAllPages.Value = False
    lblCurPage.Caption = Item.Text
    lblCurPage.Visible = True
    lblCurPageALL.Visible = False
    
    FindShapes
    
End Sub

Private Sub lstvParent_DblClick()

    Select Case FindType
        Case typeNO, typeNC 'Если макрос активировался дочерним - значит искали родителей
            'Создаем связь как и было задумано
            AddReferenceRelay shpChild, shpParent
        Case typeCoil, typeParent 'Если макрос активировался родителем - значит искали дочерних
            'Меняем местами родителя/дочернего, т.к. в переменной shpChild содержится родитель, а в shpParent дочерний
            AddReferenceRelay shpParent, shpChild
    End Select

    'Активация событий. Они чета сомодезактивируются xD
'    Set vsoPagesEvent = ActiveDocument.Pages
    
    btnClose_Click
    
End Sub

Private Sub lstvParent_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim vsoShape As Visio.Shape
    Dim ShapeID As String
    Dim PageID As String
    Dim mstrShPgID() As String
    
    lblCurParent.Caption = Item.Text
    
    mstrShPgID = Split(Item.Key, "/")
    PageID = mstrShPgID(0)   ' ID страницы
    ShapeID = mstrShPgID(1)   ' ID шейпа

    With ActiveWindow
        .Page = ActiveDocument.Pages.ItemFromID(PageID) ' активация нужной страницы
        Set vsoShape = ActivePage.Shapes.ItemFromID(ShapeID)
        If vsoShape.Parent.Type = visTypeGroup Then
            .Select vsoShape, visDeselectAll + visSubSelect  ' выделение субшейпа
            '.CenterViewOnShape ActivePage.Shapes(shName), visCenterViewSelectShape '2010+
        Else
            .Select vsoShape, visDeselectAll + visSelect     ' выделение шейпа
            '.CenterViewOnShape ActivePage.Shapes(shName) , visCenterViewSelectShape '2010+
            .SetViewRect vsoShape.Cells("PinX") - pinWidth / 2, vsoShape.Cells("PinY") + pinHeight / 2, pinWidth, pinHeight
            '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
        End If
    End With

    If vsoShape.CellExistsU("User.Location", 0) Then
        lblCurParent.Caption = Item.Text + "  " + vsoShape.Cells("User.Location").ResultStr(0)
    End If
    
    Select Case FindType
        Case typeNO, typeNC 'Если макрос активировался дочерним - значит искали родителей
            Fill_lstvChild vsoShape 'Заполняем лист контактов
        Case typeCoil, typeParent 'Если макрос активировался родителем - значит искали дочерних
            'ниче не делаем
    End Select
    
    ReSize
    
    Set shpParent = vsoShape 'передаем родителя для создания связи
    
End Sub

Private Sub lblContent_Click() ' выровнять ширину столбцов по содержимому
   Dim colNum As Long
   For colNum = 0 To lstvParent.ColumnHeaders.Count - 1
      Call SendMessage(lstvParent.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvParent.ColumnHeaders.Count - 1
      Call SendMessage(lstvParent.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

Private Sub UserForm_Initialize() ' инициализация формы
    Set colShapes = New Collection
    Set colPages = New Collection
    
    Me.Height = 213 ' высота формы
    
    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи
    
    txtShapeText.Text = "*" ' вставка текста в поле поиска
    lblCurParent.Caption = ""
    lblCurPageALL.Caption = "Все страницы"
    lblCurPage.Caption = ActivePage.Name
    chkAllPages.Value = False
    
    lstvPages.ColumnHeaders.Add , , "Страницы" ' добавить ColumnHeaders
    'Call SendMessage(lstvPages.hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE_USEHEADER) ' выровнять ширину столбцов по заголовкам
    'Call SendMessage(lstvPages.hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE) ' выровнять ширину столбцов по содержимому
    lstvPages.ColumnHeaders.Item(1).Width = lstvPages.Width - 18
 
    lstvChild.ColumnHeaders.Add , , "Контакты" ' добавить ColumnHeaders
    'Call SendMessage(lstvChild.hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE_USEHEADER)  ' выровнять ширину столбцов по заголовкам
    'Call SendMessage(lstvChild.hWnd, LVM_SETCOLUMNWIDTH, 0, ByVal LVSCW_AUTOSIZE) ' выровнять ширину столбцов по содержимому
    lstvChild.ColumnHeaders.Item(1).Width = lstvParent.Width - 4
    
    lstvPages.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvParent.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvChild.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке


End Sub

Private Sub btnClose_Click() ' выгрузка формы

    With ActiveWindow
        .Page = shpChild.ContainingPage
        .Select shpChild, visDeselectAll + visSelect     ' выделение шейпа
        .SetViewRect pinLeft, pinTop, pinWidth, pinHeight  'Восстановление вида окна после закрытия формы
                    '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    Unload Me
    
End Sub