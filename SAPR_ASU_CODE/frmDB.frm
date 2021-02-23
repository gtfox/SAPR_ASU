'Option Explicit
'------------------------------------------------------------------------------------------------------------
' Module        : frmDB - Форма поиска и задания данных для элемента схемы их БД
' Author        : gtfox
' Date          : 2021.02.22
' Description   : Выбор данных из БД прайс листа, фильтрация по категориям и полнотекстовый поиск
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
#End If
Private Const LVM_FIRST As Long = &H1000   ' 4096
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)   ' 4126
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Dim glShape As Visio.Shape 'шейп из модуля DB

Dim colShapes As Collection
Dim colPages As Collection
Dim FindType As Integer 'Кто запустил создание связи (родитль/дочерний)
Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи
Dim HyperLinkToParentPLC As String
Dim HyperLinkToParentPLCMod As String
Dim mstrAdrParentPLC() As String
Dim mstrAdrParentPLCMod() As String
Dim shpParentPLC As Visio.Shape 'Родительский плк с модулями
Dim shpParentPLCMod As Visio.Shape 'Родительский модуль со входами внутри  родительского плк
Dim vsoShp As Visio.Shape
Dim bError As Boolean

Private Sub UserForm_Initialize() ' инициализация формы
    Set colShapes = New Collection
    Set colPages = New Collection

    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи
    lstvTable.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    
    lstvTable.ColumnHeaders.Add , , "Артикул" ' добавить ColumnHeaders
    lstvTable.ColumnHeaders.Add , , "Название" ' SubItems(1)
    lstvTable.ColumnHeaders.Add , , "Цена" ' SubItems(2)
    lstvTable.ColumnHeaders.Add , , "Макс./Сущ.вх." ' SubItems(3)

    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
    
    tbtnFiltr.Caption = ChrW(9650)
    tbtnBD = True
    
    
'    cmbxProizvoditel.ColumnCount = 3 'Показ столбцов
'    cmbxProizvoditel.AddItem "111"
'    cmbxProizvoditel.Column(1, 0) = "222"
'    cmbxProizvoditel.Column(2, 0) = "+222"
'    cmbxProizvoditel.AddItem "333"
'    cmbxProizvoditel.Column(1, 1) = "444"
'    cmbxProizvoditel.Column(2, 1) = "+444"


    cmbxProizvoditel.ColumnCount = 1 'Столбцы скрыты
    cmbxProizvoditel.AddItem "111"
    cmbxProizvoditel.List(0, 1) = "222"
    cmbxProizvoditel.List(0, 2) = "+222"
    cmbxProizvoditel.AddItem "333"
    cmbxProizvoditel.List(1, 1) = "444"
    cmbxProizvoditel.List(1, 2) = "+444"
    
'    ggg = cmbxProizvoditel.List(1, 1)

End Sub

Sub Run(vsoShape As Visio.Shape) 'Приняли шейп из модуля DB
    Set glShape = vsoShape 'И определили его как глолбальный в форме frmDB

    Fill_lstvTable

    Call lblHeaders_Click ' выровнять ширину столбцов по заголовкам

    ReSize
    
    If bError Then
        Unload Me
    Else
        frmDB.Show
    End If
    
End Sub



Private Sub ReSize() ' изменение формы. Зависит от длины в lstvTable
    Dim lstvTableWidth As Single
    
    lblHeaders_Click
    
    If lstvTable.ListItems(1).Width > 381 Then
        lstvTableWidth = lstvTable.ListItems(1).Width
    Else
        lstvTableWidth = 381
    End If
    
    lstvTable.Width = lstvTableWidth + 20
    frameTab.Width = lstvTable.Width + 10
    
    frameFilters.Width = frameTab.Width
    Me.Width = frameTab.Width + 14
    cmbxKategoriya.Width = frameFilters.Width - cmbxKategoriya.Left - 6
    cmbxGruppa.Width = frameFilters.Width - cmbxGruppa.Left - 6
    cmbxPodgruppa.Width = frameFilters.Width - cmbxPodgruppa.Left - 6
    btnClose.Left = Me.Width - btnClose.Width - 10
    tbtnFiltr.Left = Me.Width - tbtnFiltr.Width - 10
    btnFavAdd.Left = btnClose.Left - btnFavAdd.Width - 10
    btnETM.Left = btnFavAdd.Left - btnETM.Width - 2
    frameProizvoditel.Width = btnETM.Left - frameProizvoditel.Left - 6
    cmbxProizvoditel.Width = frameProizvoditel.Width - 12
    'lblResult.Top = Me.Height - 35
    lblResult.Left = frameTab.Width - lblResult.Width
    btnFind.Left = frameTab.Width - btnFind.Width - 6
    frameNazvanie.Width = btnFind.Left - frameNazvanie.Left - 6
    txtNazvanie1.Width = (frameNazvanie.Width - 16) / 2
    txtNazvanie2.Left = txtNazvanie1.Left + txtNazvanie1.Width
    txtNazvanie2.Width = frameNazvanie.Width / 4
    txtNazvanie3.Left = txtNazvanie2.Left + txtNazvanie2.Width
    txtNazvanie3.Width = frameNazvanie.Width / 4
    
    
End Sub

Private Sub cmbxProizvoditel_Change()

End Sub

Private Sub tbtnBD_Click()
    tbtnFav = Not tbtnBD
    
End Sub

Private Sub tbtnFav_Click()
    tbtnBD = Not tbtnFav
End Sub

Private Sub lstvTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку

    With lstvTable
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
    
End Sub

Sub Fill_lstvTable() ' заполнение списка родительских элементов схемы
    Dim i, j, x, y, n, k As Integer
    Dim itmx As ListItem
    Dim wires As String
    Dim vsoShape As Visio.Shape
    lstvTable.ListItems.Clear
    
    Dim dbsDatabase As DAO.Database
    Dim rstRecordset As DAO.Recordset
    Dim strPath As String
    Dim SQLQuery As String
    Dim List As String
    
    'Определяем запрос SQL для отбора записей из базы данных
    SQLQuery = "SELECT КодПозиции, Артикул, Название, Цена " & _
                "FROM Прайс;"

    'Создаем набор записей для получения списка
    strPath = ThisDocument.path & "SAPR_ASU_IEK.accdb" 'Schneider IEK ABB
    Set dbsDatabase = GetDBEngine.OpenDatabase(strPath)
    Set rstRecordset = dbsDatabase.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
    
    '---Ищем необходимую запись в наборе данных и по ней создаем набор значений для списка для заданных параметров
    'With rst

    rstRecordset.MoveFirst
    Do Until rstRecordset.EOF
'        Set itmx = lstvTable.ListItems.Add(, """" & rstRecordset![КодПозиции] & """", rstRecordset![Артикул])
        Set itmx = lstvTable.ListItems.Add(, """" & rstRecordset![КодПозиции] & """", rstRecordset![КодПозиции])
        itmx.SubItems(1) = rstRecordset![Артикул]
        itmx.SubItems(2) = rstRecordset![Название]
        itmx.SubItems(3) = rstRecordset![Цена]
        
        rstRecordset.MoveNext
        i = i + 1
    Loop
    'End With
    lblResult.Caption = "Найдено записей: " & i 'rstRecordset.RecordCount
'
'
'
'
'
'    Select Case FindType
'        Case typePLCModChild  'Если макрос активировался дочерним PLCModChild - значит искали родителей PLCModParent
'            For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
'                With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
'                    Set itmx = lstvTable.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), .Cells("User.Name").ResultStr(0)) '.Cells("TheText").ResultStr("")
'                    itmx.SubItems(1) = .Cells("Prop.Model").ResultStr(0)
'                    'подсчет кол-ва связей модуля
'                    k = 0
'                    For n = 1 To .Section(visSectionHyperlink).Count
'                        k = k + IIf(.CellsU("Hyperlink." & n & ".SubAddress").ResultStr(0) = "", 0, 1)
'                    Next
'                    itmx.SubItems(2) = k
'                    itmx.SubItems(3) = .Cells("Prop.NIO").Result(0) & "  |  " & .Shapes.Count - 1
'                    x = 0
'                    y = 0
'                    For Each vsoShape In .Shapes
'                        If (vsoShape.Name Like "PLCIOL*") Or (vsoShape.Name Like "PLCIOR*") Then
'                            'подсчет кол-ва связанных входов
'                            x = x + IIf(vsoShape.CellsU("Hyperlink.IO.SubAddress").ResultStr(0) <> "", 1, 0)
'                            'подсчет кол-ва подключенных входов
'                            For n = 1 To 4
'                                If vsoShape.Cells("User.w" & n).Result(0) <> 0 Then
'                                    y = y + 1
'                                    Exit For
'                                End If
'                            Next
'                        End If
'                    Next
'                    itmx.SubItems(4) = x & "  |  " & y
'
'              End With
'            Next i
'        Case typePLCIOChild 'Если макрос активировался дочерним PLCIO - значит искали PLCIO
'            For i = 1 To colShapes.Count  ' добавить N ListItem в коллекцию ListItems
'                With ActiveDocument.Pages.ItemFromID(colPages.Item(i)).Shapes.ItemFromID(colShapes.Item(i))
'                    Set itmx = lstvTable.ListItems.Add(, colPages.Item(i) & "/" & colShapes.Item(i), .Cells("User.Name").ResultStr(0)) '.Cells("TheText").ResultStr("")
'                    itmx.SubItems(1) = .CellsU("Hyperlink.IO.ExtraInfo").ResultStr(0)
'                    wires = IIf(.Cells("User.w1").Result(0) <> 0, .Cells("User.w1").Result(0), "")
'                    For j = 2 To 4
'                        wires = IIf(.Cells("User.w" & j).Result(0) <> 0, wires & ", " & .Cells("User.w" & j).Result(0), wires & "")
'                    Next j
'                    itmx.SubItems(2) = wires
'                    wires = ""
'                End With
'            Next i
'    End Select

End Sub

Private Sub lstvTable_DblClick()

    Select Case FindType
        Case typePLCModChild  'Если макрос активировался дочерним - значит искали родителей
            'Создаем связь как и было задумано
            AddReferencePLCMod glShape, shpParent
        Case typePLCIOChild 'Если макрос активировался родителем - значит искали дочерних
            'Меняем местами родителя/дочернего, т.к. в переменной glShape содержится родитель, а в shpParent дочерний
            AddReferencePLCIO glShape, shpParent
    End Select

    'Активация событий. Они чета сомодезактивируются xD
    'Set vsoPagesEvent = ActiveDocument.Pages
    
    btnClose_Click
    
End Sub

Private Sub lstvTable_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim vsoShape As Visio.Shape
    Dim ShapeID As String
    Dim PageID As String
    Dim mstrShPgID() As String
    
    'lblCurParent.Caption = Item.Text
    
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
        'lblCurParent.Caption = Item.Text + "  " + vsoShape.Cells("User.Location").ResultStr(0)
    End If
    

    ReSize
    
    Set shpParent = vsoShape 'передаем родителя для создания связи
    
End Sub

Private Sub lblContent_Click() ' выровнять ширину столбцов по содержимому
   Dim colNum As Long
   For colNum = 0 To lstvTable.ColumnHeaders.Count - 1
      Call SendMessage(lstvTable.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvTable.ColumnHeaders.Count - 1
      Call SendMessage(lstvTable.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

Private Sub tbtnFiltr_Click()
    If tbtnFiltr.Value Then
        frameFilters.Height = 84
        tbtnFiltr.Caption = ChrW(9650) 'вверх
    Else
        frameFilters.Height = 0
        tbtnFiltr.Caption = ChrW(9660) 'вниз
    End If
    frameTab.Top = frameFilters.Top + frameFilters.Height
    Me.Height = frameTab.Top + frameTab.Height + 36
    lblResult.Top = Me.Height - 35
End Sub



 Sub btnClose_Click() ' выгрузка формы

    With ActiveWindow
        .Page = glShape.ContainingPage
        .Select glShape, visDeselectAll + visSubSelect     ' выделение шейпа
        .SetViewRect pinLeft, pinTop, pinWidth, pinHeight  'Восстановление вида окна после закрытия формы
                    '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    Unload Me
    
End Sub