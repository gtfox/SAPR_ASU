Option Explicit
'------------------------------------------------------------------------------------------------------------
' Module        : frmMenuUnLockSALayer - Форма разблокировки заблокированных шейпов
' Author        : gtfox
' Date          : 2022.02.16
' Description   : Удаление шейпов с заблокированного слоя, болкировка/разблокировка выбранного слоя
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
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

Dim colShapes As Collection
Public pinLeft As Double, pinTop As Double, pinWidth As Double, pinHeight As Double 'Для сохранения вида окна перед созданием связи
Dim vsoShp As Visio.Shape

Private Sub UserForm_Initialize() ' инициализация формы
    Set colShapes = New Collection

    ActiveWindow.GetViewRect pinLeft, pinTop, pinWidth, pinHeight   'Сохраняем вид окна перед созданием связи

    lstvShapes.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvShapes.ColumnHeaders.Add , , "Шейп"
    lstvShapes.ColumnHeaders.Add , , "Название"
    Application.ActiveWindow.Page.Layers.Add "SA_LockedLayer"
    Fill_cmbxLayers
    cmbxLayers.style = fmStyleDropDownList
    If ActivePage.Layers.Count > 0 Then
        cmbxLayers.ListIndex = ActivePage.Layers("SA_LockedLayer").Index - 1
    End If
End Sub

Private Sub cbDeleteFromLayer_Click()
    If vsoShp Is Nothing Then
        Exit Sub
    Else
        ActivePage.Layers(cmbxLayers.text).Remove vsoShp, 0
        Fill_lstvShapes
        If lstvShapes.ListItems.Count > 0 Then
            lstvShapes.ListItems(1).Selected = True
            lstvShapes.SetFocus
            lstvShapes_ItemClick lstvShapes.ListItems(1)
        End If
    End If
End Sub

Private Sub cbLockUnlockLayer_Click()
    If cmbxLayers.text <> "" Then
        If ActivePage.Layers(cmbxLayers.text).CellsC(visLayerLock).Result(0) = 1 Then
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerLock).FormulaU = "0"
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerColor).FormulaU = "255"
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerSnap).FormulaU = "1"
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerGlue).FormulaU = "1"
            cbLockUnlockLayer.Caption = "Заблокировать слой"
        Else
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerLock).FormulaU = "1"
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerColor).FormulaU = "19"
            ActivePage.Layers(cmbxLayers.text).CellsC(visLayerSnap).FormulaU = "0"
            If cmbxLayers.text = "SA_LockedWire" Then
                ActivePage.Layers(cmbxLayers.text).CellsC(visLayerGlue).FormulaU = "1"
            Else
                ActivePage.Layers(cmbxLayers.text).CellsC(visLayerGlue).FormulaU = "0"
            End If
            cbLockUnlockLayer.Caption = "Разблокировать слой"
        End If
    End If
End Sub

Private Sub cmbxLayers_Change()
    If ActivePage.Layers(cmbxLayers.text).CellsC(visLayerLock).Result(0) = 1 Then cbLockUnlockLayer.Caption = "Разблокировать слой" Else cbLockUnlockLayer.Caption = "Заблокировать слой"
    Fill_lstvShapes
End Sub

Sub Fill_cmbxLayers()
    Dim vsoLayer As Visio.Layer
    cmbxLayers.Clear
    For Each vsoLayer In ActivePage.Layers
        cmbxLayers.AddItem vsoLayer.name, vsoLayer.Index - 1
    Next
End Sub

Sub Fill_lstvShapes()
    Dim itmx As ListItem
    Dim i As Integer
    FillCollection ActivePage.Layers(cmbxLayers.text)
    lstvShapes.ListItems.Clear
    For i = 1 To colShapes.Count
        Set itmx = lstvShapes.ListItems.Add(, colShapes.Item(i).name, colShapes.Item(i).name)
        If colShapes.Item(i).CellExists("User.Name", 0) Then
            itmx.SubItems(1) = colShapes.Item(i).Cells("User.Name").ResultStr(0)
        Else
            itmx.SubItems(1) = ""
        End If
    Next
    lblHeaders_Click
End Sub

Private Sub FillCollection(vsoLayer As Visio.Layer)
    Dim vsoShape As Visio.Shape
    Dim i As Integer
    
    Set colShapes = New Collection
    For Each vsoShape In ActivePage.Shapes
        For i = 1 To vsoShape.LayerCount
            If vsoShape.Layer(i).name = vsoLayer.name Then
                colShapes.Add vsoShape
                Exit For
            End If
        Next
    Next
End Sub

Private Sub ReSize() ' изменение высоты формы. Зависит от количества элементов в listbox
    Dim H As Single
    
    H = lstvShapes.ListItems.Count
  
    H = H * 12 + 12
    If H < 48 Then H = 48
    If H > 328 Then H = 328
    
    Me.Height = lstvPages.Top + H + 26

    lstvShapes.Height = H
    
End Sub

Private Sub lstvShapes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку
    With lstvShapes
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
End Sub

Private Sub lstvShapes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set vsoShp = ActivePage.Shapes.Item(Item.Key)
    If vsoShp.Parent.Type = visTypeGroup Then
        ActiveWindow.Select vsoShp, visDeselectAll + visSubSelect  ' выделение субшейпа
        '.CenterViewOnShape ActivePage.Shapes(shName), visCenterViewSelectShape '2010+
    Else
        ActiveWindow.Select vsoShp, visDeselectAll + visSelect     ' выделение шейпа
        '.CenterViewOnShape ActivePage.Shapes(shName) , visCenterViewSelectShape '2010+
        ActiveWindow.SetViewRect vsoShp.Cells("PinX") - pinWidth / 2, vsoShp.Cells("PinY") + pinHeight / 2, pinWidth, pinHeight
        '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End If
End Sub

Private Sub lblContent_Click() ' выровнять ширину столбцов по содержимому
   Dim colNum As Long
   For colNum = 0 To lstvShapes.ColumnHeaders.Count - 1
      Call SendMessage(lstvShapes.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvShapes.ColumnHeaders.Count - 1
      Call SendMessage(lstvShapes.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub

 Sub btnClose_Click() ' выгрузка формы
    With ActiveWindow
        .SetViewRect pinLeft, pinTop, pinWidth, pinHeight  'Восстановление вида окна после закрытия формы
                    '[левый] , [верхний] угол , [ширина] , [высота](вниз) видового окна
    End With
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub
