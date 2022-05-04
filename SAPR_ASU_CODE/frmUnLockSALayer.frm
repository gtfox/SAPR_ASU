Option Explicit
'------------------------------------------------------------------------------------------------------------
' Module        : frmUnLockSALayer - Форма разблокировки заблокированных шейпов
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

    lstvSapes.LabelEdit = lvwManual 'чтобы не редактировалось первое значение в строке
    lstvSapes.ColumnHeaders.Add , , "Шейп"
    lstvSapes.ColumnHeaders.Add , , "Название"
    
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
        ActivePage.Layers(cmbxLayers.Text).Remove vsoShp, 0
        Fill_lstvSapes
    End If
End Sub

Private Sub cbLockUnlockLayer_Click()
    If cmbxLayers.Text <> "" Then
        If ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerLock).Result(0) = 1 Then
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerLock).FormulaU = "0"
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerColor).FormulaU = "255"
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerSnap).FormulaU = "0"
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerGlue).FormulaU = "0"
            cbLockUnlockLayer.Caption = "Заблокировать слой"
        Else
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerLock).FormulaU = "1"
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerColor).FormulaU = "19"
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerSnap).FormulaU = "0"
            ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerGlue).FormulaU = "0"
            cbLockUnlockLayer.Caption = "Разблокировать слой"
        End If
    End If
End Sub

Private Sub cmbxLayers_Change()
    If ActivePage.Layers(cmbxLayers.Text).CellsC(visLayerLock).Result(0) = 1 Then cbLockUnlockLayer.Caption = "Разблокировать слой" Else cbLockUnlockLayer.Caption = "Заблокировать слой"
    Fill_lstvSapes
End Sub

Sub Fill_cmbxLayers()
    Dim vsoLayer As Visio.Layer
    cmbxLayers.Clear
    For Each vsoLayer In ActivePage.Layers
        cmbxLayers.AddItem vsoLayer.name, vsoLayer.Index - 1
    Next
End Sub

Sub Fill_lstvSapes()
    Dim itmx As ListItem
    Dim i As Integer
    FillCollection ActivePage.Layers(cmbxLayers.Text)
    lstvSapes.ListItems.Clear
    For i = 1 To colShapes.Count
        Set itmx = lstvSapes.ListItems.Add(, colShapes.Item(i).NameID, colShapes.Item(i).NameID)
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
    
    H = lstvSapes.ListItems.Count
  
    H = H * 12 + 12
    If H < 48 Then H = 48
    If H > 328 Then H = 328
    
    Me.Height = lstvPages.Top + H + 26

    lstvSapes.Height = H
    
End Sub

Private Sub lstvSapes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader) ' сортировка при клике по заголовку
    With lstvSapes
        .Sorted = False
        .SortKey = ColumnHeader.SubItemIndex
        'изменить порядок сортировки на обратный имеющемуся
        .SortOrder = Abs(.SortOrder Xor 1)
        .Sorted = True
    End With
End Sub

Private Sub lstvSapes_ItemClick(ByVal Item As MSComctlLib.ListItem)
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
   For colNum = 0 To lstvSapes.ColumnHeaders.Count - 1
      Call SendMessage(lstvSapes.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE)
   Next
End Sub

Private Sub lblHeaders_Click() ' выровнять ширину столбцов по заголовкам
   Dim colNum As Long
   For colNum = 0 To lstvSapes.ColumnHeaders.Count - 1
      Call SendMessage(lstvSapes.hWnd, LVM_SETCOLUMNWIDTH, colNum, ByVal LVSCW_AUTOSIZE_USEHEADER)
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