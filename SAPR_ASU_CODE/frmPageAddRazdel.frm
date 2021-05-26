

Option Explicit

Dim NaimenovanieAdd2Ramka As String


Private Sub UserForm_Initialize()

    cmbxPageName.AddItem cListNameOD '"ОД" 'Общие указания
    cmbxPageName.AddItem cListNameFSA '"ФСА" 'Схема функциональная автоматизации
    cmbxPageName.AddItem cListNamePlan '"План" 'План расположения оборудования и приборов КИП
    cmbxPageName.AddItem cListNameCxema '"Схема" 'Схема электрическая принципиальная
    cmbxPageName.AddItem cListNameVID '"ВИД" 'Чертеж внешнего вида шкафа
    cmbxPageName.AddItem cListNameSVP '"СВП" 'Схема соединения внешних проводок
    cmbxPageName.AddItem cListNameSpec '"С" 'Спецификация оборудования, изделий и материалов
'    cmbxPageName.ListIndex = 3
    cmbxPageName.style = fmStyleDropDownList
    
    frameCx.Visible = False
    frameFSA.Visible = False
    frameCx.Top = 30
    frameFSA.Top = 30
    frameNaim.Top = 30
    Me.Height = frameNaim.Top + frameNaim.Height + 24
    
    Fill_cmbxNazvanieShemy
    Fill_cmbxNazvanieFSA
    Fill_cmbxNaimenovLista
    
End Sub

Private Sub cmbxPageName_Change()
    Select Case cmbxPageName.List(cmbxPageName.ListIndex, 0)
        Case cListNameCxema '"Схема"
            frameCx.Visible = True
            frameFSA.Visible = False
            frameNaim.Top = 60
        Case cListNameFSA '"ФСА"
            frameCx.Visible = False
            frameFSA.Visible = True
            frameNaim.Top = 60
        Case Else
            frameCx.Visible = False
            frameFSA.Visible = False
            frameNaim.Top = 30
    End Select
    Me.Height = frameNaim.Top + frameNaim.Height + 24
End Sub

Private Sub btnAddRazdel_Click()
    Dim vsoPage As Visio.Page
    Dim vsoPageNew As Visio.Page
    Dim vsoPageSource As Visio.Page
    Dim vsoPageLast As Visio.Page
    Dim shpRamka As Visio.Shape
    Dim shpRamkaSource As Visio.Shape
    Dim Ramka As Visio.Master
    Dim Setka As Visio.Master
    Dim colPagesAll As Collection
    Dim PropPageSheet As String
    Dim PageName As String
    Dim PageNumber As Integer
    Dim MaxNpage As Integer
    Dim Index As Integer
    Dim i As Integer

    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    Set Setka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("SETKA KOORD")
    If cmbxPageName.ListIndex = -1 Then Exit Sub
    PageName = cmbxPageName.List(cmbxPageName.ListIndex, 0)

    Set vsoPageSource = GetSAPageExist(PageName)
    If vsoPageSource Is Nothing Then
        Set vsoPageNew = ActiveDocument.Pages.Add
        vsoPageNew.Name = PageName
    Else
        Set colPagesAll = New Collection
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.Name Like PageName & "*" Then
                colPagesAll.Add vsoPage
                If vsoPage.Index > Index Then Index = vsoPage.Index: Set vsoPageLast = vsoPage
            End If
        Next
        PageNumber = GetPageNumber(vsoPageLast.Name)
        'Находим максимальный номер страницы в NameU и Name
        MaxNpage = MaxMinPageNumber(colPagesAll, , , True)
        'Создаем страницу раздела с максимальным номером
        Set vsoPageNew = ActiveDocument.Pages.Add
        vsoPageNew.Name = PageName & "." & CStr(MaxNpage + 1)
        'Переименовываем вставленный лист в нумерацию Name после последнего
        vsoPageNew.Name = PageName & "." & CStr(PageNumber + 1)
        'Положение новой страницы сразу за последним
        vsoPageNew.Index = Index + 1
    End If

    Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
    ActiveDocument.Masters.Item("Рамка").Delete
        
    If cmbxNaimenovLista.ListIndex = -1 Then
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(0,Prop.CHAPTER.Format)"
        shpRamka.Cells("Prop.Type.Format").FormulaU = """" & shpRamka.Cells("Prop.Type.Format").ResultStr(0) & ";" & cmbxNaimenovLista.Text & """"
        shpRamka.Cells("Prop.Type").FormulaU = "INDEX(" & cmbxNaimenovLista.ListCount & ",Prop.Type.Format)"
        shpRamka.Cells("Prop.CNUM").Formula = 0
        shpRamka.Cells("Prop.TNUM").Formula = 0
    Else
        shpRamka.Cells("Prop.CHAPTER").FormulaU = "INDEX(0,Prop.CHAPTER.Format)"
        shpRamka.Cells("Prop.Type").FormulaU = "INDEX(" & cmbxNaimenovLista.ListIndex & ",Prop.Type.Format)"
        shpRamka.Cells("Prop.CNUM").Formula = 0
        shpRamka.Cells("Prop.TNUM").Formula = 0
    End If

    If chbA4 Then
        vsoPageNew.PageSheet.Cells("PageWidth").Formula = "210 MM"
        vsoPageNew.PageSheet.Cells("PageHeight").Formula = "297 MM"
        vsoPageNew.PageSheet.Cells("Paperkind").Formula = 9
        vsoPageNew.PageSheet.Cells("PrintPageOrientation").Formula = 1
    Else
        vsoPageNew.PageSheet.Cells("PageWidth").Formula = "420 MM"
        vsoPageNew.PageSheet.Cells("PageHeight").Formula = "297 MM"
        vsoPageNew.PageSheet.Cells("Paperkind").Formula = 8
        vsoPageNew.PageSheet.Cells("PrintPageOrientation").Formula = 2
    End If
        vsoPageNew.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawingScale).FormulaU = "1 mm"
        vsoPageNew.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawScaleType).FormulaU = "0"
    
    Select Case PageName
        Case cListNameOD ' "ОД" 'Общие указания
        Case cListNameFSA ' "ФСА" 'Схема функциональная автоматизации
            SetNazvanieFSA vsoPageNew.PageSheet
            If cmbxNazvanieFSA.ListIndex = -1 Then NazvanieFSAAdd
            For i = 0 To cmbxNazvanieFSA.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNazvanieFSA.List(i) = "", "", cmbxNazvanieFSA.List(i) & IIf(i = cmbxNazvanieFSA.ListCount - 1, "", ";"))
            Next
            vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = """" & PropPageSheet & """"
            If cmbxNazvanieFSA.ListIndex <> -1 Then
                vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA").FormulaU = "INDEX(" & cmbxNazvanieFSA.ListIndex & ",Prop.SA_NazvanieFSA.Format)"
            Else
                vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieFSA").FormulaU = "INDEX(" & cmbxNazvanieFSA.ListCount - 1 & ",Prop.SA_NazvanieFSA.Format)"
            End If
            
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить оборудование со схемы"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageFSAAddSensorsFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
            End With

        Case cListNamePlan ' "План" 'План расположения оборудования и приборов КИП
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить оборудование из ФСА"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PagePLANAddElementsFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """10"""
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Проложить кабели для всего оборудования"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""AddRouteCablesOnPlan"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "2633" '2645
                .CellsSRC(visSectionAction, visRowLast, visActionSortKey).FormulaU = """10"""
            End With
        Case cListNameCxema ' "Схема" 'Схема электрическая принципиальная
            SetNazvanieShemy vsoPageNew.PageSheet
            If cmbxNazvanieShemy.ListIndex = -1 Then NazvanieShemyAdd
            For i = 0 To cmbxNazvanieShemy.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNazvanieShemy.List(i) = "", "", cmbxNazvanieShemy.List(i) & IIf(i = cmbxNazvanieShemy.ListCount - 1, "", ";"))
            Next
            vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").Formula = """" & PropPageSheet & """"
            If cmbxNazvanieShemy.ListIndex <> -1 Then
                vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShemy").FormulaU = "INDEX(" & cmbxNazvanieShemy.ListIndex & ",Prop.SA_NazvanieShemy.Format)"
            Else
                vsoPageNew.PageSheet.Cells("Prop.SA_NazvanieShemy").FormulaU = "INDEX(" & cmbxNazvanieShemy.ListCount - 1 & ",Prop.SA_NazvanieShemy.Format)"
            End If
            vsoPageNew.Drop Setka, 0, 0
        Case cListNameVID ' "ВИД" 'Чертеж внешнего вида шкафа
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить элементы со схемы"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageVIDAddElementsFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
            End With
        Case cListNameSVP ' "СВП" 'Схема соединения внешних проводок
            With vsoPageNew.PageSheet
                .AddSection visSectionAction
                .AddRow visSectionAction, visRowLast, visTagDefault
                .CellsSRC(visSectionAction, visRowLast, visActionMenu).FormulaForceU = """Вставить провода со схемы"""
                .CellsSRC(visSectionAction, visRowLast, visActionAction).FormulaForceU = "RunMacro(""PageSVPAddKabeliFrm"")"
                .CellsSRC(visSectionAction, visRowLast, visActionButtonFace).FormulaForceU = "1104" '1753
            End With
        Case cListNameSpec ' "С" 'Спецификация оборудования, изделий и материалов
        Case Else
    End Select

    LockTitleBlock
    
    ActiveWindow.DeselectAll
    
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me

End Sub

Sub Fill_cmbxNazvanieShemy()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieShemy.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieShemy.AddItem mstrPropPageSheet(i)
    Next
    cmbxNazvanieShemy.Text = ""
End Sub

Sub Fill_cmbxNazvanieFSA()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNazvanieFSA.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNazvanieFSA.AddItem mstrPropPageSheet(i)
    Next
    cmbxNazvanieFSA.Text = ""
End Sub

Sub Fill_cmbxNaimenovLista()
    Dim Ramka As Visio.Shape
    Dim PropShapeSheet As String
    Dim mstrPropShapeSheet() As String
    Dim i As Integer
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка").Shapes("Рамка")
    PropShapeSheet = Ramka.Cells("Prop.Type.Format").ResultStr(0)
    cmbxNaimenovLista.Clear
    mstrPropShapeSheet = Split(PropShapeSheet, ";")
    For i = 0 To UBound(mstrPropShapeSheet)
        cmbxNaimenovLista.AddItem mstrPropShapeSheet(i)
    Next
End Sub

Private Sub btnNazvanieShemyAdd_Click()
    If MsgBox("Добавить схему: " & cmbxNazvanieShemy.Text & vbNewLine & vbNewLine & "Это повлияет на все схемы в документе!", vbYesNo + vbInformation, "Добавить название схемы") = vbYes Then
        If cmbxNazvanieShemy.Text = "" Then
            MsgBox "Название схемы пустое" & vbNewLine & "Введите название схемы... ", vbExclamation, "Название схемы пустое"
        Else
            NazvanieShemyAdd
        End If
    End If
End Sub

Sub NazvanieShemyAdd()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    If cmbxNazvanieShemy.Text <> "" Then
        PageName = cListNameCxema
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.Name Like PageName & "*" Then
                PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").ResultStr(0)
                vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").Formula = """" & PropPageSheet & IIf(PropPageSheet = "", "", ";") & cmbxNazvanieShemy.Text & """"
            End If
        Next
        Fill_cmbxNazvanieShemy
    End If
End Sub

Private Sub btnNazvanieFSAAdd_Click()
    If MsgBox("Добавить ФСА: " & cmbxNazvanieFSA.Text & vbNewLine & vbNewLine & "Это повлияет на все листы ФСА в документе!", vbYesNo + vbInformation, "Добавить название ФСА") = vbYes Then
        If cmbxNazvanieFSA.Text = "" Then
            MsgBox "Название ФСА пустое" & vbNewLine & "Введите название ФСА... ", vbExclamation, "Название ФСА пустое"
        Else
            NazvanieFSAAdd
        End If
    End If
End Sub

Sub NazvanieFSAAdd()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    If cmbxNazvanieFSA.Text <> "" Then
        PageName = cListNameFSA
        For Each vsoPage In ActiveDocument.Pages
            If vsoPage.Name Like PageName & "*" Then
                PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").ResultStr(0)
                vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = """" & PropPageSheet & IIf(PropPageSheet = "", "", ";") & cmbxNazvanieFSA.Text & """"
            End If
        Next
        Fill_cmbxNazvanieFSA
    End If
End Sub

Private Sub btnNaimenovanieAdd2Master_Click()
    Dim Ramka As Visio.Shape
    Dim PropShapeSheet As String
    If MsgBox("Добавить наименование листа в шаблон рамки: " & cmbxNaimenovLista.Text & vbNewLine & vbNewLine & "Это повлияет на все будущие рамки всех разделов!" & vbNewLine & "Запись попадет в рамку в наборе элементов SAPR_ASU_OFORM.vss" & vbNewLine & "Чтобы это произошло набор элементов должен быть переведен в режим редактирования (изменения)", vbYesNo + vbExclamation, "Добавить Наименование листа в Шаблон рамки") = vbYes Then
        Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка").Shapes("Рамка")
        PropShapeSheet = Ramka.Cells("Prop.Type.Format").ResultStr(0)
        Ramka.Cells("Prop.Type.Format").Formula = """" & PropShapeSheet & ";" & cmbxNaimenovLista.Text & """"
        Fill_cmbxNaimenovLista
    End If
End Sub

Private Sub btnNazvanieShemyDel_Click()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim i As Integer
    If MsgBox("Удалить схему: " & cmbxNazvanieShemy.Text & vbNewLine & vbNewLine & "Это повлияет на все схемы в документе!", vbYesNo + vbCritical, "Удалить название схемы") = vbYes Then
        If cmbxNazvanieShemy.ListIndex <> -1 Then
            cmbxNazvanieShemy.RemoveItem cmbxNazvanieShemy.ListIndex
            For i = 0 To cmbxNazvanieShemy.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNazvanieShemy.List(i) = "", "", cmbxNazvanieShemy.List(i) & IIf(i = cmbxNazvanieShemy.ListCount - 1, "", ";"))
            Next
            PageName = cListNameCxema
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.Name Like PageName & "*" Then
                    vsoPage.PageSheet.Cells("Prop.SA_NazvanieShemy.Format").Formula = """" & PropPageSheet & """"
                End If
            Next
            Fill_cmbxNazvanieShemy
        End If
    End If
End Sub

Private Sub btnNazvanieFSADel_Click()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim i As Integer
    If MsgBox("Удалить ФСА: " & cmbxNazvanieFSA.Text & vbNewLine & vbNewLine & "Это повлияет на все листы ФСА в документе!", vbYesNo + vbCritical, "Удалить название ФСА") = vbYes Then
        If cmbxNazvanieFSA.ListIndex <> -1 Then
            cmbxNazvanieFSA.RemoveItem cmbxNazvanieFSA.ListIndex
            For i = 0 To cmbxNazvanieFSA.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNazvanieFSA.List(i) = "", "", cmbxNazvanieFSA.List(i) & IIf(i = cmbxNazvanieFSA.ListCount - 1, "", ";"))
            Next
            PageName = cListNameFSA
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.Name Like PageName & "*" Then
                    vsoPage.PageSheet.Cells("Prop.SA_NazvanieFSA.Format").Formula = """" & PropPageSheet & """"
                End If
            Next
            Fill_cmbxNazvanieFSA
        End If
    End If
End Sub

Private Sub btnClose_Click()
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    Unload Me
End Sub