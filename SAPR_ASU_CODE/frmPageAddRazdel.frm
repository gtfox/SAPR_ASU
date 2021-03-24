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
    cmbxPageName.ListIndex = 3
    cmbxPageName.style = fmStyleDropDownList
    
    Fill_cmbxNomerShemy
    Fill_cmbxNomerFSA
    Fill_cmbxNaimenovLista
    
End Sub

Private Sub btnAddRazdel_Click()
    Dim vsoPageNew As Visio.Page
    Dim vsoPageSource As Visio.Page
    Dim shpRamka As Visio.Shape
    Dim shpRamkaSource As Visio.Shape
    Dim Ramka As Visio.Master
    Dim Setka As Visio.Master
    Dim PropPageSheet As String
    Dim PageName As String
    Dim i As Integer
    
    Set Ramka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("Рамка")
    Set Setka = Application.Documents.Item("SAPR_ASU_OFORM.vss").Masters.Item("SETKA KOORD")
    PageName = cmbxPageName.List(cmbxPageName.ListIndex, 0)

    Set vsoPageSource = GetSAPageExist(PageName)
    If vsoPageSource Is Nothing Then
        Set vsoPageNew = ActiveDocument.Pages.Add
        vsoPageNew.Name = PageName
        Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
        ActiveDocument.Masters.Item("Рамка").Delete
    Else
        Set vsoPageNew = vsoPageSource
        Set shpRamkaSource = GetSAShapeExist(vsoPageSource, "Рамка")
        If Not shpRamkaSource Is Nothing Then
            shpRamkaSource.Delete
            Set shpRamka = vsoPageNew.Drop(Ramka, 0, 0)
            ActiveDocument.Masters.Item("Рамка").Delete
        End If
    End If
    
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
    
    If PageName = cListNameCxema Then
        SetSA_NomerShemy vsoPageNew.PageSheet
        If cmbxNomerShemy.ListIndex <> -1 Then
            For i = 0 To cmbxNomerShemy.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNomerShemy.List(i) = "", "", cmbxNomerShemy.List(i) & IIf(i = cmbxNomerShemy.ListCount - 1, "", ";"))
            Next
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerShemy.Format").Formula = """" & PropPageSheet & """"
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerShemy").FormulaU = """INDEX(" & cmbxNomerShemy.ListIndex & ",Prop.SA_NomerShemy.Format)"""
        Else
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerShemy.Format").Formula = """" & cmbxNomerShemy.Text & """"
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerShemy").FormulaU = """INDEX(0,Prop.SA_NomerShemy.Format)"""
        End If
        vsoPageNew.Drop Setka, 0, 0
    End If
    If PageName = cListNameFSA Then
        SetSA_NomerFSA vsoPageNew.PageSheet
        If cmbxNomerFSA.ListIndex <> -1 Then
            For i = 0 To cmbxNomerFSA.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNomerFSA.List(i) = "", "", cmbxNomerFSA.List(i) & IIf(i = cmbxNomerFSA.ListCount - 1, "", ";"))
            Next
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerFSA.Format").Formula = """" & PropPageSheet & """"
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerFSA").FormulaU = """INDEX(" & cmbxNomerFSA.ListIndex & ",Prop.SA_NomerFSA.Format)"""
        Else
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerFSA.Format").Formula = """" & cmbxNomerFSA.Text & """"
            vsoPageNew.PageSheet.Cells("Prop.SA_NomerFSA").FormulaU = """INDEX(0,Prop.SA_NomerFSA.Format)"""
        End If
    End If
    
    LockTitleBlock
    
    Unload Me

End Sub

Sub Fill_cmbxNomerShemy()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameCxema
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NomerShemy.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNomerShemy.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNomerShemy.AddItem mstrPropPageSheet(i)
    Next
    cmbxNomerShemy.Text = ""
End Sub

Sub Fill_cmbxNomerFSA()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim mstrPropPageSheet() As String
    Dim i As Integer
    PageName = cListNameFSA
    For Each vsoPage In ActiveDocument.Pages
        If vsoPage.Name Like PageName & "*" Then
            PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NomerFSA.Format").ResultStr(0)
            Exit For
        End If
    Next
    cmbxNomerFSA.Clear
    mstrPropPageSheet = Split(PropPageSheet, ";")
    For i = 0 To UBound(mstrPropPageSheet)
        cmbxNomerFSA.AddItem mstrPropPageSheet(i)
    Next
    cmbxNomerFSA.Text = ""
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

Private Sub btnNomerShemyAdd_Click()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    If MsgBox("Добавить схему: " & cmbxNomerShemy.Text & vbNewLine & vbNewLine & "Это повлияет на все схемы в документе!", vbYesNo + vbInformation, "Добавить название схемы") = vbYes Then
        If cmbxNomerShemy.Text <> "" Then
            PageName = cListNameCxema
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.Name Like PageName & "*" Then
                    PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NomerShemy.Format").ResultStr(0)
                    vsoPage.PageSheet.Cells("Prop.SA_NomerShemy.Format").Formula = """" & PropPageSheet & ";" & cmbxNomerShemy.Text & """"
                End If
            Next
            Fill_cmbxNomerShemy
        End If
    End If
End Sub

Private Sub btnNomerFSAAdd_Click()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    If MsgBox("Добавить ФСА: " & cmbxNomerFSA.Text & vbNewLine & vbNewLine & "Это повлияет на все листы ФСА в документе!", vbYesNo + vbInformation, "Добавить название ФСА") = vbYes Then
        If cmbxNomerFSA.Text <> "" Then
            PageName = cListNameFSA
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.Name Like PageName & "*" Then
                    PropPageSheet = vsoPage.PageSheet.Cells("Prop.SA_NomerFSA.Format").ResultStr(0)
                    vsoPage.PageSheet.Cells("Prop.SA_NomerFSA.Format").Formula = """" & PropPageSheet & ";" & cmbxNomerFSA.Text & """"
                End If
            Next
            Fill_cmbxNomerFSA
        End If
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

Private Sub btnNomerShemyDel_Click()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim i As Integer
    If MsgBox("Удалить схему: " & cmbxNomerShemy.Text & vbNewLine & vbNewLine & "Это повлияет на все схемы в документе!", vbYesNo + vbCritical, "Удалить название схемы") = vbYes Then
        If cmbxNomerShemy.ListIndex <> -1 Then
            cmbxNomerShemy.RemoveItem cmbxNomerShemy.ListIndex
            For i = 0 To cmbxNomerShemy.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNomerShemy.List(i) = "", "", cmbxNomerShemy.List(i) & IIf(i = cmbxNomerShemy.ListCount - 1, "", ";"))
            Next
            PageName = cListNameCxema
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.Name Like PageName & "*" Then
                    vsoPage.PageSheet.Cells("Prop.SA_NomerShemy.Format").Formula = """" & PropPageSheet & """"
                End If
            Next
            Fill_cmbxNomerShemy
        End If
    End If
End Sub

Private Sub btnNomerFSADel_Click()
    Dim vsoPage As Visio.Page
    Dim PageName As String
    Dim PropPageSheet As String
    Dim i As Integer
    If MsgBox("Удалить ФСА: " & cmbxNomerFSA.Text & vbNewLine & vbNewLine & "Это повлияет на все листы ФСА в документе!", vbYesNo + vbCritical, "Удалить название ФСА") = vbYes Then
        If cmbxNomerFSA.ListIndex <> -1 Then
            cmbxNomerFSA.RemoveItem cmbxNomerFSA.ListIndex
            For i = 0 To cmbxNomerFSA.ListCount - 1
                PropPageSheet = PropPageSheet & IIf(cmbxNomerFSA.List(i) = "", "", cmbxNomerFSA.List(i) & IIf(i = cmbxNomerFSA.ListCount - 1, "", ";"))
            Next
            PageName = cListNameFSA
            For Each vsoPage In ActiveDocument.Pages
                If vsoPage.Name Like PageName & "*" Then
                    vsoPage.PageSheet.Cells("Prop.SA_NomerFSA.Format").Formula = """" & PropPageSheet & """"
                End If
            Next
            Fill_cmbxNomerFSA
        End If
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub