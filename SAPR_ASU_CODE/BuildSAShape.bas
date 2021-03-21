'------------------------------------------------------------------------------------------------------------
' Module        : BuildSAShape - Создание шейпов для САПР АСУ из обычных шейпов/графики
' Author        : gtfox на основе Shishok::CopyProperties_Module
' Date          : 2021.03.12
' Description   : Применение к шейпу заранее заданных свойств, копирование свойств
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе этого:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : CopyProperties_Module Копирование свойств
                ' Author    : Shishok
                ' Date      : 10.11.2016
                ' Purpose   : Копирование свойств одного шейпа/страницы/документа в другие шейпы/страницы/документы
                ' https://github.com/Shishok/, https://yadi.sk/d/qbpj9WI9d2eqF
                '------------------------------------------------------------------------------------------------------------
                
    ' CONSTANT                  VALUE
    '--------------------------------
    ' visSectionAction          240
    ' visSectionConnectionPts   7   только именованные
    ' visSectionScratch         6
    ' visSectionControls        9
    ' visSectionHyperlink       244
    ' visSectionProp            243
    ' visSectionSmartTag        247
    ' visSectionUser            242
    '--------------------------------
    'visSectionTextField
    'visRowField
    
    'visSectionCharacter
    'visRowCharacter
    
    'visSectionParagraph
    'visRowParagraph
    
    'visSectionTab
    'visRowTab

        
    'visSectionObject
    
    ' CONSTANT                      VALUE
    '------------------------------------
    ' visRowAlign                   14
    ' visRowEvent                   5+
    ' visRowFill                    3+
    ' visRowForeign                 9
    ' visRowGroup                   22+
    ' visRowHelpCopyright           16
    ' visRowImage                   21+
    ' visRowLayerMem                6+
    ' visRowLine                    2+
    ' visRowLock                    15+
    ' visRowMisc                    17+
    ' visRowShapeLayout             23+
    ' visRowTextXForm               12+
    ' visRowText                    11+
    ' visRowXForm1D                 4+
    ' visRowXFormOut                1+
    
Public arrCellName()
Public sSectionName As String

Sub GetAllSSValue()
    Dim arrSectionNumber()
    Dim arrRowNumber()
    Dim arrSectionNumberTEXT()
    Dim arrRowNumberTEXT()
    Dim i As Integer
    Dim j As Integer
    Dim strFile As String
    Dim vsoShape As Visio.Shape
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(6)
    
    strFile = ThisDocument.path & "tempValue.vb"
    
    arrSectionNumber = Array(visSectionUser, visSectionProp, visSectionHyperlink, visSectionConnectionPts, visSectionAction, visSectionControls, visSectionScratch, visSectionTextField, visSectionCharacter, visSectionParagraph, visSectionObject)
    arrRowNumber = Array(visRowXForm1D, visRowXFormOut, visRowLock, visRowMisc, visRowGroup, visRowLine, visRowFill, visRowText, visRowTextXForm, visRowLayerMem, visRowEvent, visRowImage, visRowShapeLayout)
    
    arrSectionNumberTEXT = Array(vbTab & "Select Case arrSectionNumber(i)" & vbNewLine & vbTab & vbTab & "Case visSectionUser 'User 242" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionProp 'Prop 243" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionHyperlink  'Hyperlink 244" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionConnectionPts 'ConnectionPts 7 только именованные" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionAction 'Action 240" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionControls 'Controls 9" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionScratch 'Scratch 6" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionTextField 'Text Field" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionCharacter 'Character" & vbNewLine, _
    vbNewLine & vbTab & vbTab & "Case visSectionParagraph 'Paragraph" & vbNewLine)

    arrRowNumberTEXT = Array(vbNewLine & vbTab & vbTab & "Case visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & "Select Case arrRowNumber(j)" & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowXForm1D '1-D Endpoints" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowXFormOut 'Shape Trannsform" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowLock 'Protection" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowMisc 'Miscellaneous + Glue Info" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowGroup 'Group Propeties" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowLine 'Line Format" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowFill 'Fill Format" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowText 'Text Block Format" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowTextXForm 'Text Transform" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowLayerMem 'Layer Membership" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowEvent 'Events" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowImage 'Image Propeties" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "'Case visRowMisc 'Glue Info" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "'arrValue = " & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowShapeLayout 'Shape Layout" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case Else" & vbNewLine & vbTab & vbTab & vbTab & "End Select" & vbNewLine & vbTab & vbTab & "Case Else" & vbNewLine & vbTab & "End Select")

    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                AddIntoTXTfile strFile, arrRowNumberTEXT(j)
                GetShapeSheetValue vsoShape, strFile, arrSectionNumber(i), arrRowNumber(j)
            Next
            AddIntoTXTfile strFile, arrRowNumberTEXT(j)
        Else
            AddIntoTXTfile strFile, arrSectionNumberTEXT(i)
            GetShapeSheetValue vsoShape, strFile, arrSectionNumber(i)
        End If
    Next
End Sub

Sub GetAllSSValueSplit()
    Dim arrSectionNumber()
    Dim arrRowNumber()
    Dim arrSectionNumberTEXT()
    Dim arrRowNumberTEXT()
    Dim i As Integer
    Dim j As Integer
    Dim strFile As String
    Dim vsoShape As Visio.Shape
    
    Set vsoShape = Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Master.34").Shapes("Sheet.5")
    
'    Set vsoShape = ActivePage.Shapes.ItemFromID(6)
    
    strFile = ThisDocument.path & "tempValue.vb"
    
    arrSectionNumber = Array(visSectionUser, visSectionProp, visSectionHyperlink, visSectionConnectionPts, visSectionAction, visSectionControls, visSectionScratch, visSectionTextField, visSectionCharacter, visSectionParagraph, visSectionObject)
    arrRowNumber = Array(visRowXForm1D, visRowXFormOut, visRowLock, visRowMisc, visRowGroup, visRowLine, visRowFill, visRowText, visRowTextXForm, visRowLayerMem, visRowEvent, visRowImage, visRowShapeLayout)
    
    arrSectionNumberTEXT = Array("SectionNumber = visSectionUser 'User 242" & vbNewLine & "sSectionName = ""User.""" & vbNewLine, _
    vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionProp 'Prop 243" & vbNewLine, _
    vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionHyperlink  'Hyperlink 244" & vbNewLine, _
    vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionConnectionPts 'ConnectionPts 7 только именованные" & vbNewLine, _
    vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionAction 'Action 240" & vbNewLine, _
    vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionControls 'Controls 9" & vbNewLine, _
    vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionScratch 'Scratch 6" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionTextField 'Text Field" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionCharacter 'Character" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionParagraph 'Paragraph" & vbNewLine)

    arrRowNumberTEXT = Array(vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowXForm1D '1-D Endpoints" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowXFormOut 'Shape Trannsform" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowLock 'Protection" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowMisc 'Miscellaneous + Glue Info" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowGroup 'Group Propeties" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowLine 'Line Format" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowFill 'Fill Format" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowText 'Text Block Format" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowTextXForm 'Text Transform" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowLayerMem 'Layer Membership" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowEvent 'Events" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowImage 'Image Propeties" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowShapeLayout 'Shape Layout" & vbNewLine, _
    vbNewLine & vbNewLine & "SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbTab & "")

    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                AddIntoTXTfile strFile, arrRowNumberTEXT(j)
                GetShapeSheetValue vsoShape, strFile, arrSectionNumber(i), arrRowNumber(j)
            Next
            AddIntoTXTfile strFile, arrRowNumberTEXT(j)
        Else
            AddIntoTXTfile strFile, arrSectionNumberTEXT(i)
            GetShapeSheetValue vsoShape, strFile, arrSectionNumber(i)
        End If
    Next
End Sub

Private Sub GetShapeSheetValue(vsoShape As Visio.Shape, strFile As String, ByVal SectionNumber As Long, Optional ByVal RowNumber As Long)
    Dim sRowName As String
    Dim arrValue()
    Dim arrRowNameValue()
    Dim i As Integer
    Dim j As Integer
    Dim UBarrCellName As Integer
    Dim UBarrValue As Integer
    Dim UBarrRowNameValue As Integer
    Dim ShpRowCount As Integer
    Dim strToFile As String
    Dim NoNameConnectionPts As Boolean
    
    SelectSection SectionNumber, RowNumber

    'Значения ячеек
    UBarrCellName = UBound(arrCellName)
    If SectionNumber = visSectionObject Then
        ShpRowCount = 0
    Else
        ShpRowCount = vsoShape.RowCount(SectionNumber) - 1
        If ShpRowCount = -1 Then Exit Sub
    End If
    ReDim arrValue(ShpRowCount)
    ReDim arrRowNameValue(ShpRowCount)
    For j = 0 To ShpRowCount

        If SectionNumber = visSectionConnectionPts And vsoShape.CellExists("Connections.X1", 0) Then 'ConnectionPts Не именованные
            If vsoShape.Cells("Connections.X1").RowName = "" Then NoNameConnectionPts = True
        End If

        If NoNameConnectionPts Then 'ConnectionPts Не именованные
            sRowName = ""
            arrRowNameValue(j) = ""
            For i = 0 To UBarrCellName
               arrValue(j) = arrValue(j) & vsoShape.Cells(sSectionName & Right(arrCellName(i), 1) & IIf(i > 1, "[" & CStr(j + 1) & "]", CStr(j + 1))).FormulaU & IIf(i = UBarrCellName, "", ";")
            Next
            NoNameConnectionPts = False
            
        ElseIf SectionNumber = visSectionScratch Then    'Scratch
            sRowName = ""
            arrRowNameValue(j) = ""
            For i = 0 To UBarrCellName
               arrValue(j) = arrValue(j) & vsoShape.Cells(sSectionName & arrCellName(i) & CStr(j + 1)).FormulaU & IIf(i = UBarrCellName, "", ";")
            Next
            
        ElseIf SectionNumber = visSectionTextField Or SectionNumber = visSectionCharacter Or SectionNumber = visSectionParagraph Or SectionNumber = visSectionObject Then   'Text Field + Character + Paragraph + SectionObject=Отдельные ячейки без строк
            arrRowNameValue(0) = ""
            For i = 0 To UBarrCellName
               arrValue(0) = arrValue(0) & vsoShape.Cells(sSectionName & arrCellName(i)).FormulaU & IIf(i = UBarrCellName, "", ";")
            Next
            
        Else 'Все остальные
            sRowName = vsoShape.CellsSRC(SectionNumber, j, 0).RowName 'Replace( , sSectionName, "")
            arrRowNameValue(j) = """" & sRowName & """"
            For i = 0 To UBarrCellName
               arrValue(j) = arrValue(j) & vsoShape.Cells(sSectionName & sRowName & arrCellName(i)).FormulaU & IIf(i = UBarrCellName, "", ";")
            Next
        End If
        
        arrValue(j) = """" & Replace(arrValue(j), """", """""") & """"
    Next
    
    If Len(arrRowNameValue(0)) <> 0 Then
        UBarrRowNameValue = UBound(arrRowNameValue)
        strToFile = strToFile & vbTab & vbTab & vbTab & "arrRowName = Array("
        For i = 0 To UBarrRowNameValue
            strToFile = strToFile & arrRowNameValue(i) & IIf(i = UBarrRowNameValue, ")" & vbNewLine, ", ")
        Next
    Else
        strToFile = strToFile & vbTab & vbTab & vbTab & "arrRowName = Array("""")" & vbNewLine & vbTab & vbTab
    End If
    
    strToFile = strToFile & vbTab & vbTab & vbTab & "arrValue = Array("
    UBarrValue = UBound(arrValue)
    For i = 0 To UBarrValue
        strToFile = strToFile & IIf(i > 0, vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab, "") & arrValue(i) & IIf(i = UBarrValue, ")", ", _" & vbNewLine)
    Next

    AddIntoTXTfile strFile, strToFile
End Sub

Sub SetValueToSelSections()
    Dim vsoShape As Visio.Shape
    Dim arrValue()
    Dim arrRowName()
    Dim SectionNumber As Long
    Dim RowNumber As Long




    Set vsoShape = ActivePage.Shapes.ItemFromID(163)
    

SectionNumber = visSectionUser 'User 242
sSectionName = "User."
            arrRowName = Array("KodProizvoditelyaDB", "KodPoziciiDB")
            arrValue = Array(";""""", _
                            ";""Код позиции/Код производителя/Код единицы""")
SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber

'SectionNumber = visSectionProp 'Prop 243
'            arrRowName = Array("SymName", "Number", "Autonum", "ShowDesc", "NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
'            arrValue = Array("""Букв. обозначение"";""Букв. обозначение"";0;;""KL"";""10"";FALSE;FALSE;1033;0", _
'                            """Номер элемента"";""Номер элемента"";2;"""";1;""20"";FALSE;FALSE;1033;0", _
'                            """Автонумерация"";""Автонумерация"";3;"""";FALSE;""90"";FALSE;FALSE;1033;0", _
'                            """Показать описание"";""Показать описание"";3;"""";FALSE;""80"";FALSE;FALSE;1049;0", _
'                            """Название из БД"";""Название из БД"";0;"""";"""";""60"";FALSE;FALSE;1033;0", _
'                            """Артикул из БД"";""Код заказа из БД"";0;"""";"""";""61"";FALSE;FALSE;1033;0", _
'                            """Производитель из БД"";""Производитель из БД"";0;"""";"""";""62"";FALSE;FALSE;1033;0", _
'                            """Цена из БД"";""Цена из БД"";0;"""";"""";""63"";FALSE;FALSE;1033;0", _
'                            """Единица из БД"";""Единица измерения из БД"";0;"""";"""";""64"";FALSE;FALSE;1033;0")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber


SectionNumber = visSectionProp 'Prop 243
            arrRowName = Array("NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
            arrValue = Array("""Название из БД"";""Название из БД"";0;"""";"""";""60"";FALSE;FALSE;1033;0", _
                            """Артикул из БД"";""Код заказа из БД"";0;"""";"""";""61"";FALSE;FALSE;1033;0", _
                            """Производитель из БД"";""Производитель из БД"";0;"""";"""";""62"";FALSE;FALSE;1033;0", _
                            """Цена из БД"";""Цена из БД"";0;"""";"""";""63"";FALSE;FALSE;1033;0", _
                            """Единица из БД"";""Единица измерения из БД"";0;"""";"""";""64"";FALSE;FALSE;1033;0")
SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber



SectionNumber = visSectionAction 'Action 240
            arrRowName = Array("Row_4")
            arrValue = Array("CALLTHIS(""DB.AddDBFrm"");""База данных..."";"""";264;"""";0;0;FALSE;FALSE;FALSE")
SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionObject
'RowNumber = visRowLock 'Protection
'            arrRowName = Array("")
'                    arrValue = Array("0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionTextField 'Text Field
'            arrRowName = Array("")
'                    arrValue = Array("FIELDPICTURE(0);User.Name;0;0")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionObject
'RowNumber = visRowXFormOut 'Shape Trannsform
'                    arrRowName = Array("")
'                    arrValue = Array(";;GUARD(IF(Actions.Row_2.Action,-90 deg,0 deg));;;;;;;1")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionObject
'RowNumber = visRowMisc 'Miscellaneous + Glue Info
'            arrRowName = Array("")
'                    arrValue = Array(";;;;;;;;;;9;;;;;;;;")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionObject
'RowNumber = visRowTextXForm 'Text Transform
'            arrRowName = Array("")
'                    arrValue = Array("TEXTWIDTH(TheText);TEXTHEIGHT(TheText,Height);IF(Actions.Row_2.Action,90 deg,0 deg);SETATREF(Controls.TextPos);SETATREF(Controls.TextPos.Y);TxtWidth*0;TxtHeight*0")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionObject
'RowNumber = visRowEvent 'Events
'            arrRowName = Array("")
'                    arrValue = Array(";;;;CALLTHIS(""ThisDocument.EventDropAutoNum"");CALLTHIS(""AutoNumber.AutoNum"")")
'SetValueToOneSection vsoShape, arrValue, arrRowName, SectionNumber, RowNumber


End Sub


Sub SetValueToAllSections()
    Dim arrSectionNumber()
    Dim arrRowNumber()
    Dim arrValue()
    Dim arrRowName()
    Dim i As Integer
    Dim j As Integer
    Dim vsoShape As Visio.Shape
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(131)

    arrSectionNumber = Array(visSectionUser, visSectionProp, visSectionHyperlink, visSectionConnectionPts, visSectionAction, visSectionControls, visSectionScratch, visSectionTextField, visSectionCharacter, visSectionParagraph, visSectionObject)
    arrRowNumber = Array(visRowXForm1D, visRowXFormOut, visRowLock, visRowMisc, visRowGroup, visRowLine, visRowFill, visRowText, visRowTextXForm, visRowLayerMem, visRowEvent, visRowImage, visRowShapeLayout)

    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                GoSub SubSelect
                SetValueToOneSection vsoShape, arrValue(), arrRowName(), arrSectionNumber(i), arrRowNumber(j)
                arrRowName = Array("")
                arrValue = Array("")
            Next
        Else
            GoSub SubSelect
            SetValueToOneSection vsoShape, arrValue(), arrRowName(), arrSectionNumber(i)
            arrRowName = Array("")
            arrValue = Array("")
        End If
    Next
    
    Exit Sub
    
    
SubSelect:

    Select Case arrSectionNumber(i)
        Case visSectionUser 'User 242
            arrRowName = Array("Location", "SAType", "Name", "Dropped", "StartNumDopKont", "KodProizvoditelyaDB", "KodPoziciiDB")
            arrValue = Array("", _
                            "2;", _
                            "Prop.SymName&Prop.Number;", _
                            "1;""""", _
                            "1;""""", _
                            "2;""""", _
                            "51319/2/1;""Код позиции/Код производителя/Код единицы""")
        Case visSectionProp 'Prop 243
            arrRowName = Array("SymName", "Number", "Autonum", "Kontaktov", "ShowDesc", "RTime", "NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
            arrValue = Array("""Букв. обозначение"";""Букв. обозначение"";0;;""KL"";""10"";FALSE;FALSE;1033;0", _
                            """Номер элемента"";""Номер элемента"";2;"""";1;""20"";FALSE;FALSE;1033;0", _
                            """Автонумерация"";""Автонумерация"";3;"""";FALSE;""90"";FALSE;FALSE;1033;0", _
                            """Число контактов"";""Максимум 10"";2;"""";4;""30"";FALSE;FALSE;1049;0", _
                            """Показать описание"";""Показать описание"";3;"""";FALSE;""80"";FALSE;FALSE;1049;0", _
                            """Реле времени"";""Реле времени"";3;"""";TRUE;""40"";FALSE;FALSE;1033;0", _
                            """Название из БД"";""Название из БД"";0;"""";""ARA автоматическое устройство повторного включения для iID 4P,1 программа           "";""60"";FALSE;FALSE;1033;0", _
                            """Артикул из БД"";""Код заказа из БД"";0;"""";""A9C70344"";""61"";FALSE;FALSE;1033;0", _
                            """Производитель из БД"";""Производитель из БД"";0;"""";""Schneider Electric"";""62"";FALSE;FALSE;1033;0", _
                            """Цена из БД"";""Цена из БД"";0;"""";""21784,997"";""63"";FALSE;FALSE;1033;0", _
                            """Единица из БД"";""Единица измерения из БД"";0;"""";""шт."";""64"";FALSE;FALSE;1033;0")
        Case visSectionHyperlink  'Hyperlink 244
            arrRowName = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
            arrValue = Array("""Контакт ""&User.Name&"".""&Scratch.B1&"":""&IF(Scratch.D1=0,""NO"",""NC"")&"" ""&Scratch.C1;"""";Scratch.A1;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A1,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B2&"":""&IF(Scratch.D2=0,""NO"",""NC"")&"" ""&Scratch.C2;"""";Scratch.A2;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A2,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B3&"":""&IF(Scratch.D3=0,""NO"",""NC"")&"" ""&Scratch.C3;"""";Scratch.A3;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A3,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B4&"":""&IF(Scratch.D4=0,""NO"",""NC"")&"" ""&Scratch.C4;"""";Scratch.A4;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A4,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B5&"":""&IF(Scratch.D5=0,""NO"",""NC"")&"" ""&Scratch.C5;"""";Scratch.A5;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A5,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B6&"":""&IF(Scratch.D6=0,""NO"",""NC"")&"" ""&Scratch.C6;"""";Scratch.A6;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A6,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B7&"":""&IF(Scratch.D7=0,""NO"",""NC"")&"" ""&Scratch.C7;"""";Scratch.A7;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A7,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B8&"":""&IF(Scratch.D8=0,""NO"",""NC"")&"" ""&Scratch.C8;"""";Scratch.A8;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A8,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B9&"":""&IF(Scratch.D9=0,""NO"",""NC"")&"" ""&Scratch.C9;"""";Scratch.A9;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A9,"""")", _
                            """Контакт ""&User.Name&"".""&Scratch.B10&"":""&IF(Scratch.D10=0,""NO"",""NC"")&"" ""&Scratch.C10;"""";Scratch.A10;"""";"""";"""";FALSE;FALSE;STRSAME(Scratch.A10,"""")")
        Case visSectionConnectionPts 'ConnectionPts 7 только именованные
            arrRowName = Array("a", "b")
            arrValue = Array("Width*0.5;Height*1;;;;""A1""", _
                            "Width*0.5;Height*0;;;;""A2""")
        Case visSectionAction 'Action 240
            arrRowName = Array("Row_1", "Row_2", "Row_3", "Row_4", "Row_5")
            arrValue = Array("CALLTHIS(""CrossReferenceRelay.AddReferenceRelayFrm"");""Связать..."";"""";"""";"""";;;FALSE;FALSE;FALSE", _
                            "NOT(""Actions.Action[2]"");IF(Actions.Row_2.Action,""Вертикально"",""Горизонтально"");"""";"""";"""";;;FALSE;FALSE;FALSE", _
                            "CALLTHIS(""CrossReferenceRelay.AddLocThumb"");""Вставить миниатюры контактов"";"""";"""";"""";0;0;FALSE;FALSE;FALSE", _
                            "SETF(GetRef(Actions.Row_4.Checked),NOT(Actions.Row_4.Checked));""Показать описание"";"""";"""";"""";0;0;FALSE;TRUE;FALSE", _
                            "CALLTHIS(""DB.AddDBFrm"");""База данных..."";"""";264;"""";0;0;FALSE;FALSE;FALSE")
        Case visSectionControls 'Controls 9
            arrRowName = Array("DescPos", "TextPos")
            arrValue = Array("Width*1.25;Height*0.33333333333333;Controls.DescPos;Controls.DescPos.Y;IF(Prop.ShowDesc,0,5);0;TRUE;0", _
                            "Width*1;Height*0.33333333333333;Controls.TextPos;Controls.TextPos.Y;0;0;TRUE;0")
        Case visSectionScratch 'Scratch 6
            arrRowName = Array("")
                    arrValue = Array(";;"""";IF(STRSAME(Scratch.A1,""""),0,User.StartNumDopKont);;", _
                            ";;"""";IF(STRSAME(Scratch.A2,""""),Scratch.B1,Scratch.B1+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A3,""""),Scratch.B2,Scratch.B2+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A4,""""),Scratch.B3,Scratch.B3+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A5,""""),Scratch.B4,Scratch.B4+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A6,""""),Scratch.B5,Scratch.B5+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A7,""""),Scratch.B6,Scratch.B6+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A8,""""),Scratch.B7,Scratch.B7+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A9,""""),Scratch.B8,Scratch.B8+1);;", _
                            ";;"""";IF(STRSAME(Scratch.A10,""""),Scratch.B9,Scratch.B9+1);;")
        Case visSectionTextField 'Text Field
            arrRowName = Array("")
                    arrValue = Array("FIELDPICTURE(0);User.Name;0;0")
        Case visSectionCharacter 'Character
            arrRowName = Array("")
                    arrValue = Array("93;11 pt;100%;0 pt;0;0%;2;0;0;FALSE;FALSE;FALSE;FALSE;0;0;0;-100%;1049")
        Case visSectionParagraph 'Paragraph
            arrRowName = Array("")
                    arrValue = Array("0 mm;0 mm;0 mm;-120%;0 pt;0 pt;1;0;"""";0;0;0 mm;-100%;0")
        Case visSectionObject 'Отдельные ячейки без строк
            Select Case arrRowNumber(j)
                Case visRowXForm1D '1-D Endpoints
                    arrRowName = Array("")
                    arrValue = Array(";;;")
                Case visRowXFormOut 'Shape Trannsform
                    arrRowName = Array("")
                    arrValue = Array("GUARD(10 mm);GUARD(15 mm);GUARD(IF(Actions.Row_2.Action,-90 deg,0 deg));30 mm;202.5 mm;Width*0.5;Height*0.5;;FALSE;1")
                Case visRowLock 'Protection
                    arrRowName = Array("")
                    arrValue = Array("0;0;0;0;0;0;0;0;0;0;0;0;1;0;0;0;0;0;0;0")
                Case visRowMisc 'Miscellaneous + Glue Info
                    arrRowName = Array("")
                    arrValue = Array("FALSE;FALSE;FALSE;FALSE;1033;FALSE;FALSE;0;FALSE;0;9;FALSE;"""";100%;FALSE;;;0;0")
                Case visRowGroup 'Group Propeties
                    arrRowName = Array("")
                    arrValue = Array("1;2;TRUE;TRUE;TRUE;FALSE")
                Case visRowLine 'Line Format
                    arrRowName = Array("")
                    arrValue = Array("1;0.2 mm;0;0;0;0;0%;2;2;0 mm")
                Case visRowFill 'Fill Format
                    arrRowName = Array("")
                    arrValue = Array("1;0%;0;0%;1;0;0%;1;0%;0;0 mm;0 mm;0;0 deg;100%")
                Case visRowText 'Text Block Format
                    arrRowName = Array("")
                    arrValue = Array("2 pt;1 pt;0;2 pt;1 pt;0%;0;1;15 mm")
                Case visRowTextXForm 'Text Transform
                    arrRowName = Array("")
                    arrValue = Array("TEXTWIDTH(TheText);TEXTHEIGHT(TheText,Height);IF(Actions.Row_2.Action,90 deg,0 deg);SETATREF(Controls.TextPos);SETATREF(Controls.TextPos.Y);TxtWidth*0;TxtHeight*0")
                Case visRowLayerMem 'Layer Membership
                    arrRowName = Array("")
                    arrValue = Array("")
                Case visRowEvent 'Events
                    arrRowName = Array("")
                    arrValue = Array(";;CALLTHIS(""CrossReferenceRelay.AddReferenceRelayFrm"");;CALLTHIS(""ThisDocument.EventDropAutoNum"");CALLTHIS(""AutoNumber.AutoNum"")")
                Case visRowImage 'Image Propeties
                    arrRowName = Array("")
                    arrValue = Array("50%;50%;0%;1;0%;0%;0%")
                'Case visRowMisc 'Glue Info
                    'arrValue =
                Case visRowShapeLayout 'Shape Layout
                    arrRowName = Array("")
                    arrValue = Array("FALSE;0;0;0;0;0;FALSE;0;0;0;0;FALSE;0;0;0;0")
                Case Else
            End Select
        Case Else
    End Select

Return

End Sub

Private Sub SetValueToOneSection(vsoShape As Visio.Shape, arrValue(), arrRowName(), ByVal SectionNumber As Long, Optional ByVal RowNumber As Long)
    Dim sRowName As String
    Dim arrCellValue() As String

    Dim i As Integer
    Dim j As Integer
    Dim UBarrCellName As Integer
    Dim UBarrCellValue As Integer
    Dim UBarrValue As Integer

    SelectSection SectionNumber, RowNumber

    UBarrCellName = UBound(arrCellName)
    
    If SectionNumber = visSectionObject Then
        UBarrValue = 0
    Else
        UBarrValue = UBound(arrValue)
    End If
    For j = 0 To UBarrValue
        AddSection vsoShape, SectionNumber
        arrCellValue = Split(arrValue(j), ";")
        UBarrCellValue = UBound(arrCellValue)
        If Len(arrRowName(0)) <> 0 Then
            sRowName = arrRowName(j)
            AddNamedRow vsoShape, SectionNumber, sRowName
        Else
            If Not (SectionNumber = visSectionScratch Or SectionNumber = visSectionTextField) Then
                AddRow vsoShape, SectionNumber
            End If
        End If
        
        On Error Resume Next
        If SectionNumber = visSectionConnectionPts And Len(sRowName) = 0 Then 'ConnectionPts Не именованные
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoShape.Cells(sSectionName & Right(arrCellName(i), 1) & IIf(i > 1, "[" & CStr(j + 1) & "]", CStr(j + 1))).FormulaU = arrCellValue(i)
                End If
            Next
            
        ElseIf SectionNumber = visSectionScratch Then    'Scratch
            If Not vsoShape.CellExists("Scratch.X" & CStr(j + 1), 0) Then AddRow vsoShape, SectionNumber
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoShape.Cells(sSectionName & arrCellName(i) & CStr(j + 1)).FormulaU = arrCellValue(i)
                End If
            Next
            
        ElseIf SectionNumber = visSectionTextField Or SectionNumber = visSectionCharacter Or SectionNumber = visSectionParagraph Or SectionNumber = visSectionObject Then   'Text Field + Character + Paragraph + SectionObject=Отдельные ячейки без строк
            If SectionNumber = visSectionTextField And (Not vsoShape.CellExists("Fields.Format", 0)) Then AddRow vsoShape, SectionNumber
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoShape.Cells(sSectionName & arrCellName(i)).FormulaU = arrCellValue(i)
                End If
            Next
            
        Else 'Все остальные
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoShape.Cells(sSectionName & sRowName & arrCellName(i)).FormulaU = arrCellValue(i)
                End If
            Next
        End If
    Next
End Sub


Sub SelectSection(ByVal SectionNumber As Long, ByVal RowNumber As Long)

    Select Case SectionNumber
        Case visSectionUser 'User 242
            sSectionName = "User."
            arrCellName = Array("", ".Prompt")
        Case visSectionProp 'Prop 243
            sSectionName = "Prop."
            arrCellName = Array(".Label", ".Prompt", ".Type", ".Format", "", ".SortKey", ".Invisible", ".Verify", ".LangID", ".Calendar")
        Case visSectionHyperlink  'Hyperlink 244
            sSectionName = "Hyperlink."
            arrCellName = Array("", ".Address", ".SubAddress", ".ExtraInfo", ".Frame", ".SortKey", ".NewWindow", ".Default", ".Invisible")
        Case visSectionConnectionPts 'ConnectionPts 7 только именованные
            sSectionName = "Connections."
            arrCellName = Array(".X", ".Y", ".A", ".B", ".C", ".D")
        Case visSectionAction 'Action 240
            sSectionName = "Actions."
            arrCellName = Array(".Action", ".Menu", ".TagName", ".ButtonFace", ".SortKey", ".Checked", ".Disabled", ".ReadOnly", ".Invisible", ".BeginGroup")
        Case visSectionControls 'Controls 9
            sSectionName = "Controls."
            arrCellName = Array("", ".Y", ".XDyn", ".YDyn", ".XCon", ".YCon", ".CanGlue", ".Type")
        Case visSectionScratch 'Scratch 6
            sSectionName = "Scratch."
            arrCellName = Array("X", "Y", "A", "B", "C", "D")   'X & i
        Case visSectionTextField 'Text Field
            sSectionName = "Fields."
            arrCellName = Array("Format", "Value", "Calendar", "ObjectKind")
        Case visSectionCharacter 'Character
            sSectionName = "Char."
            arrCellName = Array("Font", "Size", "FontScale", "Letterspace", "Color", "ColorTrans", "Style", "Case", "Pos", "Strikethru", "DblUnderline", "Overline", "DoubleStrikethrough", "AsianFont", "ComplexScriptFont", "LocalizeFont", "ComplexScriptSize", "LangID")
        Case visSectionParagraph 'Paragraph
            sSectionName = "Para."
            arrCellName = Array("IndFirst", "IndLeft", "IndRight", "SpLine", "SpBefore", "SpAfter", "HorzAlign", "Bullet", "BulletStr", "BulletFont", "LocalizeBulletFont", "TextPosAfterBullet", "BulletFontSize", "Flags")
        Case visSectionObject 'Отдельные ячейки без строк
            sSectionName = ""
            Select Case RowNumber
                Case visRowXForm1D '1-D Endpoints
                    arrCellName = Array("BeginX", "BeginY", "EndX", "EndY")
                Case visRowXFormOut 'Shape Trannsform
                    arrCellName = Array("Width", "Height", "Angle", "PinX", "PinY", "LocPinX", "LocPinY", "FlipX", "FlipY", "ResizeMode")
                Case visRowLock 'Protection
                    arrCellName = Array("LockWidth", "LockHeight", "LockAspect", "LockMoveX", "LockMoveY", "LockRotate", "LockBegin", "LockEnd", "LockDelete", "LockSelect", "LockFormat", "LockCustProp", "LockTextEdit", "LockVtxEdit", "LockCrop", "LockGroup", "LockCalcWH", "LockFromGroupFormat", "LockThemeColors", "LockThemeEffects")
                Case visRowMisc 'Miscellaneous + Glue Info
                    arrCellName = Array("NoObjHandles", "NoCtlHandles", "NoAlignBox", "NonPrinting", "LangID", "HideText", "UpdateAlignBox", "DynFeedback", "NoLiveDynamics", "Calendar", "ObjType", "IsDropSource", "Comment", "DropOnPageScale", "LocalizeMerge", "BegTrigger", "EndTrigger", "GlueType", "WalkPreference")
                Case visRowGroup 'Group Propeties
                    arrCellName = Array("SelectMode", "DisplayMode", "IsTextEditTarget", "IsSnapTarget", "IsDropTarget", "DontMoveChildren")
                Case visRowLine 'Line Format
                    arrCellName = Array("LinePattern", "LineWeight", "LineColor", "LineCap", "BeginArrow", "EndArrow", "LineColorTrans", "BeginArrowSize", "EndArrowSize", "Rounding")
                Case visRowFill 'Fill Format
                    arrCellName = Array("FillForegnd", "FillForegndTrans", "FillBkgnd", "FillBkgndTrans", "FillPattern", "ShdwForegnd", "ShdwForegndTrans", "ShdwBkgnd", "ShdwBkgndTrans", "ShdwPattern", "ShapeShdwOffsetX", "ShapeShdwOffsetY", "ShapeShdwType", "ShapeShdwObliqueAngle", "ShapeShdwScaleFactor")
                Case visRowText 'Text Block Format
                    arrCellName = Array("LeftMargin", "RightMargin", "TextBkgnd", "TopMargin", "BottomMargin", "TextBkgndTrans", "TextDirection", "VerticalAlign", "DefaultTabStop")
                Case visRowTextXForm 'Text Transform
                    arrCellName = Array("TxtWidth", "TxtHeight", "TxtAngle", "TxtPinX", "TxtPinY", "TxtLocPinX", "TxtLocPinY")
                Case visRowLayerMem 'Layer Membership
                    arrCellName = Array("LayerMember")
                Case visRowEvent 'Events
                    arrCellName = Array("TheData", "TheText", "EventDblClick", "EventXFMod", "EventDrop", "EventMultiDrop")
                Case visRowImage 'Image Propeties
                    arrCellName = Array("Contrast", "Brightness", "Transparency", "Gamma", "Blur", "Sharpen", "Denoise")
'                Case visRowMisc 'Glue Info
'                    arrCellName = Array("BegTrigger", "EndTrigger", "GlueType", "WalkPreference")
                Case visRowShapeLayout 'Shape Layout
                    arrCellName = Array("ShapePermeableX", "ShapeFixedCode", "ConLineJumpDirX", "ConLineJumpCode", "ShapePlaceFlip", "ShapePlaceStyle", "ShapePermeableY", "ShapePlowCode", "ConLineJumpDirY", "ConLineJumpStyle", "ConLineRouteExt", "ShapePermeablePlace", "ShapeRouteStyle", "ConFixedCode", "ShapeSplit", "ShapeSplittable")
                Case Else
            End Select
        Case Else
    End Select

End Sub

Private Sub AddRow(vsoShape As Visio.Shape, ByVal SectionNumber As Long)
    On Error Resume Next
    vsoShape.AddRow SectionNumber, visRowLast, visTagDefault
End Sub

Private Sub AddNamedRow(vsoShape As Visio.Shape, ByVal SectionNumber As Long, ByVal sRowName As String)
    On Error Resume Next
    vsoShape.AddNamedRow SectionNumber, sRowName, visTagDefault
End Sub

Private Sub AddSection(vsoShape As Visio.Shape, ByVal SectionNumber As Long)
    If Not vsoShape.SectionExists(SectionNumber, 0) Then
        vsoShape.AddSection SectionNumber
    End If
End Sub

'Private Function AddNamedRow(vsoShape As Visio.Shape, ByVal SectionNumber As Long, ByVal sRowName As String) As Boolean
'    On Error GoTo err
'        vsoShape.AddNamedRow SectionNumber, sRowName, visTagDefault
'        AddNamedRow = True
'    Exit Function
'err:
'        AddNamedRow = False
'End Function


Function AddIntoTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(filename, 8, True): ts.Write txt: ts.Close
    Set ts = Nothing: Set fso = Nothing
    AddIntoTXTfile = err = 0
End Function


Private Sub GetCellsNames()
    ' Получение имен ячеек
    Dim Section As Integer
    Dim Row As Integer
    Dim vsoCellF As Visio.Cell, r As Integer, i As Integer, j As Integer, booAddRow As Boolean
    Dim vsoShape As Visio.Shape
    Dim strFile As String
    Dim strToFile As String

    strFile = ThisDocument.path & "tempName.vb"
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(131)
    
    Section = visSectionObject
    Row = visRowXForm1D

    'Имена ячеек
    
    For i = 0 To vsoShape.RowsCellCount(Section, Row)
        strToFile = strToFile & """." & vsoShape.CellsSRC(Section, Row, i).LocalName & """, "
    Next

     AddIntoTXTfile strFile, strToFile & vbNewLine & vbNewLine

End Sub