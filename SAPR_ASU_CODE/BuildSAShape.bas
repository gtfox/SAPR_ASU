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

Sub SetElement() 'SetValueToSelSections
    Dim vsoObject As Object
    Dim mastshp As Visio.Shape
    Dim vsoShp As Visio.Shape
    
    Dim arrRowValue()
    Dim arrRowName()
    Dim arrMast()
    Dim SectionNumber As Long
    Dim RowNumber As Long

'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("PLCParent").Shapes.Item("PLCParent").Shapes.Item("PLCModParent")

'Set vsoObject = ActivePage.Shapes.ItemFromID(219)

'Set vsoObject = ActiveWindow.Selection(1)
    
'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("PodvalCxemy").Shapes.Item(1)

'arrMast = Array("QS", "SF", "QF", "QA", "KK", "FU3P", "FU", "KL", "KM", "KT", "KV", "UZF", "UZ3P", "UZ", "SA", "SB", "HL", "HA", "EK3P", "EK", "XS", "XS3P", "TV", "UG", "R", "R3P", "RU3P", "RU", "Sensor", "Term", "TermC")
'arrMast = Array("w1", "w2", "w3")
'
'
'For i = 0 To UBound(arrMast)
'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item(arrMast(i)).Shapes.Item(arrMast(i))
'



'SectionNumber = visSectionUser 'User 242
'sSectionName = "User."
'            arrRowName = Array("Name", "Shkaf", "Mesto")
'            arrRowValue = Array("IF(Prop.HideNumber,"""",IF(AND(TheDoc!User.SA_ISO,Prop.Oboz_ISO),IF(STRSAME(User.Mesto,""""),"""",TheDoc!User.SA_PrefMesto&User.Mesto&IF(Prop.PerenosOboz,CHAR(10),""""))&IF(STRSAME(User.Shkaf,""""),"""",TheDoc!User.SA_PrefShkaf&User.Shkaf&IF(Prop.PerenosOboz,CHAR(10),""""))&TheDoc!User.SA_PrefElement,"""")&Prop.Number)&IF(Prop.HideName,"""",IF(Prop.HideNumber,"""","": "")&Prop.SymName)|", _
'                            "ThePage!Prop.SA_NazvanieShkafa|""""", _
'                            "ThePage!Prop.SA_NazvanieMesta|""""")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionProp 'Prop 243
'            arrRowName = Array("PerenosOboz", "Oboz_ISO")
'            arrRowValue = Array("""Перенос обозн.""|""Переносить обозначение (обозначение в столбец)""|3|""""|FALSE|""69""|NOT(TheDoc!User.SA_ISO)|FALSE|1049|0", _
'                            """Разделители обозн.""|""Использовать обозначение с разделителями""|3|""""|FALSE|""68""|NOT(TheDoc!User.SA_ISO)|FALSE|1049|0")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

    
'
'vsoObject.Cells("User.FullName").FormulaU = "IF(and(TheDoc!User.SA_ISO,Prop.Oboz_ISO),IF(STRSAME(User.Mesto,""""),"""",TheDoc!User.SA_PrefMesto&User.Mesto&IF(Prop.PerenosOboz,CHAR(10),""""))&IF(STRSAME(User.Shkaf,""""),"""",TheDoc!User.SA_PrefShkaf&User.Shkaf&IF(Prop.PerenosOboz,CHAR(10),""""))&TheDoc!User.SA_PrefElement,"""")&User.KlemmnikName&"":""&Prop.Number"
'vsoObject.Cells("User.Name").FormulaU = "IF(and(TheDoc!User.SA_ISO,Prop.Oboz_ISO),IF(STRSAME(User.Mesto,""""),"""",TheDoc!User.SA_PrefMesto&User.Mesto&IF(Prop.PerenosOboz,CHAR(10),""""))&IF(STRSAME(User.Shkaf,""""),"""",TheDoc!User.SA_PrefShkaf&User.Shkaf&IF(Prop.PerenosOboz,CHAR(10),""""))&TheDoc!User.SA_PrefElement,"""")&Prop.Number"



'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item(arrMast(i)).PageSheet
'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item("TRM").PageSheet
'SectionNumber = visSectionProp 'Prop 243
'            arrRowName = Array("SA_NazvanieShkafa", "SA_NazvanieMesta")
'            arrRowValue = Array("""Название Шкафа""|""Нумерация элементов идет в пределах одного шкафа""|1|""""|INDEX(0,Prop.SA_NazvanieShkafa.Format)|""""|FALSE|FALSE|1049|0", _
'                            """Название Места""|""Название места расположения или название установки""|0|""""|""""|""""|FALSE|FALSE|1049|0")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber


'
'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item(arrMast(i)).Shapes.Item(arrMast(i))
'SectionNumber = visSectionUser 'User 242
'sSectionName = "User."
'            arrRowName = Array("Shkaf", "Mesto")
'            arrRowValue = Array("ThePage!Prop.SA_NazvanieShkafa|""""", _
'                            "ThePage!Prop.SA_NazvanieMesta|""""")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
'



'SectionNumber = visSectionUser 'User 242
'sSectionName = "User."
'            arrRowName = Array("Name", "Shkaf", "Mesto")
'            arrRowValue = Array("IF(TheDoc!User.SA_ISO,IF(STRSAME(User.Mesto,""""),"""",TheDoc!User.SA_PrefMesto&User.Mesto&IF(Prop.PerenosOboz,CHAR(10),""""))&IF(STRSAME(User.Shkaf,""""),"""",TheDoc!User.SA_PrefShkaf&User.Shkaf&IF(Prop.PerenosOboz,CHAR(10),""""))&TheDoc!User.SA_PrefElement,"""")&Prop.SymName&Prop.Number", _
'                            """""|""""", _
'                            """""|""""")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
'
'SectionNumber = visSectionProp 'Prop 243
'            arrRowName = Array("PerenosOboz")
'            arrRowValue = Array("""Перенос обозн.""|""Переносить обозначение (обозначение в столбец)""|3|""""|FALSE|""69""|NOT(TheDoc!User.SA_ISO)|FALSE|1049|0")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

'Next

'
'
'arrMast = Array("PLCParent", "TRM")
'
'For i = 0 To UBound(arrMast)
'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item(arrMast(i)).Shapes.Item(arrMast(i))
'
'vsoObject.Cells("User.Name").FormulaU = "IF(TheDoc!User.SA_ISO,IF(STRSAME(User.Mesto,""""),"""",TheDoc!User.SA_PrefMesto&User.Mesto&IF(Prop.PerenosOboz,CHAR(10),""""))&IF(STRSAME(User.Shkaf,""""),"""",TheDoc!User.SA_PrefShkaf&User.Shkaf&IF(Prop.PerenosOboz,CHAR(10),""""))&TheDoc!User.SA_PrefElement,"""")&Prop.SymName&Prop.Number"
'
'Next



'arrMast = Array("QS", "SF", "QF", "QA", "KK", "FU3P", "FU", "KL", "KM", "KT", "KV", "UZF", "UZ3P", "UZ", "SA", "SB", "HL", "HA", "EK3P", "EK", "XS", "XS3P", "TV", "UG", "R", "R3P", "RU3P", "RU", "Sensor", "Term", "TermC")
'
'For i = 0 To UBound(arrMast)
'Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item(arrMast(i)).Shapes.Item(arrMast(i))
'
'vsoObject.Cells("Actions.AddDB.Action").FormulaU = "CALLTHIS(""AddDBFrm"")"
'
'Next
'
'On Error Resume Next
'For Each vsoObject In Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Panel").Shapes.Item("Panel").Shapes
'    If vsoObject.name Like "DIN*" Then
'SectionNumber = visSectionProp 'Prop 243
'            arrRowName = Array("NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
'            arrRowValue = Array("""Название из БД""|""Название из БД""|0|""""|""DIN-рейка 35мм оцинкованная 125см""|""60""|FALSE|FALSE|1049|0", _
'                            """Артикул из БД""|""Код заказа из БД""|0|""""|""32054DEK""|""61""|FALSE|FALSE|1049|0", _
'                            """Производитель из БД""|""Производитель из БД""|0|""""|""DEKraft""|""62""|FALSE|FALSE|1049|0", _
'                            """Цена из БД""|""Цена из БД""|0|""""|""193""|""63""|FALSE|FALSE|1049|0", _
'                            """Единица из БД""|""Единица измерения из БД""|0|""""|""шт.""|""64""|FALSE|FALSE|1049|0")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
'    End If
'Next
'For Each vsoObject In Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("PanelMAX").Shapes.Item("PanelMAX").Shapes
'    If vsoObject.name Like "DIN*" Then
'SectionNumber = visSectionProp 'Prop 243
'            arrRowName = Array("NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
'            arrRowValue = Array("""Название из БД""|""Название из БД""|0|""""|""DIN-рейка 35мм оцинкованная 125см""|""60""|FALSE|FALSE|1049|0", _
'                            """Артикул из БД""|""Код заказа из БД""|0|""""|""32054DEK""|""61""|FALSE|FALSE|1049|0", _
'                            """Производитель из БД""|""Производитель из БД""|0|""""|""DEKraft""|""62""|FALSE|FALSE|1049|0", _
'                            """Цена из БД""|""Цена из БД""|0|""""|""193""|""63""|FALSE|FALSE|1049|0", _
'                            """Единица из БД""|""Единица измерения из БД""|0|""""|""шт.""|""64""|FALSE|FALSE|1049|0")
'SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
'    End If
'Next

'For Each vsoObject In Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("PanelMAX").Shapes.Item("PanelMAX").Shapes 'Panel PanelMAX Shkaf ShkafKSRM
'    If vsoObject.name Like "KabKan*" Then
'        With vsoObject
''            .CellsSRC(visSectionControls, 0, visCtlY).FormulaU = "15mm*User.AntiScale"
'            .CellsSRC(visSectionObject, visRowTextXForm, visXFormWidth).FormulaForceU = "GUARD(TEXTWIDTH(TheText))"
'            .CellsSRC(visSectionObject, visRowTextXForm, visXFormHeight).FormulaForceU = "GUARD(TEXTHEIGHT(TheText,TxtWidth))"
'        End With
'    End If
'Next

arrMast = Array("WLS", "QS", "SF", "QF", "QA", "KK", "FU3P", "FU", "KL", "KM", "KT", "KV", "UZ", "UZ3P", "UZF", "SA", "SB", "HL", "HA", "EK", "EK3P", "XS", "XS3P", "TV", "UG", "R", "R3P", "RU", "RU3P") ', "Sensor", "Term", "TermC")

For i = 0 To UBound(arrMast)
Set vsoObject = Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters.Item(arrMast(i)).Shapes.Item(arrMast(i))

SectionNumber = visSectionUser 'User 242
sSectionName = "User."
            arrRowName = Array("Name")
            arrRowValue = Array("IF(TheDoc!User.SA_ISO,IF(STRSAME(User.Shkaf,ThePage!Prop.SA_NazvanieShkafa),"""",IF(STRSAME(User.Mesto,""""),"""",TheDoc!User.SA_PrefMesto&User.Mesto&IF(Prop.PerenosOboz,CHAR(10),""""))&IF(STRSAME(User.Shkaf,""""),"""",TheDoc!User.SA_PrefShkaf&User.Shkaf&IF(Prop.PerenosOboz,CHAR(10),"""")))&TheDoc!User.SA_PrefElement,"""")&Prop.SymName&Prop.Number")
SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

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
    Dim vsoObject As Object
    
'    Set vsoObject = ActivePage.PageSheet
    
'    Set vsoObject = Application.Documents.Item("SAPR_ASU_VID.vss").Masters.Item("Master.34").Shapes("Sheet.5")
    Set vsoObject = ActiveWindow.Selection(1)
'    Set vsoObject = ActivePage.Shapes.ItemFromID(70)
'    Set vsoObject = ActiveDocument.DocumentSheet
    
    strFile = ThisDocument.path & "tempValue.vb"
    
    arrSectionNumber = Array(visSectionUser, visSectionProp, visSectionHyperlink, visSectionConnectionPts, visSectionAction, visSectionControls, visSectionScratch, visSectionTextField, visSectionCharacter, visSectionParagraph, visSectionObject)
    arrRowNumber = Array(visRowXForm1D, visRowXFormOut, visRowLock, visRowMisc, visRowGroup, visRowLine, visRowFill, visRowText, visRowTextXForm, visRowLayerMem, visRowEvent, visRowImage, visRowShapeLayout)
    
    arrSectionNumberTEXT = Array("SectionNumber = visSectionUser 'User 242" & vbNewLine & "sSectionName = ""User.""" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionProp 'Prop 243" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionHyperlink  'Hyperlink 244" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionConnectionPts 'ConnectionPts 7 только именованные" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionAction 'Action 240" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionControls 'Controls 9" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionScratch 'Scratch 6" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionTextField 'Text Field" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionCharacter 'Character" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionParagraph 'Paragraph" & vbNewLine)

    arrRowNumberTEXT = Array("SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowXForm1D '1-D Endpoints" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowXFormOut 'Shape Trannsform" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowLock 'Protection" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowMisc 'Miscellaneous + Glue Info" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowGroup 'Group Propeties" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowLine 'Line Format" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowFill 'Fill Format" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowText 'Text Block Format" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowTextXForm 'Text Transform" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowLayerMem 'Layer Membership" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowEvent 'Events" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowImage 'Image Propeties" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & "SectionNumber = visSectionObject" & vbNewLine & "RowNumber = visRowShapeLayout 'Shape Layout" & vbNewLine, _
       "SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber" & vbNewLine & vbNewLine & vbTab & "")

    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                AddIntoTXTfile strFile, arrRowNumberTEXT(j)
                GetShapeSheetValue vsoObject, strFile, arrSectionNumber(i), arrRowNumber(j)
            Next
            AddIntoTXTfile strFile, arrRowNumberTEXT(j)
        Else
            AddIntoTXTfile strFile, arrSectionNumberTEXT(i)
            GetShapeSheetValue vsoObject, strFile, arrSectionNumber(i)
        End If
    Next
End Sub

Sub SetFont()
    Dim vsoMaster As Visio.Master
    For Each vsoMaster In Application.Documents.Item("SAPR_ASU_VID.vss").Masters
        vsoMaster.Shapes(1).CellsSRC(visSectionCharacter, visRowCharacter, visCharacterFont).Formula = 93
    Next
End Sub

Sub SetIcon()
    Dim vsoMaster As Visio.Master
    For Each vsoMaster In Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters
        On Error Resume Next
        With vsoMaster.Shapes(1)
            .Cells("Actions.Rotate.ButtonFace").FormulaU = "IF(Actions.Rotate.Action,""199"",""198"")" '128 129
            .Cells("Actions.AddReference.ButtonFace").FormulaU = "2651" '1623
            .Cells("Actions.Thumb.ButtonFace").FormulaU = "2871" '256
            .Cells("Actions.Tune.ButtonFace").FormulaU = "1894"
            .Cells("Actions.KlemmyProvoda.ButtonFace").FormulaU = "2601"
            .Cells("Actions.KabeliIzProvodov.ButtonFace").FormulaU = "2642"
            .Cells("Actions.KabeliSrazu.ButtonFace").FormulaU = "1187"
            .Cells("Actions.55.ButtonFace").FormulaU = "203"
            .Cells("Actions.15.ButtonFace").FormulaU = "388"
        End With
    Next
End Sub

Sub SetIconPLC()
    Dim vsoMaster As Visio.Master
    Dim vsoShape As Visio.Shape
    For Each vsoMaster In Application.Documents.Item("SAPR_ASU_CXEMA.vss").Masters
        For i = 1 To 100
            On Error Resume Next
            With vsoMaster.Shapes(1).Shapes.ItemFromID(i)
                .Cells("Actions.Rotate.ButtonFace").FormulaU = "IF(Actions.Rotate.Action,""199"",""198"")" '128 129
                .Cells("Actions.AddReference.ButtonFace").FormulaU = "2651" '1623
                .Cells("Actions.Thumb.ButtonFace").FormulaU = "2871" '256
                .Cells("Actions.Tune.ButtonFace").FormulaU = "1894"
                .Cells("Actions.KlemmyProvoda.ButtonFace").FormulaU = "2601"
                .Cells("Actions.KabeliIzProvodov.ButtonFace").FormulaU = "2642"
                .Cells("Actions.KabeliSrazu.ButtonFace").FormulaU = "1187"
                .Cells("Actions.Duplicate.ButtonFace").FormulaU = "19"
                .Cells("Actions.Glue.ButtonFace").FormulaU = "1649"
                .Cells("Actions.HideName.ButtonFace").FormulaU = "529"
                
            End With
        Next
    Next
End Sub

Sub SetDB()
    Dim vsoObject As Object
    Dim arrRowValue()
    Dim arrRowName()
    Dim arrMast()
    Dim SectionNumber As Long
    Dim RowNumber As Long
    
    Set vsoObject = ActiveWindow.Selection(1)
    
    SectionNumber = visSectionUser 'User 242
    sSectionName = "User."
                arrRowName = Array("KodProizvoditelyaDB", "KodPoziciiDB")
                arrRowValue = Array("|""""", _
                                "|""Код позиции/Код производителя/Код единицы""")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
    
    SectionNumber = visSectionProp 'Prop 243
                arrRowName = Array("NazvanieDB", "ArtikulDB", "ProizvoditelDB", "CenaDB", "EdDB")
                arrRowValue = Array("""Название из БД""|""Название из БД""|0|""""|""""|""60""|FALSE|FALSE|1033|0", _
                                """Артикул из БД""|""Код заказа из БД""|0|""""|""""|""61""|FALSE|FALSE|1033|0", _
                                """Производитель из БД""|""Производитель из БД""|0|""""|""""|""62""|FALSE|FALSE|1033|0", _
                                """Цена из БД""|""Цена из БД""|0|""""|""""|""63""|FALSE|FALSE|1033|0", _
                                """Единица из БД""|""Единица измерения из БД""|0|""""|""""|""64""|FALSE|FALSE|1033|0")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber
    
    SectionNumber = visSectionAction 'Action 240
                arrRowName = Array("AddDB")
                arrRowValue = Array("CALLTHIS(""AddDBFrm"")|""База данных...""|""""|264|""""|0|0|FALSE|FALSE|FALSE")
    SetValueToOneSection vsoObject, arrRowValue, arrRowName, SectionNumber, RowNumber

End Sub

Sub GetAllSSValue()
    Dim arrSectionNumber()
    Dim arrRowNumber()
    Dim arrSectionNumberTEXT()
    Dim arrRowNumberTEXT()
    Dim i As Integer
    Dim j As Integer
    Dim strFile As String
    Dim vsoObject As Object
    
    Set vsoObject = ActivePage.Shapes.ItemFromID(6)
    
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
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "'Case visRowMisc 'Glue Info" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "'arrRowValue = " & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowShapeLayout 'Shape Layout" & vbNewLine & vbTab & vbTab, _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case Else" & vbNewLine & vbTab & vbTab & vbTab & "End Select" & vbNewLine & vbTab & vbTab & "Case Else" & vbNewLine & vbTab & "End Select")

    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                AddIntoTXTfile strFile, arrRowNumberTEXT(j)
                GetShapeSheetValue vsoObject, strFile, arrSectionNumber(i), arrRowNumber(j)
            Next
            AddIntoTXTfile strFile, arrRowNumberTEXT(j)
        Else
            AddIntoTXTfile strFile, arrSectionNumberTEXT(i)
            GetShapeSheetValue vsoObject, strFile, arrSectionNumber(i)
        End If
    Next
End Sub

Private Sub GetShapeSheetValue(vsoObject As Object, strFile As String, ByVal SectionNumber As Long, Optional ByVal RowNumber As Long)
    Dim sRowName As String
    Dim arrRowValue()
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
        ShpRowCount = vsoObject.RowCount(SectionNumber) - 1
        If ShpRowCount = -1 Then Exit Sub
    End If
    ReDim arrRowValue(ShpRowCount)
    ReDim arrRowNameValue(ShpRowCount)
    For j = 0 To ShpRowCount

        If SectionNumber = visSectionConnectionPts And vsoObject.CellExists("Connections.X1", 0) Then 'ConnectionPts Не именованные
            If vsoObject.Cells("Connections.X1").RowName = "" Then NoNameConnectionPts = True
        End If

        If NoNameConnectionPts Then 'ConnectionPts Не именованные
            sRowName = ""
            arrRowNameValue(j) = ""
            For i = 0 To UBarrCellName
               arrRowValue(j) = arrRowValue(j) & vsoObject.Cells(sSectionName & Right(arrCellName(i), 1) & IIf(i > 1, "[" & CStr(j + 1) & "]", CStr(j + 1))).FormulaU & IIf(i = UBarrCellName, "", "|")
            Next
            NoNameConnectionPts = False
            
        ElseIf SectionNumber = visSectionScratch Then    'Scratch
            sRowName = ""
            arrRowNameValue(j) = ""
            For i = 0 To UBarrCellName
               arrRowValue(j) = arrRowValue(j) & vsoObject.Cells(sSectionName & arrCellName(i) & CStr(j + 1)).FormulaU & IIf(i = UBarrCellName, "", "|")
            Next

        ElseIf SectionNumber = visSectionTextField Then    'Text Field
            sRowName = ""
            arrRowNameValue(j) = ""
            For i = 0 To UBarrCellName
               arrRowValue(j) = arrRowValue(j) & vsoObject.Cells(sSectionName & arrCellName(i)).FormulaU & IIf(i = UBarrCellName, "", "|")
            Next
            
        ElseIf SectionNumber = visSectionCharacter Or SectionNumber = visSectionParagraph Or SectionNumber = visSectionObject Then   'Character + Paragraph + SectionObject=Отдельные ячейки без строк
            arrRowNameValue(0) = ""
            For i = 0 To UBarrCellName
               arrRowValue(0) = arrRowValue(0) & vsoObject.Cells(sSectionName & arrCellName(i)).FormulaU & IIf(i = UBarrCellName, "", "|")
            Next
            
        Else 'Все остальные
            sRowName = vsoObject.CellsSRC(SectionNumber, j, 0).RowName 'Replace( , sSectionName, "")
            arrRowNameValue(j) = """" & sRowName & """"
            For i = 0 To UBarrCellName
               arrRowValue(j) = arrRowValue(j) & vsoObject.Cells(sSectionName & sRowName & arrCellName(i)).FormulaU & IIf(i = UBarrCellName, "", "|")
            Next
        End If
        
        arrRowValue(j) = """" & Replace(arrRowValue(j), """", """""") & """"
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
    
    strToFile = strToFile & vbTab & vbTab & vbTab & "arrRowValue = Array("
    UBarrValue = UBound(arrRowValue)
    For i = 0 To UBarrValue
        strToFile = strToFile & IIf(i > 0, vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab, "") & arrRowValue(i) & IIf(i = UBarrValue, ")" & vbNewLine, ", _" & vbNewLine)
    Next

    AddIntoTXTfile strFile, strToFile
End Sub

Sub SetValueToAllSections()
    Dim arrSectionNumber()
    Dim arrRowNumber()
    Dim arrRowValue()
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
                SetValueToOneSection vsoShape, arrRowValue(), arrRowName(), arrSectionNumber(i), arrRowNumber(j)
                arrRowName = Array("")
                arrRowValue = Array("")
            Next
        Else
            GoSub SubSelect
            SetValueToOneSection vsoShape, arrRowValue(), arrRowName(), arrSectionNumber(i)
            arrRowName = Array("")
            arrRowValue = Array("")
        End If
    Next
    
    Exit Sub
    
    
SubSelect:


Return

End Sub

 Sub SetValueToOneSection(vsoObject As Object, arrRowValue(), arrRowName(), ByVal SectionNumber As Long, Optional ByVal RowNumber As Long)
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
        UBarrValue = UBound(arrRowValue)
    End If
    For j = 0 To UBarrValue
        AddSection vsoObject, SectionNumber
        arrCellValue = Split(arrRowValue(j), "|")
        UBarrCellValue = UBound(arrCellValue)
        If Len(arrRowName(0)) <> 0 Then
            sRowName = arrRowName(j)
            AddNamedRow vsoObject, SectionNumber, sRowName
        Else
            If Not (SectionNumber = visSectionScratch Or SectionNumber = visSectionTextField) Then
                AddRow vsoObject, SectionNumber
            End If
        End If
        
        On Error Resume Next
        If SectionNumber = visSectionConnectionPts And Len(sRowName) = 0 Then 'ConnectionPts Не именованные
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoObject.Cells(sSectionName & Right(arrCellName(i), 1) & IIf(i > 1, "[" & CStr(j + 1) & "]", CStr(j + 1))).FormulaU = arrCellValue(i)
                End If
            Next
            
        ElseIf SectionNumber = visSectionScratch Then    'Scratch
            If Not vsoObject.CellExists("Scratch.X" & CStr(j + 1), 0) Then AddRow vsoObject, SectionNumber
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoObject.Cells(sSectionName & arrCellName(i) & CStr(j + 1)).FormulaU = arrCellValue(i)
                End If
            Next

        ElseIf SectionNumber = visSectionTextField Then    'Text Field
            If Not vsoObject.CellExists("Fields.Format" & "[" & CStr(j + 1) & "]", 0) Then AddRow vsoObject, SectionNumber
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoObject.Cells(sSectionName & arrCellName(i) & "[" & CStr(j + 1) & "]").FormulaU = arrCellValue(i)
                End If
            Next

        ElseIf SectionNumber = visSectionCharacter Or SectionNumber = visSectionParagraph Or SectionNumber = visSectionObject Then   'Character + Paragraph + SectionObject=Отдельные ячейки без строк
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoObject.Cells(sSectionName & arrCellName(i)).FormulaU = arrCellValue(i)
                End If
            Next
            
        Else 'Все остальные
            For i = 0 To UBarrCellValue
                If Len(arrCellValue(i)) <> 0 Then
                    vsoObject.Cells(sSectionName & sRowName & arrCellName(i)).FormulaU = arrCellValue(i)
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

Private Sub AddRow(vsoObject As Object, ByVal SectionNumber As Long)
    On Error Resume Next
    vsoObject.AddRow SectionNumber, visRowLast, visTagDefault
End Sub

Private Sub AddNamedRow(vsoObject As Object, ByVal SectionNumber As Long, ByVal sRowName As String)
    On Error Resume Next
    vsoObject.AddNamedRow SectionNumber, sRowName, visTagDefault
End Sub

Private Sub AddSection(vsoObject As Object, ByVal SectionNumber As Long)
    If Not vsoObject.SectionExists(SectionNumber, 0) Then
        vsoObject.AddSection SectionNumber
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


Function AddIntoTXTfile(ByVal FileName As String, ByVal txt As String) As Boolean
    On Error Resume Next: err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set TS = fso.OpenTextFile(FileName, 8, True): TS.Write txt: TS.Close
    Set TS = Nothing: Set fso = Nothing
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