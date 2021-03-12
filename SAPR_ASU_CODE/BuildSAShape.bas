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



Private Sub GetCellsNames()
    ' Получение имен ячеек
    ' CONSTANT                  VALUE
    '--------------------------------
    ' visSectionAction          240
    ' visSectionConnectionPts   7   только именованные
    ' visSectionControls        9
    ' visSectionHyperlink       244
    ' visSectionProp            243
    ' visSectionSmartTag        247
    ' visSectionUser            242
    '--------------------------------
    
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
    ' visRowXForm1D                 4
    ' visRowXFormOut                1+
    
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

Private Sub GetNamedValue()
    Dim vsoShape As Visio.Shape
    Dim sRowName As String
    Dim arrValue()
    Dim arrCellName()
    Dim sSectionName As String
    Dim i As Integer
    Dim Section As Integer
    
    Dim strFile As String
    Dim strToFile As String

    strFile = ThisDocument.path & "tempValueNamed.vb"
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(11)

    '
    Select Case SectionNumber
        Case visSectionUser 'User            242
            sSectionName = "User."
            arrCellName = Array("", ".Prompt")
        Case visSectionProp 'Prop            243
            sSectionName = "Prop."
            arrCellName = Array(".Label", ".Prompt", ".Type", ".Format", "", ".SortKey", ".Invisible", ".Verify", ".LangID", ".Calendar")
        Case visSectionHyperlink  'Hyperlink       244
            sSectionName = "Hyperlink."
            arrCellName = Array("", ".Address", ".SubAddress", ".ExtraInfo", ".Frame", ".SortKey", ".NewWindow", ".Default", ".Invisible")
        Case visSectionConnectionPts 'ConnectionPts   7   только именованные
            sSectionName = "Connections."
            arrCellName = Array(".X", ".Y", ".A", ".B", ".C", ".D")
        Case visSectionAction 'Action          240
            sSectionName = "Actions."
            arrCellName = Array(".Action", ".Menu", ".TagName", ".ButtonFace", ".SortKey", ".Checked", ".Disabled", ".ReadOnly", ".Invisible", ".BeginGroup", ".FlyoutChild")
        Case visSectionControls 'Controls        9
            sSectionName = "Controls."
            arrCellName = Array("", ".Y", ".XDyn", ".YDyn", ".XCon", ".YCon", ".CanGlue", ".Type")
        Case visSectionScratch 'Scratch         6
            sSectionName = "Scratch."
            arrCellName = Array(".X", ".Y", ".A", ".B", ".C", ".D")   X & i
        Case visSectionTextField 'Text Field
            RowName = ""
            sSectionName = "Fields."
            arrCellName = Array("Format", "Value", "Calendar", "ObjectKind")
        Case visSectionCharacter 'Character
            RowName = ""
            sSectionName = "Char."
            arrCellName = Array("Font", "Size", "FontScale", "Letterspace", "Color", "ColorTrans", "Style", "Case", "Pos", "Strikethru", "DblUnderline", "Overline", "DoubleStrikethrough", "AsianFont", "ComplexScriptFont", "LocalizeFont", "ComplexScriptSize", "LangID")
        Case visSectionParagraph 'Paragraph
            RowName = ""
            sSectionName = "Para."
            arrCellName = Array("IndFirst", "IndLeft", "IndRight", "SpLine", "SpBefore", "SpAfter", "HorzAlign", "Bullet", "BulletStr", "BulletFont", "LocalizeBulletFont", "TextPosAfterBullet", "BulletFontSize", "Flags")
        Case Else
    End Select





    'Значения ячеек
    ReDim arrValue(vsoShape.RowCount(Section) - 1, 1)
    For j = 0 To vsoShape.RowCount(Section) - 1
        sRowName = vsoShape.CellsSRC(Section, j, 0).RowName 'Replace( , sSectionName, "")
        arrValue(j, 0) = sRowName
        For i = 0 To UBound(arrCellName)
           arrValue(j, 1) = arrValue(j, 1) & """" & Replace(vsoShape.Cells(sSectionName & sRowName & arrCellName(i)).FormulaU, """", """""") & """, "
        Next
    Next
    
    For i = 0 To UBound(arrValue)
        AddIntoTXTfile strFile, arrValue(i, 0) & arrValue(i, 1) & vbNewLine & vbNewLine
    Next

End Sub

Private Sub GetExtValue()

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
    
    'visSectionTextField
    'visRowField
    
    'visSectionCharacter
    'visRowCharacter
    
    'visSectionParagraph
    'visRowParagraph
    
    'visSectionTab
    'visRowTab

    Dim Section As Integer
    Dim Row As Integer
    Dim vsoCellF As Visio.Cell, r As Integer, i As Integer, j As Integer, booAddRow As Boolean
    Dim vsoShape As Visio.Shape
    Dim strFile As String
    Dim strToFile As String
    Dim arrCellName()


    strFile = ThisDocument.path & "tempValueExt.vb"
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(6)
    'Shape Trannsform
    'visRowXFormOut
    arrCellName = Array("Width", "Height", "Angle", "PinX", "PinY", "LocPinX", "LocPinY", "FlipX", "FlipY", "ResizeMode")

    'Значения ячеек
    For i = 0 To UBound(arrCellName)
       strToFile = strToFile & """" & Replace(vsoShape.Cells(arrCellName(i)).FormulaU, """", """""") & """, "
    Next

        AddIntoTXTfile strFile, strToFile & vbNewLine & vbNewLine

End Sub

Private Sub SetExtValue()

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
    
    'visSectionTextField
    'visRowField
    
    'visSectionCharacter
    'visRowCharacter
    
    'visSectionParagraph
    'visRowParagraph
    
    'visSectionTab
    'visRowTab

    Dim Section As Integer
    Dim Row As Integer
    Dim vsoCellF As Visio.Cell, r As Integer, i As Integer, j As Integer, booAddRow As Boolean
    Dim vsoShape As Visio.Shape
    Dim strFile As String
    Dim strToFile As String
    Dim arrCellName()


    strFile = ThisDocument.path & "tempValueExt.vb"
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(6)
    
    
    'visSectionObject
    Select Case RowNumber
        Case visRowXForm1D '1-D Endpoints
            arrCellName = Array("BeginX", "BeginY", "EndX", "EndY")
        Case visRowXFormOut 'Shape Trannsform
            arrCellName = Array("Width", "Height", "Angle", "PinX", "PinY", "LocPinX", "LocPinY", "FlipX", "FlipY", "ResizeMode")
        Case visRowLock 'Protection
            arrCellName = Array("LockWidth", "LockHeight", "LockAspect", "LockMoveX", "LockMoveY", "LockRotate", "LockBegin", "LockEnd", "LockDelete", "LockSelect", "LockFormat", "LockCustProp", "LockTextEdit", "LockVtxEdit", "LockCrop", "LockGroup", "LockCalcWH", "LockFromGroupFormat", "LockThemeColors", "LockThemeEffects")
        Case visRowMisc 'Miscellaneous
            arrCellName = Array("NoObjHandles", "NoCtlHandles", "NoAlignBox", "NonPrinting", "LangID", "HideText", "UpdateAlignBox", "DynFeedback", "NoLiveDynamics", "Calendar", "ObjType", "IsDropSource", "Comment", "DropOnPageScale", "LocalizeMerge")
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
        Case visRowMisc 'Glue Info
            arrCellName = Array("BegTrigger", "EndTrigger", "GlueType", "WalkPreference")
        Case visRowShapeLayout 'Shape Layout
            arrCellName = Array("ShapePermeableX", "ShapeFixedCode", "ConLineJumpDirX", "ConLineJumpCode", "ShapePlaceFlip", "ShapePlaceStyle", "ShapePermeableY", "ShapePlowCode", "ConLineJumpDirY", "ConLineJumpStyle", "ConLineRouteExt", "ShapePermeablePlace", "ShapeRouteStyle", "ConFixedCode", "ShapeSplit", "ShapeSplittable")
        Case Else
    End Select

    'Значения ячеек
    For i = 0 To UBound(arrCellName)
       strToFile = strToFile & """" & Replace(vsoShape.Cells(arrCellName(i)).FormulaU, """", """""") & """, "
    Next

        AddIntoTXTfile strFile, strToFile & vbNewLine & vbNewLine

End Sub

Private Function RowNameExists(Section As Integer, strName As String) As Boolean
    ' Проверка на наличие строки в секции (если нет - добавляется)
    
    On Error GoTo err
        SH2.AddNamedRow Section, strName, 0
        RowNameExists = False
    Exit Function
    
err:
        RowNameExists = True
End Function


Function AddIntoTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(filename, 8, True): ts.Write txt: ts.Close
    Set ts = Nothing: Set fso = Nothing
    AddIntoTXTfile = err = 0
End Function