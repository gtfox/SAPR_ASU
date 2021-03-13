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
    

Sub GetAllSSValue()
    Dim arrSectionNumber()
    Dim arrRowNumber()
    Dim arrSectionNumberTEXT()
    Dim arrRowNumberTEXT()
    Dim i As Integer
    Dim j As Integer
    Dim strFile As String
    
    strFile = ThisDocument.path & "tempValue.vb"
    
    arrSectionNumber = Array(visSectionUser, visSectionProp, visSectionHyperlink, visSectionConnectionPts, visSectionAction, visSectionControls, visSectionScratch, visSectionTextField, visSectionCharacter, visSectionParagraph, visSectionObject)
    arrRowNumber = Array(visRowXForm1D, visRowXFormOut, visRowLock, visRowMisc, visRowGroup, visRowLine, visRowFill, visRowText, visRowTextXForm, visRowLayerMem, visRowEvent, visRowImage, visRowShapeLayout)
    
    arrSectionNumberTEXT = Array(vbTab & "Select Case SectionNumber" & vbNewLine & vbTab & vbTab & "Case visSectionUser 'User 242" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionProp 'Prop 243" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionHyperlink  'Hyperlink 244" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionConnectionPts 'ConnectionPts 7 только именованные" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionAction 'Action 240" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionControls 'Controls 9" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionScratch 'Scratch 6" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionTextField 'Text Field" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionCharacter 'Character" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & "Case visSectionParagraph 'Paragraph" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ")

    arrRowNumberTEXT = Array(vbNewLine & vbTab & vbTab & "Case visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & "Select Case RowNumber" & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowXForm1D '1-D Endpoints" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowXFormOut 'Shape Trannsform" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowLock 'Protection" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowMisc 'Miscellaneous + Glue Info" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowGroup 'Group Propeties" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowLine 'Line Format" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowFill 'Fill Format" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowText 'Text Block Format" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowTextXForm 'Text Transform" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowLayerMem 'Layer Membership" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowEvent 'Events" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowImage 'Image Propeties" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "'Case visRowMisc 'Glue Info" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "'arrValue = " & vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case visRowShapeLayout 'Shape Layout" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & vbTab & "Case Else" & vbNewLine & vbTab & vbTab & vbTab & "End Select" & vbNewLine & vbTab & vbTab & "Case Else" & vbNewLine & vbTab & "End Select")
    
    
    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                AddIntoTXTfile strFile, arrRowNumberTEXT(j)
                GetShapeSheetValue arrSectionNumber(i), arrRowNumber(j)
            Next
            AddIntoTXTfile strFile, arrRowNumberTEXT(j)
        Else
            AddIntoTXTfile strFile, arrSectionNumberTEXT(i)
            GetShapeSheetValue arrSectionNumber(i)
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
    
    strFile = ThisDocument.path & "tempValue.vb"
    
    arrSectionNumber = Array(visSectionUser, visSectionProp, visSectionHyperlink, visSectionConnectionPts, visSectionAction, visSectionControls, visSectionScratch, visSectionTextField, visSectionCharacter, visSectionParagraph, visSectionObject)
    arrRowNumber = Array(visRowXForm1D, visRowXFormOut, visRowLock, visRowMisc, visRowGroup, visRowLine, visRowFill, visRowText, visRowTextXForm, visRowLayerMem, visRowEvent, visRowImage, visRowShapeLayout)
    
    arrSectionNumberTEXT = Array(vbTab & vbTab & vbTab & "SectionNumber = visSectionUser 'User 242" & vbNewLine & vbTab & vbTab & vbTab & "sSectionName = ""User.""" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionProp 'Prop 243" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionHyperlink  'Hyperlink 244" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionConnectionPts 'ConnectionPts 7 только именованные" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionAction 'Action 240" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionControls 'Controls 9" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionScratch 'Scratch 6" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionTextField 'Text Field" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionCharacter 'Character" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & "SectionNumber = visSectionParagraph 'Paragraph" & vbNewLine & vbTab & vbTab & vbTab & "arrValue = ")

    arrRowNumberTEXT = Array(vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowXForm1D '1-D Endpoints" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowXFormOut 'Shape Trannsform" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowLock 'Protection" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowMisc 'Miscellaneous + Glue Info" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowGroup 'Group Propeties" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowLine 'Line Format" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowFill 'Fill Format" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowText 'Text Block Format" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowTextXForm 'Text Transform" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowLayerMem 'Layer Membership" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowEvent 'Events" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowImage 'Image Propeties" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "SectionNumber = visSectionObject 'Отдельные ячейки без строк" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "sSectionName = """"" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "RowNumber = visRowShapeLayout 'Shape Layout" & vbNewLine & vbTab & vbTab & vbTab & vbTab & vbTab & "arrValue = ", _
    vbNewLine & vbNewLine & vbTab & vbTab & vbTab & vbTab & "")
    
    
    For i = 0 To UBound(arrSectionNumber)
        If arrSectionNumber(i) = visSectionObject Then
            For j = 0 To UBound(arrRowNumber)
                AddIntoTXTfile strFile, arrRowNumberTEXT(j)
                GetShapeSheetValue arrSectionNumber(i), arrRowNumber(j)
            Next
            AddIntoTXTfile strFile, arrRowNumberTEXT(j)
        Else
            AddIntoTXTfile strFile, arrSectionNumberTEXT(i)
            GetShapeSheetValue arrSectionNumber(i)
        End If
    Next
End Sub

Private Sub GetShapeSheetValue(ByVal SectionNumber As Long, Optional ByVal RowNumber As Long)
    Dim vsoShape As Visio.Shape
    Dim sRowName As String
    Dim arrValue()
    Dim arrRowNameValue()
    Dim arrCellName()
    Dim sSectionName As String
    Dim i As Integer
    Dim j As Integer
    Dim UBarrCellName As Integer
    Dim UBarrValue As Integer
    Dim ShpRowCount As Integer
    Dim strFile As String

    strFile = ThisDocument.path & "tempValue.vb"
    
    Set vsoShape = ActivePage.Shapes.ItemFromID(6)
    
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

    'Значения ячеек
    UBarrCellName = UBound(arrCellName)
    If SectionNumber = visSectionObject Then
        ShpRowCount = 0
    Else
        ShpRowCount = vsoShape.RowCount(SectionNumber) - 1
    End If
    ReDim arrValue(ShpRowCount)
    ReDim arrRowNameValue(ShpRowCount)
    For j = 0 To ShpRowCount
        
        If SectionNumber = visSectionConnectionPts And vsoShape.Cells("Connections.X1").RowName = "" Then 'ConnectionPts Не именованные
            sRowName = j + 1
            arrRowNameValue(j) = """" & sRowName & """"
            For i = 0 To UBarrCellName
               arrValue(j) = arrValue(j) & vsoShape.Cells(sSectionName & Right(arrCellName(i), 1) & sRowName).FormulaU & IIf(i = UBarrCellName, "", ";")
            Next
            
        ElseIf SectionNumber = visSectionScratch Then    'Scratch
            sRowName = j + 1
            arrRowNameValue(j) = """" & sRowName & """"
            For i = 0 To UBarrCellName
               arrValue(j) = arrValue(j) & vsoShape.Cells(sSectionName & arrCellName(i) & sRowName).FormulaU & IIf(i = UBarrCellName, "", ";")
            Next
            
        ElseIf SectionNumber = visSectionTextField Or SectionNumber = visSectionCharacter Or SectionNumber = visSectionParagraph Or SectionNumber = visSectionObject Then   'Text Field + Character + Paragraph + SectionObject=Отдельные ячейки без строк
            arrRowNameValue(0) = """"""
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

    AddIntoTXTfile strFile, "[{"
    UBarrValue = UBound(arrValue)
    For i = 0 To UBarrValue
        AddIntoTXTfile strFile, IIf(i > 0, vbTab & vbTab & vbTab & vbTab & vbTab & vbTab, "") & arrRowNameValue(i) & ", " & arrValue(i) & IIf(i = UBarrValue, "}]", "; _" & vbNewLine)
    Next

End Sub

Private Sub SetShapeSheetValue()

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