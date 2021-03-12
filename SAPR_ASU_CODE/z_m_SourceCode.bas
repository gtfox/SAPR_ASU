'------------------------------------------------------------------------------------------------------------
' Module    : m_SourceCode - экспорт кода VBA и исходника файла во внешние модули для коммита через ГитХаб
' Author    : Малютин О.С. (Obsidian)
' Date      : 2014
' Purpose   : Модуль хранить процедуры для экспорта кода VBA и исходника файла во внешние модули. Нужен чтобы была возможность коммитить код через ГитХаб
' www.graphicalfiresets.ru, vk.com/aigs_grafis, https://www.youtube.com/channel/UCWlqFPWAKnF61nRxrtk28dw
'------------------------------------------------------------------------------------------------------------

Private docStateDescription As String

Public Sub ExportGitHub()
    SaveSourceCode
End Sub

Public Sub SaveSourceCode()

Dim targetPath As String
Dim doc As Visio.Document
    
    targetPath = "D:\YandexDisk\VISIO\github\SAPR_ASU\SAPR_ASU_CODE\" 'GetCodePath(ThisDocument)
    ExportVBA ThisDocument, targetPath
    'ExportDocState ThisDocument, targetPath
    
    Debug.Print "Сохранено: " & ThisDocument.Name
    MsgBox "Исходный код экспортирован"

End Sub

Public Sub ExportVBA(ByRef doc As Visio.Document, ByVal sDestinationFolder As String)
'Собственно экспорт кода
    Dim oVBComponent As Object
    Dim fullName As String

    For Each oVBComponent In doc.VBProject.VBComponents
        
        If Not oVBComponent.Name = "m_SourceCode" Then
            If oVBComponent.Type = 1 Then
                ' Standard Module
                fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            ElseIf oVBComponent.Type = 2 Then
                ' Class
                fullName = sDestinationFolder & oVBComponent.Name & ".cls"
            ElseIf oVBComponent.Type = 3 Then
                ' Form
                fullName = sDestinationFolder & oVBComponent.Name & ".frm"
            ElseIf oVBComponent.Type = 100 Then
                ' Document
                fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            Else
                ' UNHANDLED/UNKNOWN COMPONENT TYPE
            End If
            
            oVBComponent.Export fullName
            SaveTextToFile oVBComponent.CodeModule.Lines(1, oVBComponent.CodeModule.CountOfLines), fullName
            Debug.Print "Сохранен " & fullName
        End If
        
    Next oVBComponent

End Sub

Private Function GetCodePath(ByRef doc As Visio.Document) As String
'Возвращает путь к папке с исходными кодами
Dim path As String
Dim docNameWODot As String
    
    '---Путь к текущей папке
    path = doc.path
    '---Добавляем название папки с кодами
    path = GetDirPath(path & "_codes")
        
    '---Добавляем путь к папке с кодами ДАННОГО документа
    docNameWODot = Split(doc.Name, ".")(0)
    path = GetDirPath(path & "\" & docNameWODot)
    
    GetCodePath = path & "\"
End Function

Private Function GetDirPath(ByVal path As String) As String
'Возвращает путь к папке с указанным именем, если такой папки нет, предварительно создает ее
    '---Проверяем есть ли такая папка, если нет - создаем
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    GetDirPath = path
End Function

Function SaveTextToFile(ByVal txt, ByVal filename) As Boolean
'функция сохраняет текст txt в кодировке "utf-8" в файл filename
    On Error Resume Next: err.Clear

    With CreateObject("ADODB.Stream")
        .Type = 2: .CharSet = "utf-8": .Open
        .WriteText txt

        Set binaryStream = CreateObject("ADODB.Stream")
        binaryStream.Type = 1: binaryStream.Mode = 3: binaryStream.Open
        .Position = 3: .CopyTo binaryStream        'Skip BOM bytes
        .flush: .Close
        binaryStream.SaveToFile filename, 2
        binaryStream.Close
    End With

    SaveTextToFile = err = 0: DoEvents
End Function


'--------------------Работа с состоянием документа (страницы, фигуры, мастера, стили и т.д.)---------
Private Sub ExportDocState(ByRef doc As Visio.Document, ByVal sDestinationFolder As String)
'Сохраняем состояние документа в текстовый файл
Dim docFullName As String

Dim pg As Visio.Page
Dim shp As Visio.Shape
Dim mstr As Visio.Master
Dim style As Visio.style

'---Очищаем имеющийся текст
    docStateDescription = ""

'---Получаем ссылку на документ и полный путь к нему
    docFullName = sDestinationFolder & Replace(doc.Name, ".", "-") & ".txt"
    
'---Сохраняем состояние всех видов объектов в документе
    '---Документ
    FillSheetData doc.DocumentSheet, docFullName, "Документ"
    '---Страницы
    For Each pg In doc.Pages
        FillSheetData pg.PageSheet, docFullName, pg.Name
        'Фигуры
        For Each shp In pg.Shapes
            FillSheetData shp, docFullName, shp.Name
        Next shp
    Next pg

    '---Мастера
    For Each mstr In doc.Masters
        FillSheetData mstr.PageSheet, docFullName, mstr.Name
        For Each shp In mstr.Shapes
            FillSheetData shp, docFullName, shp.Name
        Next shp
    Next mstr
    
    '---Стили
    Dim ss As Visio.Section
    For Each style In doc.Styles
        FillSheetData style, docFullName, style.Name
    Next style
    
    SaveTextToFile docStateDescription, docFullName
    Debug.Print "Сохранен " & docFullName
End Sub

Public Sub FillSheetData(ByRef sheet As Object, ByVal docFullName As String, ByVal printingName As String)
'Сохраняем в файл по адресу docFullName текущее состояние листа документа, страницы или фигуры (мастера)
Dim shp As Visio.Shape

'---Добавляем название объекта состояния документа
    docStateDescription = docStateDescription & Chr(10) & "=>sheet: " & printingName & Chr(10)
    
'---Экспортируем данные по всем возможнымсекциям
    '---Общие
    FillSectionState sheet, visSectionAction, docFullName
    FillSectionState sheet, visSectionAnnotation, docFullName
    FillSectionState sheet, visSectionCharacter, docFullName
    FillSectionState sheet, visSectionConnectionPts, docFullName
    FillSectionState sheet, visSectionControls, docFullName
    FillSectionState sheet, visSectionFirst, docFullName
    FillSectionState sheet, visSectionFirstComponent, docFullName
    FillSectionState sheet, visSectionHyperlink, docFullName
    FillSectionState sheet, visSectionInval, docFullName
    FillSectionState sheet, visSectionLast, docFullName
    FillSectionState sheet, visSectionLastComponent, docFullName
    FillSectionState sheet, visSectionLayer, docFullName
    FillSectionState sheet, visSectionNone, docFullName
    FillSectionState sheet, visSectionParagraph, docFullName
    FillSectionState sheet, visSectionProp, docFullName
    FillSectionState sheet, visSectionReviewer, docFullName
    FillSectionState sheet, visSectionScratch, docFullName
    FillSectionState sheet, visSectionSmartTag, docFullName
    FillSectionState sheet, visSectionTab, docFullName
    FillSectionState sheet, visSectionTextField, docFullName
    FillSectionState sheet, visSectionUser, docFullName
    '---Секция Объект
    FillSectionObjectState sheet, visRowAlign, docFullName
    FillSectionObjectState sheet, visRowEvent, docFullName
    FillSectionObjectState sheet, visRowDoc, docFullName
    FillSectionObjectState sheet, visRowFill, docFullName
    FillSectionObjectState sheet, visRowForeign, docFullName
    FillSectionObjectState sheet, visRowGroup, docFullName
    FillSectionObjectState sheet, visRowHelpCopyright, docFullName
    FillSectionObjectState sheet, visRowImage, docFullName
    FillSectionObjectState sheet, visRowLayerMem, docFullName
    FillSectionObjectState sheet, visRowLine, docFullName
    FillSectionObjectState sheet, visRowLock, docFullName
    FillSectionObjectState sheet, visRowMisc, docFullName
    FillSectionObjectState sheet, visRowPageLayout, docFullName
    FillSectionObjectState sheet, visRowPage, docFullName
    FillSectionObjectState sheet, visRowPrintProperties, docFullName
    FillSectionObjectState sheet, visRowShapeLayout, docFullName
    FillSectionObjectState sheet, visRowStyle, docFullName
    FillSectionObjectState sheet, visRowTextXForm, docFullName
    FillSectionObjectState sheet, visRowText, docFullName
    FillSectionObjectState sheet, visRowXForm1D, docFullName
    FillSectionObjectState sheet, visRowXFormOut, docFullName
    
    
    'Если указанный объект имеет дочерние фигуры - запускаем процедуру сохранения и для них (актуально только для фигур)
    On Error GoTo EX
    If sheet.Shapes.Count > 0 Then
        For Each shp In pg.Shapes
            FillSheetData shp, docFullName, shp.Name
        Next shp
    End If
    
EX:

End Sub


Private Sub FillSectionState(ByRef shp As Object, ByVal sectID As VisSectionIndices, ByVal docFullName As String)
'Сохраняем в файл по адресу docFullName текущее состояние указанной секции листа документа, страницы или фигуры (мастера)
'ОБЩЕЕ
Dim sect As Visio.Section
Dim rwI As Integer
Dim rw As Visio.Row
Dim cllI As Integer
Dim cll As Visio.Cell
Dim str As String
       
    On Error GoTo EX
       
    If Not TryGetSection(shp, sect, sectID) Then Exit Sub
    
    '---Записываем индекс Секции
    docStateDescription = docStateDescription & "  Section: " & GetSectionName(sectID) & ">>>" & Chr(10)
    
    '---Перебираем все row секции и для каждой из row формируем строку содержащуюю пары Имя-Формула всех ячеек. При условии, что ячейка не пустая
    For rwI = 0 To sect.Count - 1
        Set rw = sect.Row(rwI)
        str = "    "
        For cllI = 0 To rw.Count - 1
            Set cll = rw.Cell(cllI)
            If cll.Formula <> "" Then
                str = str & cll.Name & ": " & cll.Formula & "; "
            End If
        Next cllI
        'Сохраняем строку в файл
        docStateDescription = docStateDescription & str & Chr(10)
    Next rwI
    
Exit Sub
EX:
    If TypeName(shp) = "Style" Then Exit Sub
    Debug.Print "Section ERROR: " & sectID
End Sub

Private Sub FillSectionObjectState(ByRef shp As Object, ByVal rowID As VisRowIndices, ByVal docFullName As String)
'Сохраняем в файл по адресу docFullName текущее состояние листа документа, страницы или фигуры (мастера)
'!!!Для Ячейки ОБЪЕКТ!!!
Dim sect As Visio.Section
Dim rw As Visio.Row
Dim cllI As Integer
Dim cll As Visio.Cell
Dim str As String
    
    On Error GoTo EX
        
    If Not TryGetObjectRow(shp, rw, rowID) Then Exit Sub
    
    '---Записываем индекс Секции
    docStateDescription = docStateDescription & "  Section Object, row: " & GetRowName(rowID) & ">>>" & Chr(10)
    
    '---Перебираем все row секции и для каждой из row формируем строку содержащуюю пары Имя-Формула всех ячеек. При условии, что ячейка не пустая
    If rw.Count > 0 Then
        str = "    "
        For cllI = 0 To rw.Count - 1
            Set cll = rw.Cell(cllI)
            If cll.Formula <> "" Then
                str = str & cll.Name & ": " & cll.Formula & "; "
            End If
        Next cllI
        'Сохраняем строку в файл
        docStateDescription = docStateDescription & str & Chr(10)
    End If
    
Exit Sub
EX:
    If TypeName(shp) = "Style" Then Exit Sub
    Debug.Print "Section Oject ERROR: " & sectID & ", rowID: " & rowID
End Sub

Private Function TryGetSection(ByRef shp As Object, ByRef sect As Visio.Section, ByVal sectID As VisSectionIndices) As Boolean
'Пытаемся получить ссылку на указанную секцию, если это удалось - возвращаем True, иначе False
    On Error GoTo EX
    
    Set sect = shp.Section(sectID)
    
    TryGetSection = True
Exit Function
EX:
    TryGetSection = False
End Function

Private Function TryGetObjectRow(ByRef shp As Object, ByRef rw As Visio.Row, ByVal rowID As VisRowIndices) As Boolean
'Пытаемся получить ссылку на указанную строку секции Object, если это удалось - возвращаем True, иначе False
    On Error GoTo EX
    
    Set rw = shp.Section(visSectionObject).Row(rowID)
    
    TryGetObjectRow = True
Exit Function
EX:
    TryGetObjectRow = False
End Function

Private Function GetSectionName(ByVal sectID As VisSectionIndices) As String
    Select Case sectID
        Case Is = VisSectionIndices.visSectionAction
            GetSectionName = "Action"
        Case Is = VisSectionIndices.visSectionAnnotation
            GetSectionName = "Annotation"
        Case Is = VisSectionIndices.visSectionCharacter
            GetSectionName = "Character"
        Case Is = VisSectionIndices.visSectionConnectionPts
            GetSectionName = "ConnectionPts"
        Case Is = VisSectionIndices.visSectionControls
            GetSectionName = "Controls"
        Case Is = VisSectionIndices.visSectionFirst
            GetSectionName = "First"
        Case Is = VisSectionIndices.visSectionFirstComponent
            GetSectionName = "FirstComponent"
        Case Is = VisSectionIndices.visSectionHyperlink
            GetSectionName = "Hyperlink"
        Case Is = VisSectionIndices.visSectionInval
            GetSectionName = "Inval"
        Case Is = VisSectionIndices.visSectionLast
            GetSectionName = "Last"
        Case Is = VisSectionIndices.visSectionLastComponent
            GetSectionName = "LastComponent"
        Case Is = VisSectionIndices.visSectionLayer
            GetSectionName = "Layer"
        Case Is = VisSectionIndices.visSectionNone
            GetSectionName = "None"
        Case Is = VisSectionIndices.visSectionObject
            GetSectionName = "Object"
        Case Is = VisSectionIndices.visSectionParagraph
            GetSectionName = "Paragraph"
        Case Is = VisSectionIndices.visSectionProp
            GetSectionName = "Prop"
        Case Is = VisSectionIndices.visSectionReviewer
            GetSectionName = "Reviewer"
        Case Is = VisSectionIndices.visSectionScratch
            GetSectionName = "Scratch"
        Case Is = VisSectionIndices.visSectionSmartTag
            GetSectionName = "SmartTag"
        Case Is = VisSectionIndices.visSectionTab
            GetSectionName = "Tab"
        Case Is = VisSectionIndices.visSectionTextField
            GetSectionName = "TextField"
        Case Is = VisSectionIndices.visSectionUser
            GetSectionName = "User"
    End Select
End Function

Private Function GetRowName(ByVal rowID As VisRowIndices) As String
    Select Case rowID
        Case Is = VisRowIndices.visRowAlign
            GetRowName = "visRowAlign"
        Case Is = VisRowIndices.visRowEvent
            GetRowName = "visRowEvent"
        Case Is = VisRowIndices.visRowDoc
            GetRowName = "visRowDoc"
        Case Is = VisRowIndices.visRowFill
            GetRowName = "visRowFill"
        Case Is = VisRowIndices.visRowForeign
            GetRowName = "visRowForeign"
        Case Is = VisRowIndices.visRowGroup
            GetRowName = "visRowGroup"
        Case Is = VisRowIndices.visRowHelpCopyright
            GetRowName = "visRowHelpCopyright"
        Case Is = VisRowIndices.visRowImage
            GetRowName = "visRowImage"
        Case Is = VisRowIndices.visRowLayerMem
            GetRowName = "visRowLayerMem"
        Case Is = VisRowIndices.visRowLine
            GetRowName = "visRowLine"
        Case Is = VisRowIndices.visRowLock
            GetRowName = "visRowLock"
        Case Is = VisRowIndices.visRowMisc
            GetRowName = "visRowMisc"
        Case Is = VisRowIndices.visRowPageLayout
            GetRowName = "visRowPageLayout"
        Case Is = VisRowIndices.visRowPage
            GetRowName = "visRowPage"
        Case Is = VisRowIndices.visRowPrintProperties
            GetRowName = "visRowPrintProperties"
        Case Is = VisRowIndices.visRowShapeLayout
            GetRowName = "visRowShapeLayout"
        Case Is = VisRowIndices.visRowStyle
            GetRowName = "visRowStyle"
        Case Is = VisRowIndices.visRowTextXForm
            GetRowName = "visRowTextXForm"
        Case Is = VisRowIndices.visRowText
            GetRowName = "visRowText"
        Case Is = VisRowIndices.visRowXForm1D
            GetRowName = "visRowXForm1D"
        Case Is = VisRowIndices.visRowXFormOut
            GetRowName = "visRowXFormOut"
    End Select
End Function