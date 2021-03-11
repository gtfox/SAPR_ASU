Private Sub AddRowAndWriteFormulas()
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

Set vsoShape = ActivePage.Shapes.ItemFromID(6)

Section = visSectionObject
Row = visRowMisc





On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

'Имена ячеек

For i = 0 To vsoShape.RowsCellCount(Section, Row)
    strToFile = strToFile & """." & vsoShape.CellsSRC(Section, Row, i).LocalName & """, "
Next


 strToFile = strToFile & vbNewLine & vbNewLine
 
 AddIntoTXTfile strFile, strToFile

'For r = 0 To vsoShape.RowCount(Section) - 1
'    Set vsoCellF = vsoShape.CellsSRC(Section, r, 0)
'
'
'        i = SH2.CellsRowIndex(vsoCellF.Name)
'        For j = 0 To SH2.RowsCellCount(Section, i) ' Перебор ячеек и запись значений в них
'            Debug.Print SH2.CellsSRC(Section, i, j).Name
'
'        Next
'
'Next

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