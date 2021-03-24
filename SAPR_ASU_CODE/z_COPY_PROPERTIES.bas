'------------------------------------------------------------------------------------------------------------
' Module    : CopyProperties_Module Копирование свойств
' Author    : Shishok
' Date      : 10.11.2016
' Purpose   : Копирование свойств одного шейпа/страницы/документа в другие шейпы/страницы/документы
' https://github.com/Shishok/, https://yadi.sk/d/qbpj9WI9d2eqF
'------------------------------------------------------------------------------------------------------------

' Если скопировать нужные процедуры себе в документ, то все нормально будет работать.
' А если нужно использовать код не копируя в свой документ, то необходимо подключать
' этот трафарет через VBE > Tools >References. Собственно, не нужно даже открывать
' вручную из Visio этот трафарет, а просто подключить и все.

' Список основных процедур:

' ШЕЙПЫ: ----------------------------------------------------------------------

' RunCopyPropSelectedShapes
' Копирование пользовательских свойств с выделением шейпов. Именованные строки.

' RunCopyPropSelectedShapesNotName
' Копирование пользовательских свойств с выделением шейпов. Неименованные строки.

' RunCopyPropSelectedShapesExt
' Копирование штатных свойств с выделением шейпов.

' RunCopyPropShapesID
' Копирование пользовательских свойств без выделения шейпов. Именованные строки.
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.

' RunCopyPropShapesIDNotName
' Копирование пользовательских свойств без выделения шейпов. Неименованные строки.
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.

' RunCopyPropShapesIDExt
' Копирование штатных свойств без выделения шейпов.
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.


' СТРАНИЦЫ: -------------------------------------------------------------------

' RunCopyPropPages
' Копирование пользовательских свойств страницы. Именованные строки.
' Возможно копирование из страницы в страницы другого документа.

' RunCopyPropPagesNotName
' Копирование пользовательских свойств страницы. Неименованные строки.
' Возможно копирование из страницы в страницы другого документа.

' RunCopyPropPagesExt
' Копирование штатных свойств страницы.
' Возможно копирование из страницы в страницы другого документа.


' ДОКУМЕНТЫ: ------------------------------------------------------------------

' RunCopyPropDocs
' Копирование пользовательских свойств документа в другой документ/документы. Именованные строки.

' RunCopyPropDocsNotName
' Копирование пользовательских свойств документа в другой документ/документы. Неименованные строки.

' RunCopyPropDocsExt
' Копирование штатных свойств документа в другой документ/документы.

'------------------------------------------------------------------------------------------------------------

Option Explicit

Dim SelObj As Visio.Selection  ' Выделенные шейпы
Dim SH1 As Object
Dim SH2 As Object

Sub RunCopyPropSelectedShapes(Section As Integer, ReplaceValue As Boolean, RemoveRow As Boolean)
' Процедура копирования пользовательских свойств шейпа для именованных строк
' Для выделенных предварительно вручную/программно шейпов
' За один раз можно копировать только одну из секций.
' Секцию Connection_Points можно копировать только если строки в ней именованные.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность заменять/не заменять значения в совпадающих по имени строках.
' Есть возможность удалять отсутствующие в шейпе - источнике строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).
'------------------------------------------------------------------------------------------------------------
' Section       - Integer. Номер или константа секции Shapesheet
' ReplaceValue  - Boolean. При совпадении имен строк заменять(True) содержимое ячеек
' RemoveRow     - Boolean. Удалять(True) отсутствующие в источнике строки

' VisSectionIndices Constants

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

Dim i As Integer

Set SelObj = ActiveWindow.Selection ' Выделенные шейпы
Set SH1 = ActiveWindow.Selection.PrimaryItem ' Шейп - источник

If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
 
For i = 2 To SelObj.Count ' Перебор вторичных шейпов
  Set SH2 = SelObj(i) ' вторичный шейп
    Call AddRowAndWriteFormulas(Section, ReplaceValue) ' Перебор строк(и добавление отсутствующих) и запись формул в ячейки
    If RemoveRow Then Call DeleteRow(Section) ' Удаление строк отсутствующих в шейпе - источнике
Next

End Sub

Sub RunCopyPropSelectedShapesNotName(Section As Integer, RemoveRow As Boolean) ' Начальная процедура. Проверки. Перебор вторичных шейпов
' Процедура копирования пользовательских свойств шейпа для неименованных строк (секции: Scratch и Connection Points)
' Для выделенных предварительно вручную/программно шейпов
' За один раз можно копировать только одну из секций.
' Секция Connection_Points. Если строки в секции неименованные. Впрочем, для именованных Connection Points тоже работает.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность предварительно удалять существующие во вторичном шейпе строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).
'------------------------------------------------------------------------------------------------------------

' VisSectionIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visSectionConnectionPts   7   неименованные
' visSectionScratch         6
'--------------------------------

If Section <> 6 And Section <> 7 Then Exit Sub ' Если секция не Scratch или не Connection Points

Dim i As Integer, r As Integer, c As Integer

Set SelObj = ActiveWindow.Selection ' Выделенные шейпы
Set SH1 = ActiveWindow.Selection.PrimaryItem ' Шейп - источник

If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

For i = 2 To SelObj.Count ' Перебор вторичных шейпов
  Set SH2 = SelObj(i) ' вторичный шейп
  If RemoveRow Then SH2.DeleteSection (Section) ' Если нужно удалить предварительно строки секции
    For r = 0 To SH1.RowCount(Section) - 1 ' Перебор строк секции первичного шейпа
        SH2.AddRow Section, visRowLast, 0 ' Добавление новой строки в секцию вторичного шейпа
        For c = 0 To SH2.RowsCellCount(Section, visRowLast) ' Перебор ячеек строки секции и
            SH2.CellsSRC(Section, visRowLast, c).FormulaForceU = SH1.CellsSRC(Section, r, c).FormulaU ' и запись значений в них
        Next
    Next
Next

End Sub

Sub RunCopyPropSelectedShapesExt(Row As Integer) ' Начальная процедура. Проверки. Перебор вторичных шейпов
' Процедура копирования штатных свойств шейпа
' Для выделенных предварительно вручную/программно шейпов
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).
'------------------------------------------------------------------------------------------------------------

' VisRowIndices Constants

' CONSTANT                      VALUE
'------------------------------------
' visRowAlign                   14
' visRowEvent                   5
' visRowFill                    3
' visRowForeign                 9
' visRowGroup                   22
' visRowHelpCopyright           16
' visRowImage                   21
' visRowLayerMem                6
' visRowLine                    2
' visRowLock                    15
' visRowMisc                    17
' visRowShapeLayout             23
' visRowTextXForm               12
' visRowText                    11
' visRowXForm1D                 4
' visRowXFormOut                1

' FOR VISIO >= 2013 -----------------
' visRow3DRotationProperties    30
' visRowBevelProperties         29
' visRowGradientProperties      26
' visRowOtherEffectProperties   28
' visRowQuickStyleProperties    27
' visRowReplaceBehaviors        32
' visRowThemeProperties         31
'------------------------------------

Dim i As Integer, j As Integer

Set SelObj = ActiveWindow.Selection ' Выделенные шейпы
Set SH1 = ActiveWindow.Selection.PrimaryItem ' Шейп - источник

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

For i = 2 To SelObj.Count ' Перебор вторичных шейпов
  Set SH2 = SelObj(i) ' Вторичный шейп
    For j = 0 To SH2.RowsCellCount(visSectionObject, Row) ' Перебор ячеек и
        SH2.CellsSRC(visSectionObject, Row, j).FormulaForceU = SH1.CellsSRC(visSectionObject, Row, j).FormulaU ' и запись значений в них
    Next
Next

End Sub

Sub RunCopyPropShapesID(Section As Integer, ReplaceValue As Boolean, RemoveRow As Boolean, arrID)  ' Начальная процедура. Проверки. Перебор вторичных шейпов
' Процедура копирования пользовательских свойств шейпа для именованных строк
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.
' За один раз можно копировать только одну из секций.
' Секцию Connection_Points можно копировать только если строки в ней именованные.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность заменять/не заменять значения в совпадающих по имени строках.
' Есть возможность удалять отсутствующие в шейпе - источнике строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisSectionIndices Constants

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

If Not IsArray(arrID) Then Exit Sub ' если не массив
If UBound(arrID) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer
On Error Resume Next ' Пропуск ошибки на случай ошибок

With ActivePage.Shapes
    Set SH1 = .ItemFromID(arrID(0)) ' Шейп - источник
    If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
    For i = 1 To UBound(arrID) ' Перебор элементов массива, начиная с 2 элемента
      Set SH2 = .ItemFromID(arrID(i)) ' Вторичный шейп
        Call AddRowAndWriteFormulas(Section, ReplaceValue) ' Перебор строк(и добавление отсутствующих) и запись формул в ячейки
        If RemoveRow Then Call DeleteRow(Section) ' Удаление строк отсутствующих в шейпе - источнике
    Next
End With

End Sub

Sub RunCopyPropShapesIDNotName(Section As Integer, RemoveRow As Boolean, arrID) ' Начальная процедура. Проверки. Перебор вторичных шейпов
' Процедура копирования пользовательских свойств шейпа для неименованных строк
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность предварительно удалять присутствующие во вторичном шейпе строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisSectionIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visSectionConnectionPts   7   неименованные
' visSectionScratch         6
'--------------------------------

If Section <> 6 And Section <> 7 Then Exit Sub ' Если секция не Scratch или не Connection Points
If Not IsArray(arrID) Then Exit Sub ' если не массив
If UBound(arrID) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer, r As Integer, c As Integer

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

With ActivePage.Shapes
  Set SH1 = .ItemFromID(arrID(0)) ' Шейп - источник
  If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
    For i = 1 To UBound(arrID) ' Перебор вторичных шейпов
      Set SH2 = .ItemFromID(arrID(i)) ' вторичный шейп
      If RemoveRow Then SH2.DeleteSection (Section) ' Если нужно удалить предварительно строки секции
        For r = 0 To SH1.RowCount(Section) - 1 ' Перебор строк секции первичного шейпа
            SH2.AddRow Section, visRowLast, 0 ' Добавление новой строки в секцию вторичного шейпа
            For c = 0 To SH2.RowsCellCount(Section, visRowLast) ' Перебор ячеек строки секции и
                SH2.CellsSRC(Section, visRowLast, c).FormulaForceU = SH1.CellsSRC(Section, r, c).FormulaU ' и запись значений в них
            Next
        Next
    Next
End With

End Sub

Sub RunCopyPropShapesIDExt(Row As Integer, arrID) ' Начальная процедура. Проверки. Перебор вторичных шейпов
' Процедура копирования штатных свойств шейпа
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.
' За один раз можно копировать только одну из секций.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisRowIndices Constants

' CONSTANT                      VALUE
'------------------------------------
' visRowAlign                   14
' visRowEvent                   5
' visRowFill                    3
' visRowForeign                 9
' visRowGroup                   22
' visRowHelpCopyright           16
' visRowImage                   21
' visRowLayerMem                6
' visRowLine                    2
' visRowLock                    15
' visRowMisc                    17
' visRowShapeLayout             23
' visRowTextXForm               12
' visRowText                    11
' visRowXForm1D                 4
' visRowXFormOut                1

' FOR VISIO >= 2013 -----------------
' visRow3DRotationProperties    30
' visRowBevelProperties         29
' visRowGradientProperties      26
' visRowOtherEffectProperties   28
' visRowQuickStyleProperties    27
' visRowReplaceBehaviors        32
' visRowThemeProperties         31
'------------------------------------

Dim i As Integer, j As Integer

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

With ActivePage.Shapes
  Set SH1 = .ItemFromID(arrID(0)) ' Шейп - источник
    For i = 1 To UBound(arrID) ' Перебор вторичных шейпов
      Set SH2 = .ItemFromID(arrID(i)) ' Вторичный шейп
        For j = 0 To SH2.RowsCellCount(visSectionObject, Row) ' Перебор ячеек и
            SH2.CellsSRC(visSectionObject, Row, j).FormulaForceU = SH1.CellsSRC(visSectionObject, Row, j).FormulaU ' и запись значений в них
        Next
    Next
End With

End Sub

Sub RunCopyPropPages(Section As Integer, ReplaceValue As Boolean, RemoveRow As Boolean, arrPageIndex, arrDocIndex)  ' Начальная процедура. Проверки. Перебор вторичных страниц
' Процедура копирования пользовательских свойств страницы для именованных строк
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность заменять/не заменять значения в совпадающих по имени строках.
' Есть возможность удалять отсутствующие в шейпе - источнике строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisSectionIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visSectionAction          240
' visSectionHyperlink       244
' visSectionProp            243
' visSectionSmartTag        247
' visSectionUser            242
'--------------------------------

If Not IsArray(arrPageIndex) Then Exit Sub ' если не массив
If UBound(arrPageIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов
If Not IsArray(arrDocIndex) Then Exit Sub ' если не массив
If UBound(arrDocIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer
On Error Resume Next ' Пропуск ошибки на случай ошибок (красиво выразился!)

With Application.Documents
    Set SH1 = .Item(arrDocIndex(0)).Pages.Item(arrPageIndex(0)).PageSheet  ' Страница - источник
    If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
    For i = 1 To UBound(arrPageIndex) ' Перебор элементов массива, начиная с 2 элемента
      Set SH2 = .Item(arrDocIndex(1)).Pages.Item(arrPageIndex(i)).PageSheet ' Вторичная страница
        Call AddRowAndWriteFormulas(Section, ReplaceValue) ' Перебор строк(и добавление отсутствующих) и запись формул в ячейки
        If RemoveRow Then Call DeleteRow(Section) ' Удаление строк отсутствующих в странице - источнике
    Next
End With

End Sub

Sub RunCopyPropPagesNotName(Section As Integer, RemoveRow As Boolean, arrPageIndex, arrDocIndex) ' Начальная процедура. Проверки. Перебор вторичных страниц
' Процедура копирования пользовательских свойств страницы для неименованных строк
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность предварительно удалять присутствующие во вторичном шейпе строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisSectionIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visSectionScratch         6
' visSectionLayer           241
'--------------------------------

If Section <> 6 And Section <> 241 Then Exit Sub  ' Если секция не Scratch или не Connection Points
If Not IsArray(arrPageIndex) Then Exit Sub ' если не массив
If UBound(arrPageIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов
If Not IsArray(arrDocIndex) Then Exit Sub ' если не массив
If UBound(arrDocIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer, r As Integer, c As Integer

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

With Application.Documents
  Set SH1 = .Item(arrDocIndex(0)).Pages.Item(arrPageIndex(0)).PageSheet ' Страница - источник
  If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
    For i = 1 To UBound(arrPageIndex) ' Перебор вторичных страниц
      Set SH2 = .Item(arrDocIndex(1)).Pages.Item(arrPageIndex(i)).PageSheet ' Вторичная страница
      If RemoveRow Then SH2.DeleteSection (Section) ' Если нужно удалить предварительно строки секции
        For r = 0 To SH1.RowCount(Section) - 1 ' Перебор строк секции первичной страницы
            SH2.AddRow Section, visRowLast, 0 ' Добавление новой строки в секцию вторичной страницы
            For c = 0 To SH2.RowsCellCount(Section, visRowLast) ' Перебор ячеек строки секции и
                SH2.CellsSRC(Section, visRowLast, c).FormulaForceU = SH1.CellsSRC(Section, r, c).FormulaU ' и запись значений в них
            Next
        Next
    Next
End With

End Sub

Sub RunCopyPropPagesExt(Row As Integer, arrPageIndex, arrDocIndex) ' Начальная процедура. Проверки. Перебор вторичных страниц
' Процедура копирования штатных свойств страницы.
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisRowIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visRowHelpCopyright       16
' visRowPageLayout          24
' visRowPage                10
' visRowPrintProperties     25
' visRowRulerGrid           18
'--------------------------------

If Not IsArray(arrPageIndex) Then Exit Sub ' если не массив
If UBound(arrPageIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов
If Not IsArray(arrDocIndex) Then Exit Sub ' если не массив
If UBound(arrDocIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer, j As Integer

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

With Application.Documents
  Set SH1 = .Item(arrDocIndex(0)).Pages.Item(arrPageIndex(0)).PageSheet ' Страница - источник
    For i = 1 To UBound(arrPageIndex) ' Перебор вторичных страниц
      Set SH2 = .Item(arrDocIndex(1)).Pages.Item(arrPageIndex(i)).PageSheet ' Вторичная страница
        For j = 0 To SH2.RowsCellCount(visSectionObject, Row) ' Перебор ячеек и
            SH2.CellsSRC(visSectionObject, Row, j).FormulaForceU = SH1.CellsSRC(visSectionObject, Row, j).FormulaU ' и запись значений в них
        Next
    Next
End With

End Sub

Sub RunCopyPropDocs(Section As Integer, ReplaceValue As Boolean, RemoveRow As Boolean, arrDocIndex)  ' Начальная процедура. Проверки. Перебор вторичных документов
' Процедура копирования пользовательских свойств документа для именованных строк
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность заменять/не заменять значения в совпадающих по имени строках.
' Есть возможность удалять отсутствующие в шейпе - источнике строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisSectionIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visSectionHyperlink       244
' visSectionProp            243
' visSectionUser            242
'--------------------------------

If Not IsArray(arrDocIndex) Then Exit Sub
If UBound(arrDocIndex) < 1 Then Exit Sub

Dim i As Integer
On Error Resume Next ' Пропуск ошибки на случай ошибок

With Application.Documents
    Set SH1 = .Item(arrDocIndex(0)).DocumentSheet ' Документ - источник
    If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
    For i = 1 To UBound(arrDocIndex) ' Перебор элементов массива, начиная с 2 элемента
      Set SH2 = .Item(arrDocIndex(i)).DocumentSheet ' Вторичный документ
        Call AddRowAndWriteFormulas(Section, ReplaceValue) ' Перебор строк(и добавление отсутствующих) и запись формул в ячейки
        If RemoveRow Then Call DeleteRow(Section) ' Удаление строк отсутствующих в документе - источнике
    Next
End With

End Sub

Sub RunCopyPropDocsNotName(Section As Integer, RemoveRow As Boolean, arrDocIndex)  ' Начальная процедура. Проверки. Перебор вторичных документов
' Процедура копирования пользовательских свойств документа для неименованных строк
' За один раз можно копировать только одну из секций.
' Копируются все без исключения строки. Выбирать нельзя.
' Есть возможность предварительно удалять присутствующие во вторичном документе строки.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisSectionIndices Constants

' CONSTANT                  VALUE
'--------------------------------
' visSectionScratch         6
'--------------------------------

If Section <> 6 Then Exit Sub ' Если секция не Scratch или не Connection Points
If Not IsArray(arrDocIndex) Then Exit Sub ' если не массив
If UBound(arrDocIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer, r As Integer, c As Integer
On Error Resume Next ' Пропуск ошибки на случай ошибок

With Application.Documents
  Set SH1 = .Item(arrDocIndex(0)).DocumentSheet ' Документ - источник
  If Not SH1.SectionExists(Section, 0) Then Exit Sub ' Если источник не содержит секции
    For i = 1 To UBound(arrDocIndex) ' Перебор вторичных документов
      Set SH2 = .Item(arrDocIndex(i)).DocumentSheet ' Вторичный документ
      If RemoveRow Then SH2.DeleteSection (Section) ' Если нужно удалить предварительно строки секции
        For r = 0 To SH1.RowCount(Section) - 1 ' Перебор строк секции первичного документа
            SH2.AddRow Section, visRowLast, 0 ' Добавление новой строки в секцию вторичного документа
            For c = 0 To SH2.RowsCellCount(Section, visRowLast) ' Перебор ячеек строки секции и
                SH2.CellsSRC(Section, visRowLast, c).FormulaForceU = SH1.CellsSRC(Section, r, c).FormulaU ' и запись значений в них
            Next
        Next
    Next
End With

End Sub

Sub RunCopyPropDocsExt(Row As Integer, arrDocIndex) ' Начальная процедура. Проверки. Перебор вторичных документов
' Процедура копирования штатных свойств документа
' За один раз можно копировать только одну из секций.
' При наличии перекрестных ссылок в ячейках, возможны пустые(без формул) ячейки. Надо повторить процедуру(возможно несколько раз).

' VisRowIndices Constants

' CONSTANT        VALUE
'--------------------------------
' visRowDoc       20
'--------------------------------

If Not IsArray(arrDocIndex) Then Exit Sub ' если не массив
If UBound(arrDocIndex) < 1 Then Exit Sub ' если в массиве меньше 2 элементов

Dim i As Integer, j As Integer

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку

With Application.Documents
  Set SH1 = .Item(arrDocIndex(0)).DocumentSheet ' Документ - источник
    For i = 1 To UBound(arrDocIndex) ' Перебор вторичных страниц
      Set SH2 = .Item(arrDocIndex(i)).DocumentSheet ' Вторичный документ
        For j = 0 To SH2.RowsCellCount(visSectionObject, Row) ' Перебор ячеек и
            SH2.CellsSRC(visSectionObject, Row, j).FormulaForceU = SH1.CellsSRC(visSectionObject, Row, j).FormulaU ' и запись значений в них
        Next
    Next
End With

End Sub

Private Sub AddRowAndWriteFormulas(Section As Integer, ReplaceValue As Boolean)
' Перебор строк(и добавление отсутствующих) и запись формул в ячейки

Dim vsoCellF As Visio.Cell, r As Integer, i As Integer, j As Integer, booAddRow As Boolean

On Error Resume Next ' Пропуск ошибки на случай ссылки на несуществующую ячейку
            
For r = 0 To SH1.RowCount(Section) - 1
    Set vsoCellF = SH1.CellsSRC(Section, r, 0)
    booAddRow = RowNameExists(Section, vsoCellF.RowName)
    
    If Not (booAddRow And Not ReplaceValue) Then
        i = SH2.CellsRowIndex(vsoCellF.Name)
        For j = 0 To SH2.RowsCellCount(Section, i) ' Перебор ячеек и запись значений в них
            SH2.CellsSRC(Section, i, j).FormulaForceU = SH1.CellsSRC(Section, SH1.CellsRowIndex(vsoCellF.Name), j).FormulaU
        Next
    End If
Next

End Sub

Private Sub DeleteRow(Section As Integer)
' Удаление строк отсутствующих в шейпе - источнике

Dim i As Integer, j As Integer, booDelRow As Boolean

For i = SH2.RowCount(Section) - 1 To 0 Step -1
    booDelRow = True
    For j = 0 To SH1.RowCount(Section) - 1
        If SH2.CellsSRC(Section, i, 0).RowName = SH1.CellsSRC(Section, j, 0).RowName Then booDelRow = False
    Next
    If booDelRow Then SH2.DeleteRow Section, i
Next
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


































































'===========================================Примеы==============================================




'-------------------------------------Документы---------------------------------------------


' Проверки на существование документов осуществляет пользователь
' В следующих процедурах переменная arrDocIndex - это:
' Массив индексов документов. Определение индекса выполняет пользователь, вручную или программно
' Первый элемент массива - обязательно документ - источник
' В массиве должно быть не менее двух элементов
' Массив должен начинаться с нуля

Sub Test_RunCopyPropDocs()
' Копирование пользовательских свойств документа в другой документ/документы. Именованные строки.

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim ReplaceValue As Boolean ' При совпадении имен строк заменять(True) содержимое ячеек
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки
Dim arrDocIndex ' Массив индексов документов

Section = visSectionProp ' 243 Номер или константа
ReplaceValue = True
RemoveRow = True
arrDocIndex = Array(1, 4) ' Пример массива документов. Из документа 1 в документ 4

'Application.UndoEnabled = False ' Отключаем запись операций отмен (необязательно)
    Call RunCopyPropDocs(Section, ReplaceValue, RemoveRow, arrDocIndex)  ' Копируем секцию Shape Data
'Application.UndoEnabled = True ' Включаем запись операций отмен (необязательно)

End Sub

Sub Test_RunCopyPropDocsNotName()
' Копирование пользовательских свойств документа в другой документ/документы. Неименованные строки.

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки
Dim arrDocIndex ' Массив индексов документов

Section = visSectionScratch ' 6 Номер или константа
RemoveRow = True
arrDocIndex = Array(1, 4) ' Пример массива документов. Из документа 1 в документ 4

Call RunCopyPropDocsNotName(Section, RemoveRow, arrDocIndex)  ' Копируем секцию Scratch

End Sub

Sub Test_RunCopyPropDocsExt()
' Копирование штатных свойств документа в другой документ/документы.

Dim Row As Integer ' Номер или VisRowIndices Constants
Dim arrDocIndex ' Массив индексов документов.

Row = visRowDoc ' 20  Секция Document properties
arrDocIndex = Array(1, 4) ' Пример массива документов. Из документа 1 в документ 4

Call RunCopyPropDocsExt(Row, arrDocIndex)   ' Копируем секцию Document Properties

End Sub


Sub PrintDocNameIndex() ' Вспомогательная процедура
' Просмотр имен/индексов открытых документов
Dim i As Integer
    With Application.Documents
        For i = 1 To .Count
            Debug.Print .Item(i).Name & " - " & .Item(i).Index
        Next
    End With
End Sub




'-------------------------------------Страницы---------------------------------------------



' Проверки на существование страниц и документов осуществляет пользователь
' В следующих процедурах переменные arrPageIndex и arrDocIndex - это:
' Массивы индексов страниц и документов. Определение индекса выполняет пользователь, вручную или программно
' Первый элемент массивов - обязательно страница - источник и документ - источник
' В массивах должно быть не менее двух элементов
' Массивы должны начинаться с нуля

Sub Test_RunCopyPropPages()
' Копирование пользовательских свойств страницы. Именованные строки.
' Возможно копирование из страницы в страницы другого документа.

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim ReplaceValue As Boolean ' При совпадении имен строк заменять(True) содержимое ячеек
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки
Dim arrPageIndex ' Массив индексов страниц
Dim arrDocIndex ' Массив индексов документов

Section = visSectionProp ' Константа, Секция Shape Data
ReplaceValue = True
RemoveRow = True
arrPageIndex = Array(1, 2) ' Пример массива страниц

' Если надо копировать свойства на страницу в другом документе,
' то вторым элементом массива должен быть индекс этого документа
arrDocIndex = Array(1, 1) ' Пример массива документов

'Application.UndoEnabled = False ' Отключаем запись операций отмен (необязательно)
    Call RunCopyPropPages(Section, ReplaceValue, RemoveRow, arrPageIndex, arrDocIndex) ' Копируем секцию Shape Data
'Application.UndoEnabled = True ' Включаем запись операций отмен (необязательно)

End Sub

Sub Test_RunCopyPropPagesNotName()
' Копирование пользовательских свойств страницы. Неименованные строки.
' Возможно копирование из страницы в страницы другого документа.

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки
Dim arrPageIndex ' Массив индексов страниц
Dim arrDocIndex ' Массив индексов документов

Section = visSectionScratch ' Константа, Секция Scratch
RemoveRow = True
' Из страницы 1 документа 1 в страницу 1 и 2 документа 4
arrPageIndex = Array(1, 1, 2) ' Пример массива страниц
arrDocIndex = Array(1, 4) ' Пример массива документов

Call RunCopyPropPagesNotName(Section, RemoveRow, arrPageIndex, arrDocIndex)  ' Копируем секцию Scratch

End Sub

Sub Test_RunCopyPropPagesExt()
' Копирование штатных свойств страницы.
' Возможно копирование из страницы в страницы другого документа.

Dim Row As Integer ' Номер или VisRowIndices Constants
Dim arrPageIndex ' Массив индексов страниц
Dim arrDocIndex ' Массив индексов документов

Row = visRowPage ' секция Page Properties
' Из страницы 1 документа 1 в страницу 1 и 2 документа 4
arrPageIndex = Array(1, 1, 2) ' Пример массива страниц
arrDocIndex = Array(1, 4) ' Пример массива документов

Call RunCopyPropPagesExt(Row, arrPageIndex, arrDocIndex) ' Копируем секцию Page Properties

End Sub


Sub PrintPageNameIndex() ' Вспомогательная процедура
' Просмотр имен/индексов страниц в активном документе
Dim i As Integer
    With ActiveDocument
        For i = 1 To .Pages.Count
            Debug.Print .Pages.Item(i).Name & " - " & .Pages.Item(i).Index
        Next
    End With
End Sub



'-------------------------------------Шейпы---------------------------------------------



' Проверки на существование шейпов осуществляет пользователь
' В следующих процедурах переменная arrID - это:
' Массив ID шейпов. Определение ID выполняет пользователь, вручную или программно
' Первый элемент массива - обязательно шейп - источник
' В массиве должно быть не менее двух элементов
' Любой элемент массива может быть одним из шейпов сгруппированной фигуры (субшейп)
' Массив должен начинаться с нуля

Sub Test_RunCopyPropSelectedShapes()
' Копирование пользовательских свойств с выделением шейпов. Именованные строки.
' Должно быть предварительно выделенно не менее 2 шейпов

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim ReplaceValue As Boolean ' При совпадении имен строк заменять(True) содержимое ячеек
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки

Section = visSectionProp ' Константа, Секция Shape Data
ReplaceValue = True
RemoveRow = True

'Application.UndoEnabled = False ' Отключаем запись операций отмен (необязательно)
    Call RunCopyPropSelectedShapes(Section, ReplaceValue, RemoveRow) ' Копируем секцию Shape Data
'Application.UndoEnabled = True ' Включаем запись операций отмен (необязательно)
End Sub

Sub Test_RunCopyPropSelectedShapesNotName()
' Копирование пользовательских свойств с выделением шейпов. Неименованные строки.
' Должно быть предварительно выделенно не менее 2 шейпов

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim RemoveRow As Boolean ' Удалять(True) предварительно существующие в источнике строки

Section = visSectionScratch ' Номер или константа
RemoveRow = True

Call RunCopyPropSelectedShapesNotName(Section, RemoveRow) ' Копируем секцию Scratch

End Sub

Sub Test_RunCopyPropSelectedShapesExt()
' Копирование штатных свойств с выделением шейпов.
' Должно быть предварительно выделенно не менее 2 шейпов

Dim Row As Integer ' Номер или VisRowIndices Constants

Row = visRowMisc ' Номер или константа. Секция Miscellanious

Call RunCopyPropSelectedShapesExt(Row)

End Sub

Sub Test_RunCopyPropShapesID()
' Копирование пользовательских свойств без выделения шейпов. Именованные строки.
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim ReplaceValue As Boolean ' При совпадении имен строк заменять(True) содержимое ячеек
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки
Dim arrID ' Массив ID шейпов

arrID = Array(10, 32, 36, 102, 358) ' Пример массива шейпов
Section = visSectionProp ' Константа, Секция Shape Data
ReplaceValue = True
RemoveRow = True

Call RunCopyPropShapesID(Section, ReplaceValue, RemoveRow, arrID) ' Копируем секцию Shape Data

End Sub

Sub Test_RunCopyPropShapesIDNotName()
' Копирование пользовательских свойств без выделения шейпов. Неименованные строки.
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.

Dim Section As Integer ' Номер секции или VisSectionIndices Constants
Dim RemoveRow As Boolean ' Удалять(True) отсутствующие в источнике строки
Dim arrID ' Массив ID шейпов

arrID = Array(10, 32, 36, 102, 358) ' Пример массива шейпов
Section = visSectionConnectionPts ' Константа, Секция Connection Points
RemoveRow = True

Call RunCopyPropSelectedShapesNotName(Section, RemoveRow) ' Копируем секцию Connection Points

End Sub

Sub Test_RunCopyPropShapesIDExt()
' Копирование штатных свойств без выделения шейпов.
' Возможно копирование из шейпа/субшейпа в шейпы/субшейпы.

Dim Row As Integer ' Номер или VisRowIndices Constants
Dim arrID ' Массив ID шейпов

arrID = Array(10, 32, 36, 102, 358) ' Пример массива шейпов
Row = visRowLock ' секция Protection

Call RunCopyPropShapesIDExt(Row, arrID) ' Копируем секцию Protection

End Sub







