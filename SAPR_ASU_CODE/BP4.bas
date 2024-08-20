'------------------------------------------------------------------------------------------------------------
' Module        : BP4 - Ведомость рабочих чертежей
' Author        : gtfox на основе шейпа от Surrogate::speka2003
' Date          : 2021.03.21
' Description   : Запускается пкм на шейпе "Обновить ВРЧ".
'               : Обновлять ВРЧ надо в последнюю очередь, перед печатью, когда есть все листы с рамками и в больших рамках указано наименование листов.
' Link          : https://visio.getbb.ru/viewtopic.php?p=14130, https://github.com/gtfox/SAPR_ASU, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе шейпа из:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : speka2003 Спецификация
                ' Author    : Surrogate
                ' Date      : 07.11.2012
                ' Purpose   : Спецификация: перенос данных из Excel из Visio и обратно
                '           : Мастер для переноса данных из экселя в визио, для формирования спецификации
                ' Links     : https://visio.getbb.ru/viewtopic.php?f=15&t=234, https://visio.getbb.ru/download/file.php?id=106
                '------------------------------------------------------------------------------------------------------------

Sub FillBP4(shpBP4 As Visio.Shape)
    Dim colRamki As Collection
    Dim vsoPage As Visio.Page
    Dim NazvanieRazdela As String
    Set colRamki = New Collection
    
    If shpBP4.Cells("User.v").Result(0) > 0 Then
        For i = 1 To 30
            shpBP4.Shapes.Item("row" & i).Shapes.Item(i & ".1").text = " "
            shpBP4.Shapes.Item("row" & i).Shapes.Item(i & ".2").text = " "
            shpBP4.Shapes.Item("row" & i).Shapes.Item(i & ".3").text = " "
        Next
    End If

    For Each vsoPage In ActiveDocument.Pages
        On Error GoTo err
        Set shpRamka = vsoPage.Shapes("Рамка")
        If shpRamka.Shapes("FORMA3").Shapes("NaimenovLista").Characters.text <> "" And _
           shpRamka.Cells("user.n").Result(0) = 3 And _
           Right(shpRamka.Shapes("FORMA3").Shapes("Shifr").Cells("fields.value").ResultStr(""), 3) <> ".CO" _
        Then
            colRamki.Add shpRamka
        End If
err:
    Next

    For i = 1 To colRamki.Count
        NazvanieRazdela = colRamki(i).Shapes("FORMA3").Shapes("NaimenovLista").Characters.text
        NachaloRazdela = colRamki(i).Cells("User.NomerLista").Result(0)
        If colRamki.Count >= (i + 1) Then
            KonecRazdela = colRamki(i + 1).Cells("User.NomerLista").Result(0) - 1
            If KonecRazdela - NachaloRazdela <= 0 Then KonecRazdela = 0
        Else
            KonecRazdela = colRamki(i).Cells("User.ChisloListov").Result(0)
        End If
        
        shpBP4.Shapes.Item("row" & i).Shapes.Item(i & ".1").text = NachaloRazdela & IIf(KonecRazdela = 0 Or KonecRazdela = NachaloRazdela, "", "-" & KonecRazdela)
        shpBP4.Shapes.Item("row" & i).Shapes.Item(i & ".2").text = DelSpace(NazvanieRazdela)
    Next
    Application.EventsEnabled = -1
    ThisDocument.InitEvent
    
    MsgBox "ВРЧ обновлена", vbInformation, "САПР-АСУ"
End Sub

Public Function DelSpace(sStroka As String) As String
'Рекурсивное удаление лишних пробелов
    If InStr(sStroka, "  ") > 0 Then
        DelSpace = DelSpace(Replace(sStroka, "  ", " "))
    Else
        DelSpace = sStroka
    End If
End Function


Sub fff()
'Преобразует строки шейпа спецификации в шейп ВРЧ
    Dim shRow As Shape
    Dim shCell As Shape
    Dim strSource() As String
    Set shRow = ActivePage.Shapes.ItemFromID(1)
    
    For i = 1 To 30
        Set shCell = shRow.Shapes.Item("row" & i).Shapes.Item(i & ".1")
        strSource = Split(shCell.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU, "!")
        With shCell
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = strSource(0) & "!Width*15/185"
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = strSource(0) & "!Width*0"
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0"
        End With
        Set shCell = shRow.Shapes.Item("row" & i).Shapes.Item(i & ".2")
        With shCell
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = strSource(0) & "!Width*140/185"
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = strSource(0) & "!Width*15/185"
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0"
            .CellsSRC(visSectionObject, visRowText, visTxtBlkRightMargin).FormulaU = "0 pt"
            .CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
        End With
        Set shCell = shRow.Shapes.Item("row" & i).Shapes.Item(i & ".3")
        With shCell
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = strSource(0) & "!Width*30/185"
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = strSource(0) & "!Width*155/185"
            .CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width * 0"
        End With
        For j = 4 To 9
            shRow.Shapes.Item("row" & i).Shapes.Item(i & "." & j).Delete
        Next
    Next
End Sub



Sub www()
'Сохраняем потеряные формулы из работающей спецификации
    Dim vsoShape As Visio.Shape
    Dim sRowName As String
    Dim arrRowValue(29)
    Dim arrRowNameValue()
    Dim i As Integer
    Dim j As Integer
    Dim UBarrCellName As Integer
    Dim UBarrValue As Integer
    Dim UBarrRowNameValue As Integer
    Dim ShpRowCount As Integer
    Dim strToFile As String
    Dim strFile As String
    
    Dim shRow As Shape
    Dim shCell As Shape
    Dim strSource() As String
    Set shRow = ActivePage.Shapes.ItemFromID(1)
    
    strFile = ThisDocument.path & "tempValue.vb"
    
    For i = 0 To 29
        arrRowValue(i) = shRow.Shapes.Item("row" & i + 1).Cells("height").FormulaU
    Next
    
    UBarrValue = UBound(arrRowValue)
    For i = 0 To UBarrValue
        strToFile = strToFile & """" & arrRowValue(i) & """" & IIf(i = UBarrValue, ")", ", _" & vbNewLine)
    Next

    AddIntoTXTfile strFile, strToFile

End Sub


Sub eee()
'Записывает потерянные формулы в строки шейпа спецификации который станет ВРЧ
    Dim shRow As Shape
    Dim shCell As Shape
    Dim strSource() As String
    Set shRow = ActivePage.Shapes.ItemFromID(1)

www1 = Array("GUARD(MAX(Sheet.23!User.Row_1,Sheet.24!User.Row_1,Sheet.25!User.Row_1))", _
"GUARD(MAX(Sheet.33!User.Row_1,Sheet.34!User.Row_1,Sheet.35!User.Row_1))", _
"GUARD(MAX(Sheet.43!User.Row_1,Sheet.44!User.Row_1,Sheet.45!User.Row_1))", _
"GUARD(MAX(Sheet.53!User.Row_1,Sheet.54!User.Row_1,Sheet.55!User.Row_1))", _
"GUARD(MAX(Sheet.63!User.Row_1,Sheet.64!User.Row_1,Sheet.65!User.Row_1))", _
"GUARD(MAX(Sheet.73!User.Row_1,Sheet.74!User.Row_1,Sheet.75!User.Row_1))", _
"GUARD(MAX(Sheet.83!User.Row_1,Sheet.84!User.Row_1,Sheet.85!User.Row_1))", _
"GUARD(MAX(Sheet.93!User.Row_1,Sheet.94!User.Row_1,Sheet.95!User.Row_1))", _
"GUARD(MAX(Sheet.103!User.Row_1,Sheet.104!User.Row_1,Sheet.105!User.Row_1))", _
"GUARD(MAX(Sheet.113!User.Row_1,Sheet.114!User.Row_1,Sheet.115!User.Row_1))", _
"GUARD(MAX(Sheet.123!User.Row_1,Sheet.124!User.Row_1,Sheet.125!User.Row_1))", _
"GUARD(MAX(Sheet.133!User.Row_1,Sheet.134!User.Row_1,Sheet.135!User.Row_1))", _
"GUARD(MAX(Sheet.143!User.Row_1,Sheet.144!User.Row_1,Sheet.145!User.Row_1))", _
"GUARD(MAX(Sheet.153!User.Row_1,Sheet.154!User.Row_1,Sheet.155!User.Row_1))", _
"GUARD(MAX(Sheet.163!User.Row_1,Sheet.164!User.Row_1,Sheet.165!User.Row_1))", _
"GUARD(MAX(Sheet.173!User.Row_1,Sheet.174!User.Row_1,Sheet.175!User.Row_1))", _
"GUARD(MAX(Sheet.183!User.Row_1,Sheet.184!User.Row_1,Sheet.185!User.Row_1))", _
"GUARD(MAX(Sheet.193!User.Row_1,Sheet.194!User.Row_1,Sheet.195!User.Row_1))", _
"GUARD(MAX(Sheet.203!User.Row_1,Sheet.204!User.Row_1,Sheet.205!User.Row_1))", _
"GUARD(MAX(Sheet.213!User.Row_1,Sheet.214!User.Row_1,Sheet.215!User.Row_1))", _
"GUARD(MAX(Sheet.223!User.Row_1,Sheet.224!User.Row_1,Sheet.225!User.Row_1))", _
"GUARD(MAX(Sheet.233!User.Row_1,Sheet.234!User.Row_1,Sheet.235!User.Row_1))", _
"GUARD(MAX(Sheet.243!User.Row_1,Sheet.244!User.Row_1,Sheet.245!User.Row_1))", _
"GUARD(MAX(Sheet.253!User.Row_1,Sheet.254!User.Row_1,Sheet.255!User.Row_1))")

www2 = Array("GUARD(MAX(Sheet.263!User.Row_1,Sheet.264!User.Row_1,Sheet.265!User.Row_1))", _
"GUARD(MAX(Sheet.273!User.Row_1,Sheet.274!User.Row_1,Sheet.275!User.Row_1))", _
"GUARD(MAX(Sheet.283!User.Row_1,Sheet.284!User.Row_1,Sheet.285!User.Row_1))", _
"GUARD(MAX(Sheet.293!User.Row_1,Sheet.294!User.Row_1,Sheet.295!User.Row_1))", _
"GUARD(MAX(Sheet.303!User.Row_1,Sheet.304!User.Row_1,Sheet.305!User.Row_1))", _
"GUARD(MAX(Sheet.313!User.Row_1,Sheet.314!User.Row_1,Sheet.315!User.Row_1))")

    For i = 1 To 24
        shRow.Shapes.Item("row" & i).Cells("height").FormulaU = www1(i - 1)
    Next
    For i = 1 To 6
        shRow.Shapes.Item("row" & i + 24).Cells("height").FormulaU = www2(i - 1)
    Next
End Sub
