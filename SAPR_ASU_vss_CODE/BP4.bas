Attribute VB_Name = "BP4"
'------------------------------------------------------------------------------------------------------------
' Module        : BP4 - Ведомость рабочих чертежей
' Author        : gtfox на основе Surrogate::Vedomost2019.vss
' Date          : 2019.09.22
' Description   : Не требует запуска. Он запускается автоматически при добавлении шейпа ВРЧ. Чтобы обновить ВРЧ удаляем старый шейп и кидаем новый.
'               : Добавлять ВРЧ на первый лист ОД надо в последнюю очередь, когда есть все листы с рамками и в больших рамках указано наименование листов.
' Link          : https://visio.getbb.ru/viewtopic.php?p=14130, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------
                'на основе этого:
                '------------------------------------------------------------------------------------------------------------
                ' Module    : BP4:BP4_corrector Ведомость рабочих чертежей
                ' Author    : Surrogate
                ' Date      : 30.08.2019
                ' Purpose   : Мастер для создания ведомости рабочих чертежей
                ' Links     : https://visio.getbb.ru/viewtopic.php?p=200, https://visio.getbb.ru/download/file.php?id=1087
                '------------------------------------------------------------------------------------------------------------
Option Base 1

Sub BP4_corrector(shpObj As Visio.Shape, pp As Integer)
    Dim isSpec As Boolean
    isSpec = False
    Dim ma() As Integer
    Dim r%, form$
    r = ActiveDocument.Pages.Count
    ReDim ma(r)
    Dim pg As Page, sh As Shape, listing$, wn As Window, n%, pos As Shape, prim As Shape
    listing = "": n = 0
    For i = pp To ActiveDocument.Pages.Count
        Set pg = ActiveDocument.Pages(i)
        'pg.Shapes("Рамка").Cells("fields.value").FormulaU = "0"
        'pg.Shapes("Рамка").Cells("fields.value").FormulaU = "=PAGENUMBER()-1"
        On Error GoTo L1
        If pg.Shapes("Рамка").Cells("prop.type").ResultStr("") <> "" Then
            If InStr(1, pg.Shapes("Рамка").Shapes("FORMA3").Shapes("shifr").Cells("fields.value").ResultStr(""), ".CO") = 0 Then
                listing = listing & ";" & pg.Name
                n = n + 1
                ma(n) = pg.Shapes("Рамка").ID
            End If
        End If
    Next i
    Set pg = ActiveDocument.Pages(pp)
    'Set wn = Application.ActiveWindow.Page.PageSheet.OpenSheetWindow
    'Application.ActiveWindow.Shape.Cells("user.store").FormulaU = Chr(34) & listing & Chr(34)
    ActivePage.PageSheet.Cells("user.store").FormulaU = Chr(34) & listing & Chr(34)
    'wn.Close
    Set sh = shpObj
    For i = 1 To n

        Set prim = sh.Shapes("pos" & i).Shapes(3)
        Set pos = prim.Parent
        pos.Cells("prop.det.format").FormulaForceU = "GUARD(ThePage!User.store)"
        pos.Cells("prop.det.value").FormulaForceU = "INDEX(" & i & " ,Prop.det.Format)"
        'form = "IF(0=0,SETF(GetRef(User.ch)," & Chr(34) & "=Pages[" & Chr(34) & "&Prop.det&" & Chr(34) & "]!sheet." & ma(i) & "!user.ch" & Chr(34) & ")+SETF(GetRef(User.de)," & Chr(34) & "=Pages[" & Chr(34) & "&Prop.det&" & Chr(34) & "]!sheet." & ma(i) & "!user.de" & Chr(34) & ")+SETF(GetRef(User.pn)," & Chr(34) & "=Pages[" & Chr(34) & "&Prop.det&" & Chr(34) & "]!sheet." & ma(i) & "!fields.value" & Chr(34) & "),33)"
        'pos.Cells("user.set").FormulaU = form
        pos.CellsSRC(visSectionAction, 0, visActionAction).FormulaU = "GOTOPAGE(Prop.det)"
        pos.CellsSRC(visSectionAction, 0, visActionMenu).FormulaU = """Перейти на ""&Prop.det"
        
    Next
    On Error GoTo L1
    sh.Shapes("pos" & n).Shapes(4).Cells("user.text").FormulaU = "=IF(User.N-1-thedoc!user.coc-User.C>1,User.C&""-""&User.N-1-thedoc!user.coc,User.C)"
    sh.Cells("prop.n").Formula = n
    MsgBox "Ведомость рабочих чертежей добавлена" & vbCrLf & vbCrLf & "Найдено листов N:" & n, vbInformation
    Exit Sub
L1:
    MsgBox "Нет листов для ВРЧ. N:" & n & vbCrLf & vbCrLf & "На всех листах должны быть рамки" & vbCrLf & "и хотя бы в одной указано наименование листа", vbCritical, "Ошибка"
End Sub
