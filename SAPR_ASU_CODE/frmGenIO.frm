Option Explicit

'------------------------------------------------------------------------------------------------------------
' Module        : frmGenIO - Форма задания количества входов для генерации вне модуля
' Author        : gtfox
' Date          : 2020.09.14
' Description   : Формируется колонка входов с автонумерацией
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Dim shpIO As Visio.Shape 'шейп из модуля PLC

Sub Run(vsoShape As Visio.Shape) 'Приняли шейп из модуля PLC
    Set shpIO = vsoShape 'И определили его в форме frmGenIO
    frmGenIO.Show
End Sub

Private Sub CommandButton1_Click()
    gen
    Unload Me
End Sub

Private Sub gen()
    Dim NIO As Integer
    NIO = TextBox1.Text
    Call GenIOPLC(shpIO, NIO)
End Sub

Private Sub btnClose_Click() ' выгрузка формы
    Unload Me
End Sub

'Private Sub TextBox1_AfterUpdate() 'крашится visio
'    gen
'    DoEvents
'    Unload Me
'End Sub
