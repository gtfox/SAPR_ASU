VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Dim WithEvents vsoPagesEvent As Visio.Pages 'события
Attribute vsoPagesEvent.VB_VarHelpID = -1
Dim WithEvents vsoWindowEvent As Visio.Window 'мышь
Attribute vsoWindowEvent.VB_VarHelpID = -1

Dim vsoShapePaste As Visio.Shape
Dim Click As Boolean

'Перед удалением шейпа чистим что-либо
Private Sub vsoPagesEvent_BeforeShapeDelete(ByVal Shape As IVShape)

    If Shape.CellExists("User.Type", 0) Then   'Если в шейпе есть тип, то он элемент SAPR ASU
        Select Case Shape.Cells("User.Type").Result(0) 'В зависимости от типа выбираем способ удаления
            Case typeNO, typeNC 'Контакт реле NO,NC (дочерний)
                DeleteChild Shape
            Case typeCoil  'Катушка реле KL (родительский)
                DeleteParent Shape
            Case typeWireLinkR  'Разрыв провода (дочерний)
                DeleteChildWireLink Shape
            Case typeWireLinkS 'Разрыв провода (родительский)
                DeleteParentWireLink Shape
        End Select
    End If

End Sub

'Активация событий при старте
Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    Set vsoPagesEvent = doc.Pages
End Sub

'Таскаем фируру за мышкой
Private Sub vsoWindowEvent_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)

    If Not Click Then
        vsoShapePaste.Cells("PinX") = x
        vsoShapePaste.Cells("PinY") = y
    End If

End Sub

'Закончили таскать фигуру кликом мыши
Private Sub vsoWindowEvent_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
    Click = True
    Set vsoWindowEvent = Nothing
End Sub

Sub AutoNumEvent(vsoShapeEvent As Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : AutoNumEvent - Автонумерация для одиночной вставки
                'Когда происходит вставка применяется привязка к курсору
                'В EventDrop должна быть формула =CALLTHIS("ThisDocument.AutoNumEvent")
'------------------------------------------------------------------------------------------------------------
    Set vsoWindowEvent = ActiveWindow
    Set vsoShapePaste = vsoShapeEvent
    Click = False
    AutoNum vsoShapePaste
    
End Sub

Sub ClearReferenceEvent(vsoShapeEvent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearReferenceEvent - Чистит дочерний при копировании
                'Чистим ссылки в дочернем при его копировании.
                'В EventDrop должна быть формула = CALLTHIS("ThisDocument.ClearReferenceEvent")
'------------------------------------------------------------------------------------------------------------
    Set vsoWindowEvent = ActiveWindow
    Set vsoShapePaste = vsoShapeEvent
    Click = False
    ClearReference vsoShapePaste
    
End Sub

Sub ClearReferenceWireLinkEvent(vsoShapeEvent As Visio.Shape)
'------------------------------------------------------------------------------------------------------------
' Macros        : ClearReferenceWireLinkEvent - Чистит при копировании
                'Чистим ссылки в при копировании разрыва провода.
                'В EventDrop должна быть формула = CALLTHIS("ThisDocument.ClearReferenceWireLinkEvent")
'------------------------------------------------------------------------------------------------------------
    Set vsoWindowEvent = ActiveWindow
    Set vsoShapePaste = vsoShapeEvent
    Click = False
    ClearReferenceWireLink vsoShapePaste
End Sub


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    Set vsoPagesEvent = Nothing
    Set vsoWindowEvent = Nothing
End Sub

'Активация событий по кнопке в меню/на пенели
Sub InitEvent()
    Set vsoPagesEvent = ThisDocument.Pages
End Sub


