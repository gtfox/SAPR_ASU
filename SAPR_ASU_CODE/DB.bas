'------------------------------------------------------------------------------------------------------------
' Module        : DB - База данных прайс листов и избранного
' Author        : gtfox
' Date          : 2021.02.22
' Description   : База данных прайс листов, избранного и их обеспечение
' Link          : https://visio.getbb.ru/viewtopic.php?f=44&t=1491, https://yadi.sk/d/24V8ngEM_8KXyg
'------------------------------------------------------------------------------------------------------------

Option Explicit

'Активация формы выбора элементов схемы из БД
Public Sub AddDBFrm(vsoShape As Visio.Shape) 'Получили шейп с листа
    Load frmDB
    frmDB.Run vsoShape 'Передали его в форму
End Sub

Public Function GetDBEngine() As Object
'Function returns DBEngine for current Office Engine Type (DAO.DBEngine.60 or DAO.DBEngine.120)
Dim engine As Object
    On Error GoTo EX
    Set GetDBEngine = DBEngine
Exit Function
EX:
    Set GetDBEngine = CreateObject("DAO.DBEngine.120")
End Function