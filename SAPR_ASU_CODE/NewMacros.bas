'
Sub Macro1()
Dim col As Collection
Dim vsoPage As Visio.Page
Dim i As Integer
Set col = New Collection
For Each vsoPage In ActiveDocument.Pages
col.Add vsoPage, CStr(i + 1)
Debug.Print vsoPage.Name
Debug.Print col(i + 1).Key
Next

End Sub
Sub Macro2()
    Application.CommandBars("Standard").Visible = False
    Application.CommandBars("Formatting").Visible = False
    Application.CommandBars("View").Visible = False
    Application.CommandBars("Data").Visible = False
    Application.CommandBars("Action").Visible = False
    Application.CommandBars("Stencil").Visible = False
    Application.CommandBars("Stop Recording").Visible = False
    Application.CommandBars("Snap & Glue").Visible = False
    Application.CommandBars("Developer").Visible = False
    Application.CommandBars("Drawing").Visible = False
    Application.CommandBars("Picture").Visible = False
    Application.CommandBars("Format Text").Visible = False
    Application.CommandBars("Format Shape").Visible = False
    Application.CommandBars("САПР АСУ").Visible = False
    Application.CommandBars("Standard").Visible = True
    Application.CommandBars("Formatting").Visible = True
    Application.CommandBars("Web").Visible = True
    Application.CommandBars("View").Visible = True
    Application.CommandBars("Data").Visible = True
    Application.CommandBars("Action").Visible = True
    Application.CommandBars("Layout & Routing").Visible = True
    Application.CommandBars("Stencil").Visible = True
    Application.CommandBars("Stop Recording").Visible = True
    Application.CommandBars("Snap & Glue").Visible = True
    Application.CommandBars("Developer").Visible = True
    Application.CommandBars("Reviewing").Visible = True
    Application.CommandBars("Drawing").Visible = True
    Application.CommandBars("Picture").Visible = True
    Application.CommandBars("Ink").Visible = True
    Application.CommandBars("Format Text").Visible = True
    Application.CommandBars("Format Shape").Visible = True
    Application.CommandBars("САПР АСУ").Visible = True

End Sub
Sub Macro3()
    Application.CommandBars("Reviewing").Visible = False
    Application.CommandBars("Web").Visible = False
    Application.CommandBars("Ink").Visible = False
    Application.CommandBars("Stencil").Visible = False
    Application.CommandBars("Picture").Visible = False
    Application.CommandBars("Layout & Routing").Visible = False
    Application.CommandBars("Data").Visible = False

End Sub