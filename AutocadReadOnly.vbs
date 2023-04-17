On Error Resume Next
dim cmd_arg
set cmd_arg=wscript.Arguments
Dim acadApp
Set acadApp = GetObject(, "AutoCAD.Application")
If Err Then
Err.Clear
Set acadApp = CreateObject("AutoCAD.Application")
End if
acadApp.visible=True
acadApp.WindowState = normal
For Each item in cmd_arg
acadApp.Documents.Open item, True
Next