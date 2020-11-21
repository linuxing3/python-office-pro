' 命名空间
Attribute VB_Name="module_fs"

''namespace=vba-files\helper
Sub test()
    '// add declarations
    On Error GoTo catchError
    MsgBox "hello xvba", vbButtonType, "Tips"
exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub

Public Sub test_ai()
    dim ai as ClassAI
End Sub