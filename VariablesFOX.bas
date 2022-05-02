Attribute VB_Name = "VariablesFOX"
Global cb As Long 'Used to store CODE4 pointer
Global db As Long 'Used to store DATA4 pointer
Global rc As Integer 'Used as general return code
Global fieldinfo() As FIELD4INFO
Global tagInfo() As TAG4INFO


Sub ConectarFox(frm As Form)
    
    cb = code4init()  'Initialize CodeBase
    If cb = 0 Then
        MsgBox "code4init( ) failed"
        Exit Sub
    Else
        frm.Show (1)  'Show Form 1 modally
    End If

    rc = code4initUndo(cb)   'Close everything and free resources

End Sub


