Sub Timer1_Timer()
    Dim lNotActive
    'lNotActive = Long
    Dim LastCursorPos
    'LastCursorPos = POINTAPI
    Dim CursorPos
    'CursorPos = POINTAPI
    Dim lCounter
   ' lCounter = Integer

    Dim User32
    Set User32 = XNHost.LoadDll( "user32.dll" )
    Dim Point
    Set Point = XNHost.Struct
    Point.Add "x" , "i32"
    Point.Add "y" , "i32"
    User32.GetCursorPos Point

    MsgBox "Mouse x Position = " & Point.x 
    MsgBox "Mouse y Position = " & Point.y
    lNotActive = lNotActive + 1
    WScript.echo GetCursorPos.y
    GetCursorPos CursorPos
    If (CursorPos.x <> LastCursorPos.x) Or (CursorPos.y <> LastCursorPos.y) Then
        'The cursor of the mouse is not the same position as before
        LastCursorPos = CursorPos
        lNotActive = 0
    Else
        'The mouse is at the exact same position as 1 second before... check if any key was pressed.
        For lCounter = 0 To 255
            If GetAsyncKeyState(lCounter) <> 0 Then
                'A key has been pressed.
                lNotActive = 0
            End If
        Next
    End If
    
    If lNotActive = 10 Then
        '  An idle-minute.
        'DO WHATEVER YOU NEED HERE
        lNotActive = 0
        WScript.echo "is Inactive"
    End If
End Sub

DO
	Timer1_Timer()
    WScript.Sleep 1000
loop