Attribute VB_Name = "modControls"

Public Sub PlayerMove(KeyNum As Long)

    If vbKeyUp = KeyNum Then
        DirUp = True
        DirDown = False
        DirRight = False
        DirLeft = False
    Else
        DirUp = False
    End If
    
    If vbKey2 = KeyNum Then
        DirUp = False
        DirDown = True
        DirRight = False
        DirLeft = False
    Else
        DirDown = False
    End If
    
End Sub

Public Sub CheckKeys()

    If Not GetAsyncKeyState(vbKeyZ) >= 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    
    'Si la Defensa esta activada salimos del SUB
    If ShiftDown = True Then Exit Sub
    
    If Not GetAsyncKeyState(vbKeySpace) >= 0 And inGround = True Then Jumping = True
    
    If Not GetAsyncKeyState(vbKeyUp) >= 0 Then
        DirUp = True
        DirDown = False
        DirRight = False
        DirLeft = False
        PlayerDirection = DIR_UP
    Else
        DirUp = False
    End If

    If Not GetAsyncKeyState(vbKeyDown) >= 0 Then
        DirUp = False
        DirDown = True
        DirRight = False
        DirLeft = False
        PlayerDirection = DIR_DOWN
    Else
        DirDown = False
    End If
    

    If Not GetAsyncKeyState(vbKeyRight) >= 0 Then
        DirUp = False
        DirDown = False
        DirRight = True
        DirLeft = False
        PlayerDirection = DIR_RIGHT
    Else
        DirRight = False
    End If

    If Not GetAsyncKeyState(vbKeyLeft) >= 0 Then
        DirUp = False
        DirDown = False
        DirRight = False
        DirLeft = True
        PlayerDirection = DIR_LEFT
    Else
        DirLeft = False
    End If
    
    If DirRight = True Or DirLeft = True Then
        InMove = True
    Else
        InMove = False
    End If
End Sub
