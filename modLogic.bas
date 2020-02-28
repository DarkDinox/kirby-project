Attribute VB_Name = "modLogic"
Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    
    Rand = Int((High - Low + 1) * Rnd) + Low

End Function

Public Sub CreateProjectil()
    frmGame.picProjectil.Visible = True
    
    Select Case PlayerDirection
        Case DIR_RIGHT
            ProX = PosX + 150
            ProY = PosY + 100
        Case DIR_LEFT
            ProX = PosX + 80
            ProY = PosY + 100
    End Select
    
    Projectil(1).Active = True
    
    Projectil(1).Dir = PlayerDirection
    
    Path = App.Path & "\sprites\anim" & 0 & ".gif"
    frmGame.picProjectil.Picture = LoadPicture(Path)
    
End Sub

Public Sub CheckNPCCollision()

    With frmGame
    If .picPlayer.Visible = True And .picEnemy(1).Visible = True Then
        If (.picEnemy(1).Left < .picPlayer.Left + .picPlayer.Width And .picEnemy(1).Left + .picEnemy(1).Width > .picPlayer.Left And .picEnemy(1).Top < .picPlayer.Top + .picPlayer.Height And .picEnemy(1).Top + .picEnemy(1).Height > .picPlayer.Top) Then
            If HP > 0 Then
                Enemy(1).Target = True
                HP = HP - 1
                .lblHP.Caption = HP
            End If
        End If
    End If
    End With
End Sub

Public Function CheckCollision() As Boolean
Dim i As Long

With frmGame
    For i = 0 To 17
        If .Tile(i).Visible = True Then
            If ((.picPlayer.Left + 250) < .Tile(i).Left + .Tile(i).Width And (.picPlayer.Left + 250) + PWIDTH > .Tile(i).Left And (.picPlayer.Top + 400) < .Tile(i).Top + .Tile(i).Height And (.picPlayer.Top + 400) + PHEIGHT > .Tile(i).Top) Then
            'If (picPlayer.Left >= Tile(0)Left And picPlayer.Left <= Tile(0)Left + Tile(0)Width) And (picPlayer.Top >= Tile(0)Top And picPlayer.Top <= Tile(0)Top + Tile(0)Height) Then
                If i = 9 Then
                    If Lado = False Then
                        PosX = PosX + 10
                    Else
                        PosX = PosX - 10
                    End If
                End If
                inGround = True
                If AntiFly = True Then AntiFly = False ' 1 Time
                CheckCollision = False
                Exit Function
            Else
                inGround = False
                CheckCollision = True
            End If
        End If
    Next
    End With
End Function

