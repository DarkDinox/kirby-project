Attribute VB_Name = "modGraphics"
Public Sub RenderGraphics()

    Dim Path As String
    
    If AttackDown = True Then
        If PlayerDirection = DIR_RIGHT Then
            If Anim >= 42 Or Anim < 38 Then Anim = 38
                            
            Anim = Anim + 1
        Else
            If Anim >= 47 Or Anim < 43 Then Anim = 43
                            
            Anim = Anim + 1
        End If
        Path = App.Path & "\sprites\spr" & Anim & ".gif"
        frmGame.picPlayer.Picture = LoadPicture(Path)
        If Anim = 42 Or Anim = 47 Then AttackDown = False
        Exit Sub
    End If
    
    If inGround = False And Jumping = False Then
        Select Case PlayerDirection
            Case DIR_RIGHT
                Path = App.Path & "\sprites\spr" & 17 & ".gif"
                frmGame.picPlayer.Picture = LoadPicture(Path)
            Case DIR_LEFT
                Path = App.Path & "\sprites\spr" & 19 & ".gif"
                frmGame.picPlayer.Picture = LoadPicture(Path)
            End Select
        Exit Sub
    End If
    
    If HP = 0 Then
        If PlayerDirection = DIR_RIGHT Then
            If Anim >= 54 Or Anim < 48 Then Anim = 47
                            
            Anim = Anim + 1
        Else
            If Anim >= 47 Or Anim < 43 Then Anim = 43
                            
            Anim = Anim + 1
        End If

        If Anim = 54 Then
            'HP = 100
            Exit Sub
        End If
        
        Path = App.Path & "\sprites\spr" & Anim & ".gif"
        frmGame.picPlayer.Picture = LoadPicture(Path)
        Exit Sub
        
    End If
    
    If ShiftDown = True Then
                
        If PlayerDirection = DIR_RIGHT Then
            Anim = 20
        Else
            Anim = 21
        End If
                
        Path = App.Path & "\sprites\spr" & Anim & ".gif"
        frmGame.picPlayer.Picture = LoadPicture(Path)
        Exit Sub
    End If
    
    Select Case PlayerDirection
        Case DIR_DOWN
            
            Path = App.Path & "\sprites\spr17.gif"
            frmGame.picPlayer.Picture = LoadPicture(Path)
        Case DIR_UP
            Path = App.Path & "\sprites\spr0.gif"
            frmGame.picPlayer.Picture = LoadPicture(Path)
        Case DIR_RIGHT
            Select Case Jumping
                Case True
                    If InAir = True Then
                        Anim = 17
                    Else
                        Anim = 16
                    End If
                    
                    Path = App.Path & "\sprites\spr" & Anim & ".gif"
                    frmGame.picPlayer.Picture = LoadPicture(Path)
                Case False
                    If ShiftUp = True And InMove = True Then
                        If Anim >= 29 Or Anim < 22 Then Anim = 22
                                
                        If DirRight = True Then
                            Anim = Anim + 1
                        Else
                            Anim = 22
                        End If
                            
                        Path = App.Path & "\sprites\spr" & Anim & ".gif"
                        frmGame.picPlayer.Picture = LoadPicture(Path)
                    Else
                        If Anim >= 7 Then Anim = 0
                            
                        If DirRight = True Then
                            Anim = Anim + 1
                        Else
                            Anim = 0
                        End If
                        
                        Path = App.Path & "\sprites\spr" & Anim & ".gif"
                        frmGame.picPlayer.Picture = LoadPicture(Path)
                    End If
            End Select
        Case DIR_LEFT
            Select Case Jumping
                Case True
                    If InAir = True Then
                        Anim = 19
                    Else
                        Anim = 18
                    End If
                    Path = App.Path & "\sprites\spr" & Anim & ".gif"
                    frmGame.picPlayer.Picture = LoadPicture(Path)
                Case False
                    If ShiftUp = True And InMove = True Then
                       If Anim >= 37 Or Anim < 30 Then Anim = 30
                        
                        If DirLeft = True Then
                            Anim = Anim + 1
                        Else
                            Anim = 30
                        End If
                        Path = App.Path & "\sprites\spr" & Anim & ".gif"
                        frmGame.picPlayer.Picture = LoadPicture(Path)
                    Else
                        If Anim >= 15 Or Anim < 8 Then Anim = 8
                        
                        If DirLeft = True Then
                            Anim = Anim + 1
                        Else
                            Anim = 8
                        End If
                        Path = App.Path & "\sprites\spr" & Anim & ".gif"
                        frmGame.picPlayer.Picture = LoadPicture(Path)
                    End If
            End Select
        Case Else
            'NOTHING
    End Select


End Sub

Public Sub DrawEnemy()
If Enemy(1).HP > 0 Then
    Select Case Enemy(1).Dir
        Case 1
                If Enemy(1).Anim >= 5 Or Enemy(1).Anim < 1 Then Enemy(1).Anim = 1
                                    
                If Enemy(1).Dir = 1 And Enemy(1).Movetimer > 0 Then
                    Enemy(1).Anim = Enemy(1).Anim + 1
                Else
                    Enemy(1).Anim = 1
                End If
                                
                Path = App.Path & "\sprites\enemigos\spr" & Enemy(1).Anim & ".gif"
                frmGame.picEnemy(1).Picture = LoadPicture(Path)
        Case 2
            If Enemy(1).Anim >= 10 Or Enemy(1).Anim < 6 Then Enemy(1).Anim = 6
                                
            If Enemy(1).Dir = 2 And Enemy(1).Movetimer > 0 Then
                Enemy(1).Anim = Enemy(1).Anim + 1
            Else
                Enemy(1).Anim = 6
            End If
                            
            Path = App.Path & "\sprites\enemigos\spr" & Enemy(1).Anim & ".gif"
            frmGame.picEnemy(1).Picture = LoadPicture(Path)
    End Select
    End If
End Sub


Public Sub DrawMuerte()

 If Not Enemy(1).HP = 0 Then Exit Sub
 
 
    Select Case Enemy(1).Dir
        Case 1
                If Enemy(1).Anim >= 15 Or Enemy(1).Anim < 11 Then Enemy(1).Anim = 10

                Enemy(1).Anim = Enemy(1).Anim + 1
                
                If Enemy(1).Anim > 14 Then
                    frmGame.picEnemy(1).Visible = False
                    Exit Sub
                End If
                
                Path = App.Path & "\sprites\enemigos\spr" & Enemy(1).Anim & ".gif"
                frmGame.picEnemy(1).Picture = LoadPicture(Path)
        Case 2
                If Enemy(1).Anim >= 19 Or Enemy(1).Anim < 15 Then Enemy(1).Anim = 14

                Enemy(1).Anim = Enemy(1).Anim + 1
                
                If Enemy(1).Anim > 18 Then
                    frmGame.picEnemy(1).Visible = False
                    Exit Sub
                End If
                
                Path = App.Path & "\sprites\enemigos\spr" & Enemy(1).Anim & ".gif"
                frmGame.picEnemy(1).Picture = LoadPicture(Path)
    End Select
End Sub

Public Sub DrawProjectil()

With frmGame
    If ProjectilLaunch = True Then
        Select Case Projectil(1).Dir
            Case DIR_RIGHT
                If ProX >= 11600 Then
                    Projectil(1).Active = False
                    ProjectilExplosion = True
                    ProjectilLaunch = False
                    Exit Sub
                End If
                ProX = ProX + 35
                .picProjectil.Left = ProX
                .picProjectil.Top = ProY
            Case DIR_LEFT
                If ProX <= 0 Then
                    Projectil(1).Active = False
                    ProjectilExplosion = True
                    ProjectilLaunch = False
                    Exit Sub
                End If
                ProX = ProX - 35
                .picProjectil.Left = ProX
                .picProjectil.Top = ProY
        End Select
    End If
    
    If .picProjectil.Visible = True And Projectil(1).Active = True Then
        If (.picEnemy(1).Left < .picProjectil.Left + .picProjectil.Width And .picEnemy(1).Left + .picEnemy(1).Width > .picProjectil.Left And .picEnemy(1).Top < .picProjectil.Top + .picProjectil.Height And .picEnemy(1).Top + .picEnemy(1).Height > .picProjectil.Top) Then
            If Enemy(1).HP > 0 Then Enemy(1).HP = Enemy(1).HP - 10
                'Target del Enemigo
                Enemy(1).Target = True
                Projectil(1).Active = False
                ProjectilExplosion = True
                ProjectilLaunch = False
        End If
    End If
End With
End Sub
