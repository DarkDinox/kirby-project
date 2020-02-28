VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGame 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kirby Project"
   ClientHeight    =   9000
   ClientLeft      =   150
   ClientTop       =   3900
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrxd 
      Interval        =   200
      Left            =   120
      Top             =   2160
   End
   Begin VB.Timer NPCLogic 
      Interval        =   1000
      Left            =   120
      Top             =   1680
   End
   Begin VB.Timer tmrGraphics 
      Interval        =   50
      Left            =   120
      Top             =   1200
   End
   Begin VB.Timer Loop 
      Interval        =   20
      Left            =   600
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1200
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "PLAY"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "REINICIAR"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lblFio 
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   17
      Left            =   10870
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   16
      Left            =   8880
      Top             =   3530
      Width           =   855
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   15
      Left            =   9240
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   14
      Left            =   10080
      Top             =   5450
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   13
      Left            =   8400
      Top             =   5820
      Width           =   1335
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   12
      Left            =   10800
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   11
      Left            =   8040
      Top             =   6960
      Width           =   855
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   10
      Left            =   3120
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   9
      Left            =   4080
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image picFace 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   10320
      Top             =   45
      Width           =   975
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   8
      Left            =   9240
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   7
      Left            =   1200
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   6
      Left            =   1560
      Top             =   4280
      Width           =   1575
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   5
      Left            =   2400
      Top             =   5450
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   4
      Left            =   360
      Top             =   6960
      Width           =   735
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   3
      Left            =   720
      Top             =   5790
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   2
      Left            =   3120
      Top             =   6555
      Width           =   1215
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11400
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.Image picProjectil 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1560
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblEnemyHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NPC 01"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   7320
      Width           =   855
   End
   Begin VB.Image picEnemy 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   1
      Left            =   7200
      Top             =   7560
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpMusic 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   8400
      Visible         =   0   'False
      Width           =   2175
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3836
      _cy             =   873
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   1560
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Image Tile 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   8040
      Width           =   12015
   End
   Begin VB.Image picPlayer 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   1680
      Top             =   1560
      Width           =   735
   End
   Begin VB.Image picPoint 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   4920
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KIRBY PROJECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblResultado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TU PUNTUACIÓN: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label lblDireccion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Coordenadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image BackGroundS 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   11280
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdReset_Click()
    GameState = Main_Menu
    Puntos = 0
End Sub

Private Sub cmdStart_Click()
    GameState = Main_Game
    HP = 100
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 88 Or KeyAscii = 120 And AttackDown = False And ProjectilLaunch = False Then
        If ProjectilExplosion = False And AntiFly = False Then
            CreateProjectil
            ProjectilLaunch = True
            AttackDown = True
        End If
    End If
End Sub


Public Sub CheckPoints()
    
    If picPoint.Visible = True Then
        If (picPlayer.Left < picPoint.Left + picPoint.Width And picPlayer.Left + picPlayer.Width > picPoint.Left And picPlayer.Top < picPoint.Top + picPoint.Height And picPlayer.Top + picPlayer.Height > picPoint.Top) Then
        'If (picPlayer.Left >= picPoint.Left And picPlayer.Left <= picPoint.Left + picPoint.Width) And (picPlayer.Top >= picPoint.Top And picPlayer.Top <= picPoint.Top + picPoint.Height) Then
            Puntos = Puntos + 50
            picPoint.Visible = False
        End If
    End If
    
End Sub

Private Sub Form_Load()
Dim Path As String
    'Inicia el Menu al cargar el Formulario
    frmGame.Height = 9420
    frmGame.Width = 12090
    
    Path = App.Path & "\sprites\backgrounds\bga0.jpg"
    frmGame.Picture = LoadPicture(Path)
    
    TileX = Tile(9).Left
    BackGroundS.Top = 0
    BackGroundS.Left = 0
    
    Path = App.Path & "\sprites\backgrounds\bgacap0.gif"
    BackGroundS.Picture = LoadPicture(Path)
    
    Path = App.Path & "\sprites\faces\char0.gif"
    picFace.Picture = LoadPicture(Path)
    
    GameState = Main_Menu
    PlayerDirection = DIR_RIGHT
    AntiFly = True
    'Crear NPC-
    picPlayer.Top = 0
    picPlayer.Left = 0
    'Variables de Enemigos ;D
    Enemy(1).Dir = 1
    Enemy(1).HP = 100
    wmpMusic.URL = App.Path & "\music\ost01.mp3"
End Sub

Private Sub NPCLogic_Timer()
    Dim Move As Byte
    
    'Si tiene HP continuamos
    If Enemy(1).HP > 0 And picEnemy(1).Visible Then
    
    If Enemy(1).Target = True Then
        'Movimiento del NPC Según la coor del objetivo
        If EnmX > PosX Then
            Move = DIR_LEFT
        Else
            Move = DIR_RIGHT
        End If
    Else
        'Movimiento del NPC sin objetivo
        Move = Rand(1, 2)
    End If
    
    Select Case Move
        Case 1
            Enemy(1).Dir = 1
            Enemy(1).Movetimer = 1
            EnmX = picEnemy(1).Left
            EnmY = picEnemy(1).Top
        Case 2
            Enemy(1).Dir = 2
            Enemy(1).Movetimer = 1
            EnmX = picEnemy(1).Left
            EnmY = picEnemy(1).Top
    End Select
    
    End If
    
End Sub



Private Sub Timer1_Timer()
    If lblFio.Left > 4500 Then lblFio.Left = -2000
    lblFio.Left = lblFio.Left + 25
    lblFio.Caption = "Fio <3"
    
    Call DrawProjectil
    
    'Check Collision
    CheckCollision
    
    If Enemy(1).Dir > 0 And Enemy(1).Movetimer > 0 Then
        Select Case Enemy(1).Dir
            Case 1 'Derecha
                 If Not EnmX >= 11600 Then EnmX = EnmX + 10
            Case 2 'Izquierda
                If Not EnmX <= 0 Then EnmX = EnmX - 10
        End Select
        Enemy(1).Movetimer = Enemy(1).Movetimer + 1
        picEnemy(1).Left = EnmX
        If Enemy(1).Movetimer >= 50 Then
            'Enemy(1).Dir = 0
            Enemy(1).Movetimer = 0
        End If
    End If
    
    
    If TileX <= 4080 Then Lado = False
    If TileX >= 7560 Then Lado = True
    
    If Lado = False Then
        TileX = TileX + 10
    Else
        TileX = TileX - 10
    End If
    
    Tile(9).Left = TileX
    
    
    
    'Dibujar nombre de Enemigo
    DrawEnemyName
    
    lblDireccion.Caption = "X: " & Int(PosX / 10) & " " & "Y: " & Int(PosY / 10)
    
    If Puntos >= 500 Or HP = 0 And GameState = Main_Game Then
        GameState = Main_Result
    End If
    
    If GameState = Main_Result Then
        lblResultado.Visible = True
        cmdReset.Visible = True
        lblResultado.Caption = "TU PUNTUACIÓN: " & Puntos
        Exit Sub
    Else
        cmdReset.Visible = False
        lblResultado.Visible = False
    End If
    
    If GameState = Main_Menu Then
        picPlayer.Visible = False
        lblTitulo.Visible = True
        cmdStart.Visible = True
        lblPoints.Visible = False
        Exit Sub
    Else
        lblTitulo.Visible = False
        cmdStart.Visible = False
    End If
    
    If GameState = Main_Game Then
        picPlayer.Visible = True
        picPoint.Visible = True
        lblPoints.Visible = True
        GameState = Main_Play
    End If
    
    If ProjectilExplosion = True Then
        TimerExplosion = TimerExplosion + 10
        Path = App.Path & "\sprites\anim" & 1 & ".gif"
        frmGame.picProjectil.Picture = LoadPicture(Path)
        If TimerExplosion >= 100 Then
            Path = App.Path & "\sprites\anim" & 1 & ".gif"
            frmGame.picProjectil.Picture = LoadPicture(Path)
            picProjectil.Visible = False
            ProjectilExplosion = False
            TimerExplosion = 0
        End If
    End If
    
    'Ejecutar movimiento
    CheckKeys
    CheckPoints
    
    If Not GetAsyncKeyState(vbKeyShift) >= 0 Then
        ShiftUp = True
    Else
        ShiftUp = False
    End If
    
    lblPoints.Caption = "Puntos: " & Trim$(Puntos)
    
    'If CanMove = True Then
    
    If picPoint.Visible = False Then
        If Not PointTimer >= 250 Then PointTimer = PointTimer + 10
    End If
    
    If PointTimer >= 250 Then
        picPoint.Visible = True
        picPoint.Left = Rand(0, 7000)
        picPoint.Top = Rand(0, 5500)
        'Volver a Cero
        PointTimer = 0
    End If
        
        If DirUp Then
            'Movimiento del Jugador
            If Not picPlayer.Top <= 0 Then
                Select Case ShiftUp
                    Case True
                        PosY = PosY - RunSpeed
                    Case False
                        PosY = PosY - WalkSpeed
                End Select
            End If
        End If
        
        'If DirDown = True Then
        '    'Movimiento del Jugador
        '    If Not picPlayer.Top >= 5650 Then
        '        If CheckCollision = True Then
        '            Select Case ShiftUp
        '                Case True
        '                    PosY = PosY + RunSpeed
        '                Case False
        '                    PosY = PosY + WalkSpeed
        '            End Select
        '        End If
        '    End If
        'End If
        
        If DirLeft = True And Not ShiftDown Then
            'Movimiento del Jugador
            If Not picPlayer.Left <= 0 Then
                'If CheckCollision = True Then
                    Select Case ShiftUp
                        Case True
                            PosX = PosX - RunSpeed
                        Case False
                            PosX = picPlayer.Left - WalkSpeed
                    End Select
                'End If
            End If
        End If
        
        If DirRight = True And ShiftDown = False Then
            'Movimiento del Jugador
            If Not picPlayer.Left >= 11600 Then
                Select Case ShiftUp
                    Case True
                        PosX = PosX + RunSpeed
                    Case False
                        PosX = PosX + WalkSpeed
                End Select
            End If
        End If
        
        picPlayer.Left = PosX
        picPlayer.Top = PosY
    
End Sub

Private Sub Loop_Timer()
    
    If Jumping = True Then
       ' If Gravity = 0 Then Exit Sub
        'Gravedad Saltar
        If Not Gravity >= 150 And InAir = False Then
            Gravity = Gravity + 10
            PosY = PosY - Gravity
            If Gravity >= 150 Then InAir = True
        End If
        
        'Gravedad Caida
        If InAir = True And Gravity > 0 Then
            Gravity = Gravity - 5
            PosY = PosY + (Gravity + 10)
            If inGround = False Then
                InAir = False
                Gravity = 0
                Jumping = False
                'inGround = True
                'Exit Sub
            End If
        ElseIf InAir = True And Gravity = 0 Then
            InAir = False
            Gravity = 0
            Jumping = False
        End If
        
    End If
    
    If Jumping = False And inGround = False Then
         PosY = PosY + 50
    End If

    picPlayer.Top = PosY
End Sub

Private Sub tmrGraphics_Timer()
    Call RenderGraphics
    Call DrawEnemy
    Call CheckNPCCollision
End Sub

Private Sub tmrxd_Timer()
    Call DrawMuerte
End Sub
