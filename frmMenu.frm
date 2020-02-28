VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Image picChar 
      Height          =   975
      Index           =   5
      Left            =   6240
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image picChar 
      Height          =   975
      Index           =   4
      Left            =   5040
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image picChar 
      Height          =   975
      Index           =   3
      Left            =   3840
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image picChar 
      Height          =   975
      Index           =   2
      Left            =   2640
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image picChar 
      Height          =   975
      Index           =   1
      Left            =   1440
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmGame.Visible = True
    Unload Me
End Sub

Private Sub Form_Load()
Dim Path As String
Dim i As Long
    frmMenu.Height = 9420
    frmMenu.Width = 12090
    
    For i = 1 To 5
        Path = App.Path & "\sprites\faces\char" & i & ".gif"
        picChar(i).Picture = LoadPicture(Path)
        
    Next
    
End Sub

Private Sub picChar_Click(Index As Integer)
    lblTitulo.Caption = "Char Seleccionado " & Index
    Players(1).Char = Index
End Sub

