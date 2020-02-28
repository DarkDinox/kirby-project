Attribute VB_Name = "modText"
Public Sub DrawEnemyName()
    MsgY = EnmY - 280
    MsgX = EnmX - 200
    frmGame.lblEnemyHP.Top = MsgY
    frmGame.lblEnemyHP.Left = MsgX
    
    If Enemy(1).HP > 0 Then
        frmGame.lblEnemyHP.Caption = Trim$(Enemy(1).HP)
    Else
        frmGame.lblEnemyHP.Caption = "Derrotado"
        frmGame.lblEnemyHP.Visible = False
    End If
    
End Sub
