VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   735
      Top             =   1440
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PNG As LayeredWindow



Private Sub Form_Load()



Set PNG = New LayeredWindow

PNG.MakeTrans App.Path & "\splash.png", Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
PNG.UnloadPNGForm
End Sub



Private Sub Timer1_Timer()

delay = delay + 1
If delay < 300 Then Exit Sub

Move Left, Top + 20
DoEvents
If Top > Screen.Height Then
Unload Me
End If

End Sub
