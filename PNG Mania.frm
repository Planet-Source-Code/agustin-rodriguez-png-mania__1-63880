VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00BB1E88&
   BorderStyle     =   0  'None
   ClientHeight    =   1380
   ClientLeft      =   -1005
   ClientTop       =   -105
   ClientWidth     =   1350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   3  'I-Beam
   Picture         =   "PNG Mania.frx":0000
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PNG As LayeredWindow

Private Sub Form_Load()
Set PNG = New LayeredWindow

Choise_mode (5)

End Sub

Private Sub Form_Unload(Cancel As Integer)
PNG.UnloadPNGForm
End Sub

Public Sub Choise_mode(x As Integer)
PNG.MakeTrans App.Path & "\" & Chr(65 + x) & Trim(Str(Int(Rnd * 5) + 1)) + ".png", Me
End Sub

