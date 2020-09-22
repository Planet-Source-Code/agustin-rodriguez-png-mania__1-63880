Attribute VB_Name = "Module1"
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim f(0 To 30) As New Form1
Public delay As Integer

Private Sub Main()
'Credits to this procedure to Takis Firipidis
'Credits to PNG class to Apeiron

Dim buf(100, 4), rel As Single
Dim x0 As Integer
Form2.Show
For i = 0 To 30
f(i).Show
Next


x0 = Screen.Width / 2    'position x  that bubbles appear
  rel = 0.54             'relation between bubble velocity and wave angle
 

 
  For q = 0 To 30
    buf(q, 0) = Rnd(1) * 1000 + 50           'wave amplitude
    buf(q, 1) = Rnd(1) * 3                'wave angle
    buf(q, 2) = Screen.Height + 10          'current y position
    buf(q, 3) = Rnd(1) * 200 + 50            'bubble velocity
    buf(q, 4) = Rnd(1) * 200 + 50            'bubble radius
  Next
  
 Do While DoEvents()

   For q = 0 To 30
           
    platos = buf(q, 0)
    ph = buf(q, 1)
    Y = buf(q, 2)
    vel = buf(q, 3)
    angle = buf(q, 4)

    x = x0 + platos * Sin(ph * 6.28)
    f(q).Move x, Y
      
    Y = Y - vel
    ph = ph + rel / vel
    
    If Y < 0 Then
      buf(q, 0) = Rnd(1) * 1000 + 20
      Y = Screen.Height + 10
      ph = Rnd(1) * 3
      buf(q, 3) = Rnd(1) * 200 + 20
      buf(q, 4) = Rnd(1) * 200 + 20
      x0 = x0 + 10
      If x0 > Screen.Width Then x0 = 0
    End If
             
    buf(q, 1) = ph
    buf(q, 2) = Y
    
    x = x0 + platos * Sin(ph * 6.28)
    f(q).Move x, Y
      
    For t = 1 To 10000: Next

xx:
    If GetAsyncKeyState(16) Then
        If GetAsyncKeyState(27) Then
            Finalize
            End
        End If
               
        If GetAsyncKeyState(112) Then
            For i = 0 To 30
            f(i).Choise_mode (1)
            Next i
        End If
        
        If GetAsyncKeyState(113) Then
            For i = 0 To 30
            f(i).Choise_mode (2)
            Next i
        End If
        
        If GetAsyncKeyState(114) Then
            For i = 0 To 30
            f(i).Choise_mode (3)
            Next i
        End If
        
        If GetAsyncKeyState(115) Then
            For i = 0 To 30
            f(i).Choise_mode (4)
            Next i
        End If
        
        If GetAsyncKeyState(116) Then
            For i = 0 To 30
            f(i).Choise_mode (5)
            Next i
        End If
        
        If GetAsyncKeyState(117) Then
            For i = 0 To 30
            f(i).Choise_mode (6)
            Next i
        End If
        
        If GetAsyncKeyState(118) Then
            For i = 0 To 30
            f(i).Choise_mode (Int(Rnd * 6 + 1))
            Next i
        End If
        
        If GetAsyncKeyState(119) Then
            delay = 0
            Form2.Show
        End If
    End If
    
    Next
   Loop
   
End Sub
Public Sub Finalize()
Dim i As Integer
For i = 0 To 30
    Unload f(i)
Next

Unload Form2
End Sub
