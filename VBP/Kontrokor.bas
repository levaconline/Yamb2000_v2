Attribute VB_Name = "Kontrokor"
Option Explicit
Public r As Integer, k As Integer, f As Integer

Sub Kont()
Dim r1 As Integer, r2 As Integer, r3 As Integer, r4 As Integer, r5 As Integer
Dim A1 As Integer, A2 As Integer, A3 As Integer, A4 As Integer, A5 As Integer
Dim k1 As Integer, k6 As Integer, rk As Integer, f As Integer
Dim rf As Integer, ka As Integer, Y As Integer

A1 = Val(Form1.Text8.Text)
A2 = Val(Form1.Text9.Text)
A3 = Val(Form1.Text10.Text)
A4 = Val(Form1.Text11.Text)
A5 = Val(Form1.Text12.Text)

Form1.B1 = Form1.Text13.Text
Form1.B2 = Form1.Text14.Text
Form1.B3 = Form1.Text15.Text
Form1.B4 = Form1.Text16.Text
Form1.B5 = Form1.Text17.Text
Form1.B6 = Form1.Text18.Text
k = 0
k1 = 0
k6 = 0
r = 0
f = 0

'RAZVRSTAVANJE
'***********************************
        If A1 = 1 Then
            r1 = 1
        End If
        
        If A2 = 1 Then
            r2 = 1
        End If
       
        If A3 = 1 Then
            r3 = 1
        End If
        
        If A4 = 1 Then
            r4 = 1
        End If
        
        If A5 = 1 Then
            r5 = 1
        End If
rk = r1 + r2 + r3 + r4 + r5
        
If rk >= 3 Then
    r = 3
    If rk = 4 Then
    ka = 4
    End If
    If rk = 5 Then
    Y = 5
    End If
    
Else
  '2222222222222222222
    If rk = 2 Then
    f = 1
    End If
'////////////////////
  
    If rk = 1 Then
    k1 = 1
    End If
    
End If
     r1 = 0: r2 = 0: r3 = 0: r4 = 0: r5 = 0

       
        If A1 = 2 Then
            r1 = 2
        End If
        
        If A2 = 2 Then
            r2 = 2
        End If
       
        If A3 = 2 Then
            r3 = 2
        End If
        
        If A4 = 2 Then
            r4 = 2
        End If
        
        If A5 = 2 Then
            r5 = 2
        End If
    rk = r1 + r2 + r3 + r4 + r5
        
If rk >= 6 Then
    r = 6
    If rk = 8 Then
    ka = 8
    End If
    If rk = 10 Then
    Y = 10
    End If
    
    
Else
'222222222222222222222
    If rk = 4 Then
    f = 1
    End If
'///////////////////////

    If rk = 0 Then
    k = k + 0
    Else
    k = k + 1
    End If
    
End If
      r1 = 0: r2 = 0: r3 = 0: r4 = 0: r5 = 0

      
        If A1 = 3 Then
            r1 = 3
        End If
        
        If A2 = 3 Then
            r2 = 3
        End If
       
        If A3 = 3 Then
            r3 = 3
        End If
        
        If A4 = 3 Then
            r4 = 3
        End If
        
        If A5 = 3 Then
            r5 = 3
        End If
    rk = r1 + r2 + r3 + r4 + r5
        
If rk >= 9 Then
    r = 9
    If rk = 12 Then
    ka = 12
    End If
    If rk = 15 Then
    Y = 15
    End If
    
    
Else
'222222222222222222222
    If rk = 6 Then
    f = 1
    End If
'/////////////////////////

    If rk = 0 Then
    k = k + 0
    Else
    k = k + 1
    End If
    
End If

'4444444444444444444444444444444444444444444444
'----------------------------------------------
      r1 = 0: r2 = 0: r3 = 0: r4 = 0: r5 = 0

      
        If A1 = 4 Then
            r1 = 4
        End If
        
        If A2 = 4 Then
            r2 = 4
        End If
       
        If A3 = 4 Then
            r3 = 4
        End If
        
        If A4 = 4 Then
            r4 = 4
        End If
        
        If A5 = 4 Then
            r5 = 4
        End If
    rk = r1 + r2 + r3 + r4 + r5
        
If rk >= 12 Then
    r = 12
    If rk = 16 Then
    ka = 16
    End If
    If rk = 20 Then
    Y = 20
    End If
    
    
Else
'22222222222222222222
    If rk = 8 Then
    f = 1
    End If
'////////////////////

    If rk = 0 Then
    k = k + 0
    Else
    k = k + 1
    End If
    
End If

'5555555555555555555555555555555555555555555
'-------------------------------------------
    r1 = 0: r2 = 0: r3 = 0: r4 = 0: r5 = 0


        If A1 = 5 Then
            r1 = 5
        End If
        
        If A2 = 5 Then
            r2 = 5
        End If
       
        If A3 = 5 Then
            r3 = 5
        End If
        
        If A4 = 5 Then
            r4 = 5
        End If
        
        If A5 = 5 Then
            r5 = 5
        End If
    rk = r1 + r2 + r3 + r4 + r5
        
If rk >= 15 Then
    r = 15
    If rk = 20 Then
        ka = 20
    End If
    If rk = 25 Then
    Y = 25
    End If
    
    
Else
'222222222222222222222
    If rk = 10 Then
    f = 1
    End If
'/////////////////////

    If rk = 0 Then
    k = k + 0
    Else
    k = k + 1
    End If
    
End If

'6666666666666666666666666666666666666666666666
'----------------------------------------------
    r1 = 0: r2 = 0: r3 = 0: r4 = 0: r5 = 0


        If A1 = 6 Then
            r1 = 6
        End If
        
        If A2 = 6 Then
            r2 = 6
        End If
       
        If A3 = 6 Then
            r3 = 6
        End If
        
        If A4 = 6 Then
            r4 = 6
        End If
        
        If A5 = 6 Then
            r5 = 6
        End If
        rk = r1 + r2 + r3 + r4 + r5
        
If rk >= 18 Then
    r = 18
    If rk = 24 Then
    ka = 24
    End If
    If rk = 30 Then
    Y = 30
    End If
    
    
Else
'22222222222222222222222
    If rk = 12 Then
    f = 1
    End If
'//////////////////////


    If rk = 6 Then
    k6 = 1
    End If
    
    r1 = 0: r2 = 0: r3 = 0: r4 = 0: r5 = 0

End If
'***********************************************


'FUL
'--------------------------------
        If r = 3 Or r = 6 Or r = 9 Or r = 12 Or r = 15 Or r = 18 Then
        rf = 1
        Else
        rf = 0
        End If
        
        If rf = 1 And f = 1 Then
        Form1.f = 1
        Else
        Form1.f = 0
        End If
'------------------------------------

'KENTA
'------------------------------------
If (k + k1) = 5 Or (k + k6) = 5 Then
Form1.k = 1
Else
Form1.k = 0
End If
'-------------------------------------

'KARO
'-------------------------------------
Form1.ka = ka
'-------------------------------------

'YAMB
'-------------------------------------
Form1.Ya = Y
'-------------------------------------



9:
End Sub
