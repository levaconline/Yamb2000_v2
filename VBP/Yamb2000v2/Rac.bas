Attribute VB_Name = "Rac"
Option Explicit

Public A1 As Integer, A2 As Integer, A3 As Integer, A4 As Integer, A5 As Integer
Public Re As Integer, Gla As Integer, Bla As Integer, Ceta As Integer, Col As Integer


Sub Prv()
Dim r As Integer
r = 0
'Popis
Popis

'Rac
    If A1 = 1 Then
        r = r + 1
    End If
    If A2 = 1 Then
        r = r + 1
    End If
    If A3 = 1 Then
        r = r + 1
    End If
    If A4 = 1 Then
        r = r + 1
    End If
    If A5 = 1 Then
        r = r + 1
    End If

Re = r
' Brisanje polja
    Form1.Text8.Text = ""
    Form1.Text9.Text = ""
    Form1.Text10.Text = ""
    Form1.Text11.Text = ""
    Form1.Text12.Text = ""
    Form1.Text13.Text = ""
    Form1.Text14.Text = ""
    Form1.Text15.Text = ""
    Form1.Text16.Text = ""
    Form1.Text17.Text = ""
    Form1.Text18.Text = ""
    
'  Pospremanje
    Form1.a = 0: Form1.Command1.Enabled = True

End Sub

Sub Dru()
Dim r As Integer
r = 0
'Popis
Popis
'Rac
    If A1 = 2 Then
        r = r + 2
    End If
    If A2 = 2 Then
        r = r + 2
    End If
    If A3 = 2 Then
        r = r + 2
    End If
    If A4 = 2 Then
        r = r + 2
    End If
    If A5 = 2 Then
        r = r + 2
    End If
    
Re = r

'Brisanje polja
    Form1.Text8.Text = ""
    Form1.Text9.Text = ""
    Form1.Text10.Text = ""
    Form1.Text11.Text = ""
    Form1.Text12.Text = ""
    Form1.Text13.Text = ""
    Form1.Text14.Text = ""
    Form1.Text15.Text = ""
    Form1.Text16.Text = ""
    Form1.Text17.Text = ""
    Form1.Text18.Text = ""
    
'Pospremanje
    Form1.a = 0: Form1.Command1.Enabled = True

End Sub

Sub Tre()
Dim r As Integer
r = 0
'Popis
Popis
'Rac
        If A1 = 3 Then
            r = r + 3
        End If
        If A2 = 3 Then
            r = r + 3
        End If
        If A3 = 3 Then
            r = r + 3
        End If
        If A4 = 3 Then
            r = r + 3
        End If
        If A5 = 3 Then
            r = r + 3
        End If

Re = r
'Brisanje polja
    Form1.Text8.Text = ""
    Form1.Text9.Text = ""
    Form1.Text10.Text = ""
    Form1.Text11.Text = ""
    Form1.Text12.Text = ""
    Form1.Text13.Text = ""
    Form1.Text14.Text = ""
    Form1.Text15.Text = ""
    Form1.Text16.Text = ""
    Form1.Text17.Text = ""
    Form1.Text18.Text = ""
    
'Pospremanje
    Form1.a = 0: Form1.Command1.Enabled = True

End Sub

Sub Cet()
Dim r As Integer
r = 0
'Popis
Popis
'Rac
        If A1 = 4 Then
            r = r + 4
        End If
        If A2 = 4 Then
            r = r + 4
        End If
        If A3 = 4 Then
            r = r + 4
        End If
        If A4 = 4 Then
            r = r + 4
        End If
        If A5 = 4 Then
            r = r + 4
        End If

Re = r
'Brisanje polja
    Form1.Text8.Text = ""
    Form1.Text9.Text = ""
    Form1.Text10.Text = ""
    Form1.Text11.Text = ""
    Form1.Text12.Text = ""
    Form1.Text13.Text = ""
    Form1.Text14.Text = ""
    Form1.Text15.Text = ""
    Form1.Text16.Text = ""
    Form1.Text17.Text = ""
    Form1.Text18.Text = ""
    
'Pospremanje
    Form1.a = 0: Form1.Command1.Enabled = True


End Sub

Sub Pet()
Dim r As Integer
r = 0
'Popis
Popis
'Rac
        If A1 = 5 Then
            r = r + 5
        End If
        If A2 = 5 Then
            r = r + 5
        End If
        If A3 = 5 Then
            r = r + 5
        End If
        If A4 = 5 Then
            r = r + 5
        End If
        If A5 = 5 Then
            r = r + 5
        End If

Re = r
'Brisanje polja
    Form1.Text8.Text = ""
    Form1.Text9.Text = ""
    Form1.Text10.Text = ""
    Form1.Text11.Text = ""
    Form1.Text12.Text = ""
    Form1.Text13.Text = ""
    Form1.Text14.Text = ""
    Form1.Text15.Text = ""
    Form1.Text16.Text = ""
    Form1.Text17.Text = ""
    Form1.Text18.Text = ""
    
'Pospremanje
    Form1.a = 0: Form1.Command1.Enabled = True


End Sub
Sub Ses()
Dim r As Integer
r = 0
'Popis
Popis
'Rac
        If A1 = 6 Then
            r = r + 6
        End If
        If A2 = 6 Then
            r = r + 6
        End If
        If A3 = 6 Then
            r = r + 6
        End If
        If A4 = 6 Then
            r = r + 6
        End If
        If A5 = 6 Then
            r = r + 6
        End If

Re = r
'Brisanje polja
    Form1.Text8.Text = ""
    Form1.Text9.Text = ""
    Form1.Text10.Text = ""
    Form1.Text11.Text = ""
    Form1.Text12.Text = ""
    Form1.Text13.Text = ""
    Form1.Text14.Text = ""
    Form1.Text15.Text = ""
    Form1.Text16.Text = ""
    Form1.Text17.Text = ""
    Form1.Text18.Text = ""
    
'Pospremanje
    Form1.a = 0: Form1.Command1.Enabled = True


End Sub

Sub Popis()
A1 = Val(Form1.Text8.Text)
A2 = Val(Form1.Text9.Text)
A3 = Val(Form1.Text10.Text)
A4 = Val(Form1.Text11.Text)
A5 = Val(Form1.Text12.Text)
Form1.Deda = Form1.a
Form1.B1 = Form1.Text13.Text
Form1.B2 = Form1.Text14.Text
Form1.B3 = Form1.Text15.Text
Form1.B4 = Form1.Text16.Text
Form1.B5 = Form1.Text17.Text
Form1.B6 = Form1.Text18.Text

End Sub
Sub Glad()
Gla = 0
Col = 0


If Form1.Sn.ForeColor = "&h000000ff" Or Form1.nnn.ForeColor = "&h000000ff" Then Col = 1


If Form1.Label13.Caption = "" Or Form1.Label27.Caption = "" Then
GoTo 336
End If

If Form1.Label14.Caption = "" Or Form1.Label15.Caption = "" Then
GoTo 336
End If

If Form1.Label16.Caption = "" Or Form1.Label17.Caption = "" Then
GoTo 336
End If

If Form1.Label18.Caption = "" Or Form1.Label19.Caption = "" Then
GoTo 336
End If

If Form1.Label20.Caption = "" Or Form1.Label21.Caption = "" Then
GoTo 336
End If

If Form1.Label22.Caption = "" Or Form1.Label23.Caption = "" Then
GoTo 336
End If

If Form1.Label24.Caption = "" Or Form1.Label25.Caption = "" Then
GoTo 336
End If

If Form1.Label26.Caption = "" Or Form1.Label52.Caption = "" Or Form1.Label40.Caption = "" Then
GoTo 336
End If
If Col = 1 Then GoTo 336
Gla = 1
336:
Ceta = 0
    If Form1.Label53.Caption = "" Or Form1.Label54.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If
    If Form1.Label55.Caption = "" Or Form1.Label56.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If
    If Form1.Label57.Caption = "" Or Form1.Label58.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If
    If Form1.Label59.Caption = "" Or Form1.Label60.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If
    If Form1.Label61.Caption = "" Or Form1.Label62.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If
    If Form1.Label63.Caption = "" Or Form1.Label64.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If
    If Form1.Label65.Caption = "" Then
    Ceta = 1
    Exit Sub
    End If

If Col = 1 Then Exit Sub
If Gla = 0 Then Exit Sub
MsgBox ("Najavite ili igrajte, jer nebiste imali gde da odigrate!")

End Sub
Sub RuB()
Bla = 0
If Gla = 0 Then Exit Sub
If (Form1.Ruc = 0 Or Form1.Ruc = 6) And Ceta = 1 Then Exit Sub
If Form1.Sn.ForeColor = "&h000000ff" Or Form1.nnn.ForeColor = "&h000000ff" Then Exit Sub

Form1.Text13.Text = Form1.B1
Form1.Text14.Text = Form1.B2
Form1.Text15.Text = Form1.B3
Form1.Text16.Text = Form1.B4
Form1.Text17.Text = Form1.B5
Form1.Text18.Text = Form1.B6

Bla = 1
If Form1.a > 1 Then
MsgBox ("Mozete da igrate samo >>RUCNU<<, dakle morate da >>bacite svih 6 kockica<<!")
Else
MsgBox ("Najavite ili igrajte >>RUCNU<<!")
End If

End Sub
