Attribute VB_Name = "Und"
Option Explicit
Sub Und()
Dim p As Integer, t As Integer
Form1.Undo = False
t = Form1.t
Form1.a = Form1.Deda
p = Form1.p
Form1.Text8.Text = Rac.A1
If Rac.A1 = 0 Then Form1.Text8.Text = ""
Form1.Text9.Text = Rac.A2
If Rac.A2 = 0 Then Form1.Text9.Text = ""
Form1.Text10.Text = Rac.A3
If Rac.A3 = 0 Then Form1.Text10.Text = ""
Form1.Text11.Text = Rac.A4
If Rac.A4 = 0 Then Form1.Text11.Text = ""
Form1.Text12.Text = Rac.A5
If Rac.A5 = 0 Then Form1.Text12.Text = ""

Form1.Text13.Text = Form1.B1
Form1.Text14.Text = Form1.B2
Form1.Text15.Text = Form1.B3
Form1.Text16.Text = Form1.B4
Form1.Text17.Text = Form1.B5
Form1.Text18.Text = Form1.B6

If p = 92 Then
    Form1.Label92.Caption = ""
    Form1.Text19.Text = Val(Form1.Label92.Caption) + Val(Form1.Label2.Caption) + Val(Form1.Label3.Caption) + Val(Form1.Label4.Caption) + Val(Form1.Label5.Caption) + Val(Form1.Label6.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 2 Then
    Form1.Label2.Caption = ""
    Form1.Text19.Text = Val(Form1.Label92.Caption) + Val(Form1.Label2.Caption) + Val(Form1.Label3.Caption) + Val(Form1.Label4.Caption) + Val(Form1.Label5.Caption) + Val(Form1.Label6.Caption)
Form1.t = t - 1

GoTo 10

End If

If p = 3 Then
    Form1.Label3.Caption = ""
    Form1.Text19.Text = Val(Form1.Label92.Caption) + Val(Form1.Label2.Caption) + Val(Form1.Label3.Caption) + Val(Form1.Label4.Caption) + Val(Form1.Label5.Caption) + Val(Form1.Label6.Caption)
Form1.t = t - 1

GoTo 10

End If

If p = 4 Then
    Form1.Label4.Caption = ""
    Form1.Text19.Text = Val(Form1.Label92.Caption) + Val(Form1.Label2.Caption) + Val(Form1.Label3.Caption) + Val(Form1.Label4.Caption) + Val(Form1.Label5.Caption) + Val(Form1.Label6.Caption)
Form1.t = t - 1
GoTo 10

End If

If p = 5 Then
    Form1.Label5.Caption = ""
    Form1.Text19.Text = Val(Form1.Label92.Caption) + Val(Form1.Label2.Caption) + Val(Form1.Label3.Caption) + Val(Form1.Label4.Caption) + Val(Form1.Label5.Caption) + Val(Form1.Label6.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 6 Then
    Form1.Label6.Caption = ""
    Form1.Text19.Text = Val(Form1.Label92.Caption) + Val(Form1.Label2.Caption) + Val(Form1.Label3.Caption) + Val(Form1.Label4.Caption) + Val(Form1.Label5.Caption) + Val(Form1.Label6.Caption)
Form1.t = t - 1
GoTo 10

End If
'************************
If p = 7 Then
    Form1.Label7.Caption = ""
Form1.Text41.Text = 0
Form1.t = t - 1
GoTo 10

End If
If p = 8 Then
    Form1.Label8.Caption = ""
Form1.Text41.Text = 0
Form1.t = t - 1
GoTo 10

End If
'========================
If p = 9 Then
    Form1.Label9.Caption = ""
    Form1.Text120.Text = Val(Form1.Label9.Caption) + Val(Form1.Label10.Caption) + Val(Form1.Label11.Caption) + Val(Form1.Label12.Caption) + Val(Form1.Label13.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 10 Then
    Form1.Label10.Caption = ""
    Form1.Text120.Text = Val(Form1.Label9.Caption) + Val(Form1.Label10.Caption) + Val(Form1.Label11.Caption) + Val(Form1.Label12.Caption) + Val(Form1.Label13.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 11 Then
    Form1.Label11.Caption = ""
    Form1.Text120.Text = Val(Form1.Label9.Caption) + Val(Form1.Label10.Caption) + Val(Form1.Label11.Caption) + Val(Form1.Label12.Caption) + Val(Form1.Label13.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 12 Then
    Form1.Label12.Caption = ""
    Form1.Text120.Text = Val(Form1.Label9.Caption) + Val(Form1.Label10.Caption) + Val(Form1.Label11.Caption) + Val(Form1.Label12.Caption) + Val(Form1.Label13.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 13 Then
    Form1.Label13.Caption = ""
    Form1.Text120.Text = Val(Form1.Label9.Caption) + Val(Form1.Label10.Caption) + Val(Form1.Label11.Caption) + Val(Form1.Label12.Caption) + Val(Form1.Label13.Caption)
Form1.t = t - 1
GoTo 10

End If
'\\\\\\\\\\\\\\\\\\\\\\\\\



If p = 14 Then
    Form1.Label14.Caption = ""
    Form1.Text20.Text = Val(Form1.Label14.Caption) + Val(Form1.Label15.Caption) + Val(Form1.Label16.Caption) + Val(Form1.Label17.Caption) + Val(Form1.Label18.Caption) + Val(Form1.Label19.Caption)

    Form1.Text42.Text = 0
Form1.t = t - 1
GoTo 10

End If

If p = 15 Then
    Form1.Label15.Caption = ""
    Form1.Text20.Text = Val(Form1.Label14.Caption) + Val(Form1.Label15.Caption) + Val(Form1.Label16.Caption) + Val(Form1.Label17.Caption) + Val(Form1.Label18.Caption) + Val(Form1.Label19.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 16 Then
    Form1.Label16.Caption = ""
    Form1.Text20.Text = Val(Form1.Label14.Caption) + Val(Form1.Label15.Caption) + Val(Form1.Label16.Caption) + Val(Form1.Label17.Caption) + Val(Form1.Label18.Caption) + Val(Form1.Label19.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 17 Then
    Form1.Label17.Caption = ""
    Form1.Text20.Text = Val(Form1.Label14.Caption) + Val(Form1.Label15.Caption) + Val(Form1.Label16.Caption) + Val(Form1.Label17.Caption) + Val(Form1.Label18.Caption) + Val(Form1.Label19.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 18 Then
    Form1.Label18.Caption = ""
    Form1.Text20.Text = Val(Form1.Label14.Caption) + Val(Form1.Label15.Caption) + Val(Form1.Label16.Caption) + Val(Form1.Label17.Caption) + Val(Form1.Label18.Caption) + Val(Form1.Label19.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 19 Then
    Form1.Label19.Caption = ""
    Form1.Text20.Text = Val(Form1.Label14.Caption) + Val(Form1.Label15.Caption) + Val(Form1.Label16.Caption) + Val(Form1.Label17.Caption) + Val(Form1.Label18.Caption) + Val(Form1.Label19.Caption)
Form1.t = t - 1
GoTo 10

End If
'*******************************
If p = 20 Then
    Form1.Label20.Caption = ""
Form1.Text42.Text = 0
Form1.t = t - 1
GoTo 10

End If
If p = 21 Then
    Form1.Label21.Caption = ""
Form1.Text42.Text = 0
Form1.t = t - 1
GoTo 10

End If

'==============================
If p = 22 Then
    Form1.Label22.Caption = ""
    Form1.Text121.Text = Val(Form1.Label22.Caption) + Val(Form1.Label23.Caption) + Val(Form1.Label24.Caption) + Val(Form1.Label25.Caption) + Val(Form1.Label26.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 23 Then
    Form1.Label23.Caption = ""
    Form1.Text121.Text = Val(Form1.Label22.Caption) + Val(Form1.Label23.Caption) + Val(Form1.Label24.Caption) + Val(Form1.Label25.Caption) + Val(Form1.Label26.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 24 Then
    Form1.Label24.Caption = ""
    Form1.Text121.Text = Val(Form1.Label22.Caption) + Val(Form1.Label23.Caption) + Val(Form1.Label24.Caption) + Val(Form1.Label25.Caption) + Val(Form1.Label26.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 25 Then
    Form1.Label25.Caption = ""
    Form1.Text121.Text = Val(Form1.Label22.Caption) + Val(Form1.Label23.Caption) + Val(Form1.Label24.Caption) + Val(Form1.Label25.Caption) + Val(Form1.Label26.Caption)
Form1.t = t - 1
GoTo 10

End If
If p = 26 Then
    Form1.Label26.Caption = ""
    Form1.Text121.Text = Val(Form1.Label22.Caption) + Val(Form1.Label23.Caption) + Val(Form1.Label24.Caption) + Val(Form1.Label25.Caption) + Val(Form1.Label26.Caption)
Form1.t = t - 1
GoTo 10

End If
'//////////////////




If p = 27 Then
    Form1.Label27.Caption = ""
    Form1.Text21.Text = Val(Form1.Label27.Caption) + Val(Form1.Label28.Caption) + Val(Form1.Label29.Caption) + Val(Form1.Label30.Caption) + Val(Form1.Label31.Caption) + Val(Form1.Label32.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 28 Then
    Form1.Label28.Caption = ""
    Form1.Text21.Text = Val(Form1.Label27.Caption) + Val(Form1.Label28.Caption) + Val(Form1.Label29.Caption) + Val(Form1.Label30.Caption) + Val(Form1.Label31.Caption) + Val(Form1.Label32.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 29 Then
    Form1.Label29.Caption = ""
    Form1.Text21.Text = Val(Form1.Label27.Caption) + Val(Form1.Label28.Caption) + Val(Form1.Label29.Caption) + Val(Form1.Label30.Caption) + Val(Form1.Label31.Caption) + Val(Form1.Label32.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 30 Then
    Form1.Label30.Caption = ""
    Form1.Text21.Text = Val(Form1.Label27.Caption) + Val(Form1.Label28.Caption) + Val(Form1.Label29.Caption) + Val(Form1.Label30.Caption) + Val(Form1.Label31.Caption) + Val(Form1.Label32.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 31 Then
    Form1.Label31.Caption = ""
    Form1.Text21.Text = Val(Form1.Label27.Caption) + Val(Form1.Label28.Caption) + Val(Form1.Label29.Caption) + Val(Form1.Label30.Caption) + Val(Form1.Label31.Caption) + Val(Form1.Label32.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 32 Then
    Form1.Label32.Caption = ""
    Form1.Text21.Text = Val(Form1.Label27.Caption) + Val(Form1.Label28.Caption) + Val(Form1.Label29.Caption) + Val(Form1.Label30.Caption) + Val(Form1.Label31.Caption) + Val(Form1.Label32.Caption)
Form1.t = t - 1
GoTo 10

End If
'*****************
If p = 33 Then
    Form1.Label33.Caption = ""
    Form1.Text41.Text = 0
Form1.t = t - 1
GoTo 10

End If
If p = 34 Then
    Form1.Label34.Caption = ""
    Form1.Text41.Text = 0

Form1.t = t - 1
GoTo 10

End If
'==================
If p = 35 Then
    Form1.Label35.Caption = ""
Form1.Text122.Text = Val(Form1.Label35.Caption) + Val(Form1.Label36.Caption) + Val(Form1.Label37.Caption) + Val(Form1.Label38.Caption) + Val(Form1.Label39.Caption)

Form1.t = t - 1
GoTo 10

End If
If p = 36 Then
    Form1.Label36.Caption = ""
Form1.Text122.Text = Val(Form1.Label35.Caption) + Val(Form1.Label36.Caption) + Val(Form1.Label37.Caption) + Val(Form1.Label38.Caption) + Val(Form1.Label39.Caption)

Form1.t = t - 1
GoTo 10

End If
If p = 37 Then
    Form1.Label37.Caption = ""
Form1.Text122.Text = Val(Form1.Label35.Caption) + Val(Form1.Label36.Caption) + Val(Form1.Label37.Caption) + Val(Form1.Label38.Caption) + Val(Form1.Label39.Caption)

Form1.t = t - 1
GoTo 10

End If
If p = 38 Then
    Form1.Label38.Caption = ""
Form1.Text122.Text = Val(Form1.Label35.Caption) + Val(Form1.Label36.Caption) + Val(Form1.Label37.Caption) + Val(Form1.Label38.Caption) + Val(Form1.Label39.Caption)

Form1.t = t - 1
GoTo 10

End If
If p = 39 Then
    Form1.Label39.Caption = ""
Form1.Text122.Text = Val(Form1.Label35.Caption) + Val(Form1.Label36.Caption) + Val(Form1.Label37.Caption) + Val(Form1.Label38.Caption) + Val(Form1.Label39.Caption)

Form1.t = t - 1
GoTo 10

End If
'//////////////////////
If p = 40 Then
    Form1.Label40.Caption = ""
Form1.Text22.Text = Val(Form1.Label40.Caption) + Val(Form1.Label41.Caption) + Val(Form1.Label42.Caption) + Val(Form1.Label43.Caption) + Val(Form1.Label44.Caption) + Val(Form1.Label45.Caption)
Form1.Text44.Text = 0
Form1.t = t - 1
GoTo 10

End If

If p = 41 Then
    Form1.Label41.Caption = ""
Form1.Text22.Text = Val(Form1.Label40.Caption) + Val(Form1.Label41.Caption) + Val(Form1.Label42.Caption) + Val(Form1.Label43.Caption) + Val(Form1.Label44.Caption) + Val(Form1.Label45.Caption)
Form1.t = t - 1
GoTo 10

End If

If p = 42 Then
    Form1.Label42.Caption = ""
Form1.Text22.Text = Val(Form1.Label40.Caption) + Val(Form1.Label41.Caption) + Val(Form1.Label42.Caption) + Val(Form1.Label43.Caption) + Val(Form1.Label44.Caption) + Val(Form1.Label45.Caption)
Form1.t = t - 1
GoTo 10

End If

If p = 43 Then
    Form1.Label43.Caption = ""
Form1.Text22.Text = Val(Form1.Label40.Caption) + Val(Form1.Label41.Caption) + Val(Form1.Label42.Caption) + Val(Form1.Label43.Caption) + Val(Form1.Label44.Caption) + Val(Form1.Label45.Caption)
Form1.t = t - 1
GoTo 10

End If

If p = 44 Then
    Form1.Label44.Caption = ""
Form1.Text22.Text = Val(Form1.Label40.Caption) + Val(Form1.Label41.Caption) + Val(Form1.Label42.Caption) + Val(Form1.Label43.Caption) + Val(Form1.Label44.Caption) + Val(Form1.Label45.Caption)
Form1.t = t - 1
GoTo 10

End If

If p = 45 Then
    Form1.Label45.Caption = ""
Form1.Text22.Text = Val(Form1.Label40.Caption) + Val(Form1.Label41.Caption) + Val(Form1.Label42.Caption) + Val(Form1.Label43.Caption) + Val(Form1.Label44.Caption) + Val(Form1.Label45.Caption)
Form1.t = t - 1
GoTo 10

End If
'***************************

If p = 46 Then
    Form1.Label46.Caption = ""
    Form1.Text44.Text = 0
Form1.t = t - 1
GoTo 10
End If
If p = 47 Then
    Form1.Label47.Caption = ""
    Form1.Text44.Text = 0
Form1.t = t - 1
GoTo 10
End If
'========================

If p = 48 Then
    Form1.Label48.Caption = ""
    Form1.Text123.Text = Val(Form1.Label48.Caption) + Val(Form1.Label49.Caption) + Val(Form1.Label50.Caption) + Val(Form1.Label51.Caption) + Val(Form1.Label52.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 49 Then
    Form1.Label49.Caption = ""
    Form1.Text123.Text = Val(Form1.Label48.Caption) + Val(Form1.Label49.Caption) + Val(Form1.Label50.Caption) + Val(Form1.Label51.Caption) + Val(Form1.Label52.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 50 Then
    Form1.Label50.Caption = ""
    Form1.Text123.Text = Val(Form1.Label48.Caption) + Val(Form1.Label49.Caption) + Val(Form1.Label50.Caption) + Val(Form1.Label51.Caption) + Val(Form1.Label52.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 51 Then
    Form1.Label51.Caption = ""
    Form1.Text123.Text = Val(Form1.Label48.Caption) + Val(Form1.Label49.Caption) + Val(Form1.Label50.Caption) + Val(Form1.Label51.Caption) + Val(Form1.Label52.Caption)
Form1.t = t - 1
GoTo 10
End If
If p = 52 Then
    Form1.Label52.Caption = ""
    Form1.Text123.Text = Val(Form1.Label48.Caption) + Val(Form1.Label49.Caption) + Val(Form1.Label50.Caption) + Val(Form1.Label51.Caption) + Val(Form1.Label52.Caption)
Form1.t = t - 1
GoTo 10
End If
'///////////////////////
'************
If p = 53 Then
    Form1.Label53.Caption = ""
    Form1.Text23.Text = Val(Form1.Label53.Caption) + Val(Form1.Label54.Caption) + Val(Form1.Label55.Caption) + Val(Form1.Label56.Caption) + Val(Form1.Label57.Caption) + Val(Form1.Label58.Caption)
    Form1.Text45.Text = 0
Form1.t = t - 1
GoTo 10
End If

If p = 54 Then
    Form1.Label54.Caption = ""
    Form1.Text23.Text = Val(Form1.Label53.Caption) + Val(Form1.Label54.Caption) + Val(Form1.Label55.Caption) + Val(Form1.Label56.Caption) + Val(Form1.Label57.Caption) + Val(Form1.Label58.Caption)
Form1.t = t - 1
  GoTo 10
  
End If

If p = 55 Then
    Form1.Label55.Caption = ""
    Form1.Text23.Text = Val(Form1.Label53.Caption) + Val(Form1.Label54.Caption) + Val(Form1.Label55.Caption) + Val(Form1.Label56.Caption) + Val(Form1.Label57.Caption) + Val(Form1.Label58.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 56 Then
    Form1.Label56.Caption = ""
    Form1.Text23.Text = Val(Form1.Label53.Caption) + Val(Form1.Label54.Caption) + Val(Form1.Label55.Caption) + Val(Form1.Label56.Caption) + Val(Form1.Label57.Caption) + Val(Form1.Label58.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 57 Then
    Form1.Label57.Caption = ""
    Form1.Text23.Text = Val(Form1.Label53.Caption) + Val(Form1.Label54.Caption) + Val(Form1.Label55.Caption) + Val(Form1.Label56.Caption) + Val(Form1.Label57.Caption) + Val(Form1.Label58.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 58 Then
    Form1.Label58.Caption = ""
    Form1.Text23.Text = Val(Form1.Label53.Caption) + Val(Form1.Label54.Caption) + Val(Form1.Label55.Caption) + Val(Form1.Label56.Caption) + Val(Form1.Label57.Caption) + Val(Form1.Label58.Caption)
Form1.t = t - 1
GoTo 10
End If
'****************
If p = 59 Then
    Form1.Label59.Caption = ""
    Form1.Text45.Text = 0
Form1.t = t - 1
GoTo 10
End If

If p = 60 Then
    Form1.Label60.Caption = ""
    Form1.Text45.Text = 0
Form1.t = t - 1
GoTo 10
End If
'========================
If p = 61 Then
    Form1.Label61.Caption = ""
    Form1.Text124.Text = Val(Form1.Label61.Caption) + Val(Form1.Label62.Caption) + Val(Form1.Label63.Caption) + Val(Form1.Label64.Caption) + Val(Form1.Label65.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 62 Then
    Form1.Label62.Caption = ""
    Form1.Text124.Text = Val(Form1.Label61.Caption) + Val(Form1.Label62.Caption) + Val(Form1.Label63.Caption) + Val(Form1.Label64.Caption) + Val(Form1.Label65.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 63 Then
    Form1.Label63.Caption = ""
    Form1.Text124.Text = Val(Form1.Label61.Caption) + Val(Form1.Label62.Caption) + Val(Form1.Label63.Caption) + Val(Form1.Label64.Caption) + Val(Form1.Label65.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 64 Then
    Form1.Label64.Caption = ""
    Form1.Text124.Text = Val(Form1.Label61.Caption) + Val(Form1.Label62.Caption) + Val(Form1.Label63.Caption) + Val(Form1.Label64.Caption) + Val(Form1.Label65.Caption)
Form1.t = t - 1
GoTo 10
End If

If p = 65 Then
    Form1.Label65.Caption = ""
    Form1.Text124.Text = Val(Form1.Label61.Caption) + Val(Form1.Label62.Caption) + Val(Form1.Label63.Caption) + Val(Form1.Label64.Caption) + Val(Form1.Label65.Caption)
    'Form1.a = Pom + 1
Form1.t = t - 1
GoTo 10
End If
'************************
If p = 66 Then
    Form1.Label66.Caption = ""
    Form1.Text24.Text = Val(Form1.Label66.Caption) + Val(Form1.Label67.Caption) + Val(Form1.Label68.Caption) + Val(Form1.Label69.Caption) + Val(Form1.Label70.Caption) + Val(Form1.Label71.Caption)
    Form1.Text46.Text = 0
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If

If p = 67 Then
    Form1.Label67.Caption = ""
    Form1.Text24.Text = Val(Form1.Label66.Caption) + Val(Form1.Label67.Caption) + Val(Form1.Label68.Caption) + Val(Form1.Label69.Caption) + Val(Form1.Label70.Caption) + Val(Form1.Label71.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If

If p = 68 Then
    Form1.Label68.Caption = ""
    Form1.Text24.Text = Val(Form1.Label66.Caption) + Val(Form1.Label67.Caption) + Val(Form1.Label68.Caption) + Val(Form1.Label69.Caption) + Val(Form1.Label70.Caption) + Val(Form1.Label71.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If

If p = 69 Then
    Form1.Label69.Caption = ""
    Form1.Text24.Text = Val(Form1.Label66.Caption) + Val(Form1.Label67.Caption) + Val(Form1.Label68.Caption) + Val(Form1.Label69.Caption) + Val(Form1.Label70.Caption) + Val(Form1.Label71.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If

If p = 70 Then
    Form1.Label70.Caption = ""
    Form1.Text24.Text = Val(Form1.Label66.Caption) + Val(Form1.Label67.Caption) + Val(Form1.Label68.Caption) + Val(Form1.Label69.Caption) + Val(Form1.Label70.Caption) + Val(Form1.Label71.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If
If p = 71 Then
    Form1.Label71.Caption = ""
    Form1.Text24.Text = Val(Form1.Label66.Caption) + Val(Form1.Label67.Caption) + Val(Form1.Label68.Caption) + Val(Form1.Label69.Caption) + Val(Form1.Label70.Caption) + Val(Form1.Label71.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If
'******************
If p = 72 Then
    Form1.Label72.Caption = ""
    Form1.Text46.Text = 0
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If
If p = 73 Then
    Form1.Label73.Caption = ""
    Form1.Text46.Text = 0
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If
'========================
If p = 74 Then
    Form1.Label74.Caption = ""
    Form1.Text125.Text = Val(Form1.Label74.Caption) + Val(Form1.Label75.Caption) + Val(Form1.Label76.Caption) + Val(Form1.Label77.Caption) + Val(Form1.Label78.Caption)
    Form1.Sn.ForeColor = "&h0000ff00"
    blok.blok
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok

Form1.t = t - 1
GoTo 10
End If
If p = 75 Then
    Form1.Label75.Caption = ""
    Form1.Text125.Text = Val(Form1.Label74.Caption) + Val(Form1.Label75.Caption) + Val(Form1.Label76.Caption) + Val(Form1.Label77.Caption) + Val(Form1.Label78.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok
Form1.t = t - 1
GoTo 10
End If
If p = 76 Then
    Form1.Label76.Caption = ""
    Form1.Text125.Text = Val(Form1.Label74.Caption) + Val(Form1.Label75.Caption) + Val(Form1.Label76.Caption) + Val(Form1.Label77.Caption) + Val(Form1.Label78.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok
Form1.t = t - 1
GoTo 10
End If
If p = 77 Then
    Form1.Label77.Caption = ""
    Form1.Text125.Text = Val(Form1.Label74.Caption) + Val(Form1.Label75.Caption) + Val(Form1.Label76.Caption) + Val(Form1.Label77.Caption) + Val(Form1.Label78.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok
Form1.t = t - 1
GoTo 10
End If
If p = 78 Then
    Form1.Label78.Caption = ""
    Form1.Text125.Text = Val(Form1.Label74.Caption) + Val(Form1.Label75.Caption) + Val(Form1.Label76.Caption) + Val(Form1.Label77.Caption) + Val(Form1.Label78.Caption)
    Form1.Sn.ForeColor = "&h000000ff"
    blok.blok
Form1.t = t - 1
GoTo 10
End If

'**********************
If p = 79 Or p = 80 Or p = 81 Or p = 82 Or p = 83 Or p = 84 Or p = 85 Or p = 86 Or p = 87 Or p = 88 Or p = 89 Or p = 90 Or p = 91 Then
Form1.Text8.Text = ""
Form1.Text9.Text = ""
Form1.Text10.Text = ""
Form1.Text11.Text = ""
Form1.Text12.Text = ""
Form1.a = 0
blok.deb
blok.DENBL
MsgBox ("U NAJAVI morate da igrate u polje koje ste najavili")
End If



10:
End Sub
