Attribute VB_Name = "SnimPoz"
Public Sub Snimanje()
Dim Sry As Integer, Bry As Integer
Dim Jil As Integer, Kir As String

Jil = FreeFile
If Right$(App.Path, 1) <> "\" Then
Open App.Path & "\" & "Poza.lex" For Output As #Jil
Else
Open App.Path & "Poza.lex" For Output As #Jil
End If

'/ Upisivanje u fajl
Print #Jil, Form1.Label2.Caption
Print #Jil, Form1.Label3.Caption
Print #Jil, Form1.Label4.Caption
Print #Jil, Form1.Label5.Caption
Print #Jil, Form1.Label6.Caption
Print #Jil, Form1.Label7.Caption
Print #Jil, Form1.Label8.Caption
Print #Jil, Form1.Label9.Caption
Print #Jil, Form1.Label10.Caption
Print #Jil, Form1.Label11.Caption
Print #Jil, Form1.Label12.Caption
Print #Jil, Form1.Label13.Caption
Print #Jil, Form1.Label14.Caption
Print #Jil, Form1.Label15.Caption
Print #Jil, Form1.Label16.Caption
Print #Jil, Form1.Label17.Caption
Print #Jil, Form1.Label18.Caption
Print #Jil, Form1.Label19.Caption
Print #Jil, Form1.Label20.Caption
Print #Jil, Form1.Label21.Caption
Print #Jil, Form1.Label22.Caption
Print #Jil, Form1.Label23.Caption
Print #Jil, Form1.Label24.Caption
Print #Jil, Form1.Label25.Caption
Print #Jil, Form1.Label26.Caption
Print #Jil, Form1.Label27.Caption
Print #Jil, Form1.Label28.Caption
Print #Jil, Form1.Label29.Caption
Print #Jil, Form1.Label30.Caption
Print #Jil, Form1.Label31.Caption
Print #Jil, Form1.Label32.Caption
Print #Jil, Form1.Label33.Caption
Print #Jil, Form1.Label34.Caption
Print #Jil, Form1.Label35.Caption
Print #Jil, Form1.Label36.Caption
Print #Jil, Form1.Label37.Caption
Print #Jil, Form1.Label38.Caption
Print #Jil, Form1.Label39.Caption
Print #Jil, Form1.Label40.Caption
Print #Jil, Form1.Label41.Caption
Print #Jil, Form1.Label42.Caption
Print #Jil, Form1.Label43.Caption
Print #Jil, Form1.Label44.Caption
Print #Jil, Form1.Label45.Caption
Print #Jil, Form1.Label46.Caption
Print #Jil, Form1.Label47.Caption
Print #Jil, Form1.Label48.Caption
Print #Jil, Form1.Label49.Caption
Print #Jil, Form1.Label50.Caption
Print #Jil, Form1.Label51.Caption
Print #Jil, Form1.Label52.Caption
Print #Jil, Form1.Label53.Caption
Print #Jil, Form1.Label54.Caption
Print #Jil, Form1.Label55.Caption
Print #Jil, Form1.Label56.Caption
Print #Jil, Form1.Label57.Caption
Print #Jil, Form1.Label58.Caption
Print #Jil, Form1.Label59.Caption
Print #Jil, Form1.Label60.Caption
Print #Jil, Form1.Label61.Caption
Print #Jil, Form1.Label62.Caption
Print #Jil, Form1.Label63.Caption
Print #Jil, Form1.Label64.Caption
Print #Jil, Form1.Label65.Caption
Print #Jil, Form1.Label66.Caption
Print #Jil, Form1.Label67.Caption
Print #Jil, Form1.Label68.Caption
Print #Jil, Form1.Label69.Caption
Print #Jil, Form1.Label70.Caption
Print #Jil, Form1.Label71.Caption
Print #Jil, Form1.Label72.Caption
Print #Jil, Form1.Label73.Caption
Print #Jil, Form1.Label74.Caption
Print #Jil, Form1.Label75.Caption
Print #Jil, Form1.Label76.Caption
Print #Jil, Form1.Label77.Caption
Print #Jil, Form1.Label78.Caption
Print #Jil, Form1.Label79.Caption
Print #Jil, Form1.Label80.Caption
Print #Jil, Form1.Label81.Caption
Print #Jil, Form1.Label82.Caption
Print #Jil, Form1.Label83.Caption
Print #Jil, Form1.Label84.Caption
Print #Jil, Form1.Label85.Caption
Print #Jil, Form1.Label86.Caption
Print #Jil, Form1.Label87.Caption
Print #Jil, Form1.Label88.Caption
Print #Jil, Form1.Label89.Caption
Print #Jil, Form1.Label90.Caption
Print #Jil, Form1.Label91.Caption
Print #Jil, Form1.Label92.Caption
'/ Odabrane
Print #Jil, Form1.Text8.Text
Print #Jil, Form1.Text9.Text
Print #Jil, Form1.Text10.Text
Print #Jil, Form1.Text11.Text
Print #Jil, Form1.Text12.Text
'/ Preostale
Print #Jil, Form1.Text13.Text
Print #Jil, Form1.Text14.Text
Print #Jil, Form1.Text15.Text
Print #Jil, Form1.Text16.Text
Print #Jil, Form1.Text17.Text
Print #Jil, Form1.Text18.Text
'/ Prvi rezultati
Print #Jil, Form1.Text19.Text
Print #Jil, Form1.Text20.Text
Print #Jil, Form1.Text21.Text
Print #Jil, Form1.Text22.Text
Print #Jil, Form1.Text23.Text
Print #Jil, Form1.Text24.Text
Print #Jil, Form1.Text25.Text
Print #Jil, Form1.Text26.Text
'/ Drugi rezultati
Print #Jil, Form1.Text41.Text
Print #Jil, Form1.Text42.Text
Print #Jil, Form1.Text43.Text
Print #Jil, Form1.Text44.Text
Print #Jil, Form1.Text45.Text
Print #Jil, Form1.Text46.Text
Print #Jil, Form1.Text47.Text
Print #Jil, Form1.Text48.Text
'/ Treci rezultati
Print #Jil, Form1.Text120.Text
Print #Jil, Form1.Text121.Text
Print #Jil, Form1.Text122.Text
Print #Jil, Form1.Text123.Text
Print #Jil, Form1.Text124.Text
Print #Jil, Form1.Text125.Text
Print #Jil, Form1.Text126.Text
Print #Jil, Form1.Text127.Text
'/ Glavni rezultat
Print #Jil, Form1.Text119.Text
'/ Broj preostalih polja
Print #Jil, Form1.t
'/Broj bacanja
If Form1.Bro.Caption >= 3 Then Form1.Bro.Caption = 0
Print #Jil, Form1.Bro.Caption
Close #Jil
MsgBox "Snimljeno...", vbInformation, "Obavestenje"

End Sub

Public Sub Citanje()
On Error Resume Next
If Form1.t <> 0 Then
If MsgBox("Zelite da nastavite zadnju snimljenu igru? Ovim se prekida igra koja je u toku!", vbYesNo, "Provera") = vbNo Then Exit Sub
End If
Dim Pol(1 To 129) As String
Dim Ada As Integer, FrrF As Integer
Ng.Ng
'/Otvaranje fajla
FrrF = FreeFile
If Right$(App.Path, 1) = "\" Then
Open App.Path & "Poza.lex" For Input As #FrrF
Else
Open App.Path & "\" & "Poza.lex" For Input As #FrrF
End If

'/Citanje iz fajla
For Ada = 1 To 129
Input #FrrF, Pol(Ada)
Next Ada
Close #FrrF

'/Upisivanje u polaja
With Form1
 .Label2.Caption = Pol(1)
 .Label3.Caption = Pol(2)
 .Label4.Caption = Pol(3)
 .Label5.Caption = Pol(4)
 .Label6.Caption = Pol(5)
 .Label7.Caption = Pol(6)
 .Label8.Caption = Pol(7)
 .Label9.Caption = Pol(8)
 .Label10.Caption = Pol(9)
 .Label11.Caption = Pol(10)
 .Label12.Caption = Pol(11)
 .Label13.Caption = Pol(12)
 .Label14.Caption = Pol(13)
 .Label15.Caption = Pol(14)
 .Label16.Caption = Pol(15)
 .Label17.Caption = Pol(16)
 .Label18.Caption = Pol(17)
 .Label19.Caption = Pol(18)
 .Label20.Caption = Pol(19)
 .Label21.Caption = Pol(20)
 .Label22.Caption = Pol(21)
 .Label23.Caption = Pol(22)
 .Label24.Caption = Pol(23)
 .Label25.Caption = Pol(24)
 .Label26.Caption = Pol(25)
 .Label27.Caption = Pol(26)
 .Label28.Caption = Pol(27)
 .Label29.Caption = Pol(28)
 .Label30.Caption = Pol(29)
 .Label31.Caption = Pol(30)
 .Label32.Caption = Pol(31)
 .Label33.Caption = Pol(32)
 .Label34.Caption = Pol(33)
 .Label35.Caption = Pol(34)
 .Label36.Caption = Pol(35)
 .Label37.Caption = Pol(36)
 .Label38.Caption = Pol(37)
 .Label39.Caption = Pol(38)
 .Label40.Caption = Pol(39)
 .Label41.Caption = Pol(40)
 .Label42.Caption = Pol(41)
 .Label43.Caption = Pol(42)
 .Label44.Caption = Pol(43)
 .Label45.Caption = Pol(44)
 .Label46.Caption = Pol(45)
 .Label47.Caption = Pol(46)
 .Label48.Caption = Pol(47)
 .Label49.Caption = Pol(48)
 .Label50.Caption = Pol(49)
 .Label51.Caption = Pol(50)
 .Label52.Caption = Pol(51)
 .Label53.Caption = Pol(52)
 .Label54.Caption = Pol(53)
 .Label55.Caption = Pol(54)
 .Label56.Caption = Pol(55)
 .Label57.Caption = Pol(56)
 .Label58.Caption = Pol(57)
 .Label59.Caption = Pol(58)
 .Label60.Caption = Pol(59)
 .Label61.Caption = Pol(60)
 .Label62.Caption = Pol(61)
 .Label63.Caption = Pol(62)
 .Label64.Caption = Pol(63)
 .Label65.Caption = Pol(64)
 .Label66.Caption = Pol(65)
 .Label67.Caption = Pol(66)
 .Label68.Caption = Pol(67)
 .Label69.Caption = Pol(68)
 .Label70.Caption = Pol(69)
 .Label71.Caption = Pol(70)
 .Label72.Caption = Pol(71)
 .Label73.Caption = Pol(72)
 .Label74.Caption = Pol(73)
 .Label75.Caption = Pol(74)
 .Label76.Caption = Pol(75)
 .Label77.Caption = Pol(76)
 .Label78.Caption = Pol(77)
 .Label79.Caption = Pol(78)
 .Label80.Caption = Pol(79)
 .Label81.Caption = Pol(80)
 .Label82.Caption = Pol(81)
 .Label83.Caption = Pol(82)
 .Label84.Caption = Pol(83)
 .Label85.Caption = Pol(84)
 .Label86.Caption = Pol(85)
 .Label87.Caption = Pol(86)
 .Label88.Caption = Pol(87)
 .Label89.Caption = Pol(88)
 .Label90.Caption = Pol(89)
 .Label91.Caption = Pol(90)
 .Label92.Caption = Pol(91)
'/ Odabrane
 .Text8.Text = Pol(92)
 .Text9.Text = Pol(93)
 .Text10.Text = Pol(94)
 .Text11.Text = Pol(95)
 .Text12.Text = Pol(96)
'/ Preostale
 .Text13.Text = Pol(97)
 .Text14.Text = Pol(98)
 .Text15.Text = Pol(99)
 .Text16.Text = Pol(100)
 .Text17.Text = Pol(101)
 .Text18.Text = Pol(102)
'/ Prvi rezultati
 .Text19.Text = Pol(103)
 .Text20.Text = Pol(104)
 .Text21.Text = Pol(105)
 .Text22.Text = Pol(106)
 .Text23.Text = Pol(107)
 .Text24.Text = Pol(108)
 .Text25.Text = Pol(109)
 .Text26.Text = Pol(110)
'/ Drugi rezultati
 .Text41.Text = Pol(111)
 .Text42.Text = Pol(112)
 .Text43.Text = Pol(113)
 .Text44.Text = Pol(114)
 .Text45.Text = Pol(115)
 .Text46.Text = Pol(116)
 .Text47.Text = Pol(117)
 .Text48.Text = Pol(118)
'/ Treci rezultati
 .Text120.Text = Pol(119)
 .Text121.Text = Pol(120)
 .Text122.Text = Pol(121)
 .Text123.Text = Pol(122)
 .Text124.Text = Pol(123)
 .Text125.Text = Pol(124)
 .Text126.Text = Pol(125)
 .Text127.Text = Pol(126)
'/ Glavni rezultat
 .Text119.Text = Pol(127)
'/ Broj preostalih polja
 .t = Val(Pol(128))
 .Bro.Caption = Pol(129)
 '/Broj poteza
 .a = Val(Pol(129))

End With

End Sub
