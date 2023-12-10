VERSION 5.00
Begin VB.Form Adios 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REZULTAT"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4740
   ForeColor       =   &H00400000&
   Icon            =   "Adios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SuperNagrada "
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nagrada "
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rekord"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BRAVO!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Vas    rezultat   je:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Adios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()


Unload Form1
Set Form1 = Nothing
Unload Adios
Set Adios = Nothing
End

End Sub

Private Sub Command2_Click()
MsgBox "Da biste u buduce mogli da vidite nagradu, pritisnite F4 i ukucajte sifru: zdenka ."
Unload Form1
Set Form1 = Nothing
Unload Adios
Set Adios = Nothing


Load Nagrada
Nagrada.Show

End Sub

Private Sub Command3_Click()
MsgBox "Kad pozelite da vidite nagradu pritisnite F4 i ukucajte sifru: zdenkas , ili: cicamica ."
Unload Form1
Set Form1 = Nothing
Unload Adios
Set Adios = Nothing


Load SuperNagrada
SuperNagrada.Show

End Sub

Private Sub Form_Load()
Dim Rez As String, sl As Integer, Zez As Integer
Dim Postoi As String
 Poruke.mAplauz
sl = FreeFile
Poruke.mAplauz
Postoi = Dir$(Form1.Path)
While Postoi <> ""
If Postoi = "Res.rec" Then GoTo 2
Postoi = Dir$
Wend
GoTo 3

2: Open Form1.Path & "Res.rec" For Input As #sl
Line Input #sl, Rez
Close #sl

3: Text2.Text = Rez
 Zez = Val(Form1.Text119.Text)
If Zez > Rez Then
    Open Form1.Path & "Res.rec" For Output As #sl
    Print #sl, Val(Zez)
    Close #sl
    If Val(Zez) > 270 Then MsgBox ("Postignut je novi rekord!")
    Command2.Visible = True
    Command3.Visible = True
    If Zez > 3300 Then
    Command3.Enabled = True
    End If
    
End If


Text1.Text = Form1.Text119
End Sub

