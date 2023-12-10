VERSION 5.00
Begin VB.Form Odabir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muzika"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Odabir.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PLAY"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   4080
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   3450
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Odabir.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Odabir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Puta As String ' Cela putanja sa imenom
Public Nam As String  ' Ime fajla
Public PathANJA As String 'Putanja do muzickog fajla

Public Sub Command1_Click()
If Right$(PathANJA, 1) <> "\" Then
Puta = PathANJA & "\" & Nam
Else
Puta = PathANJA & Nam
End If
If UCase(Right$(Puta, 3)) <> "WAV" Then MsgBox "Neodgovarajuci fajl! Odaberite fajl sa extenzijom '.wav'", vbCritical, "Oops!": Exit Sub
Poruke.Muz (Puta)
Unload Odabir
Form1.tis.Enabled = True
Set Odabir = Nothing
End Sub

Private Sub Command2_Click()
Unload Odabir
Set Odabir = Nothing
End Sub

Private Sub Dir1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
File1.Path = Dir1.List(Dir1.ListIndex)
PathANJA = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
On Error GoTo 332
Dir1.Path = Drive1.Drive
File1.Path = Dir1.List(Dir1.ListIndex)
PathANJA = Dir1.List(Dir1.ListIndex)
Exit Sub
332:
Drive1.Drive = Left$(App.Path, 3)
File1.Path = Dir1.List(Dir1.ListIndex)
PathANJA = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub File1_Click()
Label1.Caption = " " & File1.List(File1.ListIndex)
Nam = File1.List(File1.ListIndex)

End Sub

Private Sub File1_DblClick()
Label1.Caption = " " & File1.List(File1.ListIndex)
Nam = File1.List(File1.ListIndex)
Command1_Click
End Sub

Private Sub Form_Load()
Dim Naziv As String
'On Error Resume Next

If Right$(App.Path, 1) = "\" Then
File1.Path = App.Path & "Muzika" & "\"
PathANJA = App.Path & "\" & "Muzika"
Else
File1.Path = App.Path & "\" & "Muzika" & "\"
PathANJA = App.Path & "\" & "Muzika"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Odabir = Nothing

End Sub
