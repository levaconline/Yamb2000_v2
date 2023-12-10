VERSION 5.00
Begin VB.Form Yamb 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yamb"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "HRONOS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "HRONOS.frx":030A
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000080&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   3360
      Picture         =   "HRONOS.frx":03A4
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   840
      Picture         =   "HRONOS.frx":06AE
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   1680
      X2              =   3120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "YAMB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Yamb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Yamb
Set Yamb = Nothing

End Sub

Private Sub Form_Load()
Poruke.mDobos
End Sub

