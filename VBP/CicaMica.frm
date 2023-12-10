VERSION 5.00
Begin VB.Form CicaMica 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CicaMica"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7470
      Left            =   360
      Picture         =   "CicaMica.frx":0000
      ScaleHeight     =   7470
      ScaleWidth      =   11265
      TabIndex        =   2
      Top             =   480
      Width           =   11265
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         MouseIcon       =   "CicaMica.frx":16EFF
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Ovuda..."
         Top             =   5040
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      MouseIcon       =   "CicaMica.frx":17051
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C  I  C  A   -   M  I  C  A  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   8040
      Width           =   11295
   End
End
Attribute VB_Name = "CicaMica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Poruke.mTada
End Sub

Private Sub Label3_Click()
Unload CicaMica
Set CicaMica = Nothing

End Sub
