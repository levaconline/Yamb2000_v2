VERSION 5.00
Begin VB.Form Nagrada 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Nagrada"
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
   Icon            =   "Nagrada.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   9015
      Left            =   6000
      Picture         =   "Nagrada.frx":030A
      ScaleHeight     =   8955
      ScaleWidth      =   5955
      TabIndex        =   1
      ToolTipText     =   "Zdenka"
      Top             =   0
      Width           =   6015
      Begin VB.Label Label1 
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
         Left            =   3840
         MouseIcon       =   "Nagrada.frx":10F5E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "SAMO NAPRED!"
         Top             =   5400
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   0
      Picture         =   "Nagrada.frx":110B0
      ScaleHeight     =   8955
      ScaleWidth      =   6075
      TabIndex        =   0
      ToolTipText     =   "Zdenka"
      Top             =   0
      Width           =   6135
      Begin VB.Label Label2 
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
         Height          =   1455
         Left            =   4080
         MouseIcon       =   "Nagrada.frx":2040F
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "PRAVI PUT..."
         Top             =   5760
         Width           =   975
      End
   End
End
Attribute VB_Name = "Nagrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 Poruke.mTada
End Sub

Private Sub Label1_Click()
Unload Nagrada
Set Nagrada = Nothing

End Sub

Private Sub Label2_Click()
Unload Nagrada
Set Nagrada = Nothing

End Sub
