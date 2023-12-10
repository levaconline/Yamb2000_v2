VERSION 5.00
Begin VB.Form Informacije 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPUTSTVA"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "Informacije.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   546.954
   ScaleMode       =   0  'User
   ScaleWidth      =   823.598
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10320
      Picture         =   "Informacije.frx":030A
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   8640
      Picture         =   "Informacije.frx":0DAE
      ScaleHeight     =   840
      ScaleWidth      =   2730
      TabIndex        =   1
      Top             =   480
      Width           =   2730
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8205
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Informacije.frx":35C0
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "Informacije"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Poruke.mGong

End Sub

