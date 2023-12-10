VERSION 5.00
Begin VB.Form SuperNagrada 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SuperNagrada"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillStyle       =   0  'Solid
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
   Picture         =   "SuperNagrada.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z D E N K A"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   8160
      Width           =   6375
   End
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
      Height          =   615
      Left            =   4680
      MouseIcon       =   "SuperNagrada.frx":2A40C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "EXIT"
      Top             =   2280
      Width           =   615
   End
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
      Height          =   615
      Left            =   6120
      MouseIcon       =   "SuperNagrada.frx":2A716
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Tajni prolaz..."
      Top             =   5040
      Width           =   495
   End
End
Attribute VB_Name = "SuperNagrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Poruke.mTada
End Sub

Private Sub Label1_Click()
Unload SuperNagrada
Set SuperNagrada = Nothing
End Sub

Private Sub Label2_Click()
Unload SuperNagrada
Set SuperNagrada = Nothing

End Sub
