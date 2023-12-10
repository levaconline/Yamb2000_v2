VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   " Yamb"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   11880
   DrawStyle       =   1  'Dash
   ForeColor       =   &H00004000&
   Icon            =   "Yamb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6360
      Top             =   7440
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   840
      TabIndex        =   66
      Top             =   3600
      Width           =   4455
      Begin VB.Label Label86 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   146
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label85 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   145
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   133
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   132
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   120
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   119
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   107
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   106
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   94
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   93
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   81
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   80
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   840
      TabIndex        =   60
      Top             =   480
      Width           =   4455
      Begin VB.Label Label92 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   153
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label84 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   144
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label83 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   143
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label82 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   142
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label81 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   141
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label80 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   140
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label79 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   139
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   131
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   130
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   129
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   128
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   127
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   126
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   118
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   117
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   116
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   115
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   114
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   113
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   105
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   104
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   103
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   102
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   101
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   100
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   92
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   91
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   90
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   89
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   88
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   87
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   79
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   78
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   77
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   76
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   75
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   74
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.TextBox Text128 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10560
      MousePointer    =   1  'Arrow
      TabIndex        =   59
      ToolTipText     =   "Broj praznih polja"
      Top             =   600
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   240
      Top             =   7440
   End
   Begin VB.TextBox Text127 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   58
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text126 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   57
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text125 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   56
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text124 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   55
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text123 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   54
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text122 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   53
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text121 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   52
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text120 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   51
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text119 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   7080
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   49
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text48 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   28
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text47 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   27
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text46 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   26
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text45 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   25
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text44 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   24
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text43 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text42 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   22
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text41 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   21
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   4560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1560
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bacanje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   5160
      Width           =   4455
      Begin VB.Label Label91 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   151
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label90 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   150
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label89 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   149
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label88 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   148
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label87 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   147
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label78 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   138
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label77 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   137
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   136
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   135
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         TabIndex        =   134
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   125
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   124
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   123
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label62 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   122
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         TabIndex        =   121
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   112
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   111
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   110
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   109
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   108
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   99
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   98
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   97
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   96
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   95
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   86
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   85
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   84
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   83
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   82
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1500
      Left            =   5400
      Picture         =   "Yamb.frx":0BC2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Sn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   152
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      Height          =   1575
      Left            =   6480
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   735
      Left            =   6480
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label rezx 
      BackStyle       =   0  'Transparent
      Caption         =   "Rezultat:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   7200
      TabIndex        =   50
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      DrawMode        =   10  'Mask Pen
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1935
      Left            =   6720
      Shape           =   2  'Oval
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label gfgh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KARO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label tyu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YAMB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label hhh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FUL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label aassdd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label ddssa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRILING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   44
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label nnn 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4560
      TabIndex        =   43
      ToolTipText     =   "Najava"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label ru 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   42
      Top             =   120
      Width           =   615
   End
   Begin VB.Label op 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "ih"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   41
      Top             =   120
      Width           =   615
   End
   Begin VB.Label po 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   40
      Top             =   120
      Width           =   615
   End
   Begin VB.Label oop 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   39
      Top             =   120
      Width           =   615
   End
   Begin VB.Label pop 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   38
      Top             =   120
      Width           =   255
   End
   Begin VB.Label sdd 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label dfd 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "MAX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label gh 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   35
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label pe 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   34
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label ce 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   33
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label reyu 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   32
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label tr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   31
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label dva 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   30
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Bro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   8640
      TabIndex        =   29
      ToolTipText     =   "Broj bacanja"
      Top             =   2400
      Width           =   735
   End
   Begin VB.Menu rr 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu ddf 
         Caption         =   ""
      End
   End
   Begin VB.Menu Opcije 
      Caption         =   "Opcije"
      Begin VB.Menu Nga 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Undo 
         Caption         =   "Undo"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Muz 
         Caption         =   "Muzika"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Crta 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Inf 
      Caption         =   "Inf"
      Begin VB.Menu rec 
         Caption         =   "Rekord"
      End
      Begin VB.Menu Uputstva 
         Caption         =   "Uputstva"
         Shortcut        =   {F5}
      End
      Begin VB.Menu et 
         Caption         =   "-"
      End
      Begin VB.Menu Yamb1 
         Caption         =   "Yamb"
      End
   End
   Begin VB.Menu Cha 
      Caption         =   "ch"
      Begin VB.Menu Cheat 
         Caption         =   "Cheat"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public a As Integer, A1 As Integer, A2 As Integer
Public A3 As Integer, A4 As Integer, A5 As Integer
Public p As Integer, n As Integer, q As Integer, t As Integer
Public k As Integer, b As Integer, f As Integer, ccc1 As Integer
Public ka As Integer, Ya As Integer, Deda As Integer
Public vre As Integer, Ruc As Integer, kas As Integer, has As Integer
Public B1 As String, B2 As String, B3 As String, B4 As String, B5 As String, B6 As String
Public Path As String

Private Sub Cheat_Click()

Dim ce As String, lo As Integer, sl As Integer
ce = InputBox("Upisite sifru:", "CHEAT")
If ce = "zdenka" Then
Load Nagrada
Nagrada.Show
End If
If ce = "zdenkas" Then
Load SuperNagrada
SuperNagrada.Show
End If

If ce = "ariadne" Then Image1.Visible = True

If ce = "cicamica" Then
Load CicaMica
CicaMica.Show
End If


If ce = "rekord" Then
sl = FreeFile
lo = InputBox("Upisite rekord", "Rekord")
    Open Path & "Res.rec" For Output As #sl
    Print #sl, Val(lo)
    Close #sl

End If

If ce = "rnd" Then
sl = FreeFile
lo = InputBox("RNDstart", "Rnd")
    Open Path & "Rnd.zec" For Output As #sl
    Print #sl, Val(lo)
    Close #sl
End If

If ce = "sloboda" Then
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
End If

End Sub

Private Sub Command1_Click()
a = a + 1
If a > 3 Then
Command1.Enabled = False
Exit Sub
End If
Bro.Caption = a
If t = 91 Then
    Load Adios
    Adios.Show
    Unload Form1
    Set Form1 = Nothing
    Exit Sub
End If

Rac.Popis

If t > 50 Then
If a > 1 Then

Rac.Glad
If Rac.Gla = 1 Then
If Rac.Ceta = 1 Then GoTo 545
a = 1
Bro.Caption = 1
Exit Sub
End If
545:
End If
End If

If a > 2 Then
Command1.Enabled = False
Else
Undo = False


End If
Ruc = 0

If a = 1 Then
Text13.Text = Int(Rnd * 6) + 1
Text14.Text = Int(Rnd * 6) + 1
Text15.Text = Int(Rnd * 6) + 1
Text16.Text = Int(Rnd * 6) + 1
Text17.Text = Int(Rnd * 6) + 1
Text18.Text = Int(Rnd * 6) + 1

Else

    If Val(Text13.Text) <> 0 Then
    Text13.Text = Int(Rnd * 6) + 1
    Ruc = Ruc + 1
    End If
    If Val(Text14.Text) <> 0 Then
    Text14.Text = Int(Rnd * 6) + 1
    Ruc = Ruc + 1
    End If
    If Val(Text15.Text) <> 0 Then
    Text15.Text = Int(Rnd * 6) + 1
    Ruc = Ruc + 1
    End If
    If Val(Text16.Text) <> 0 Then
    Text16.Text = Int(Rnd * 6) + 1
    Ruc = Ruc + 1
    End If
    If Val(Text17.Text) <> 0 Then
    Text17.Text = Int(Rnd * 6) + 1
    Ruc = Ruc + 1
    End If
    If Val(Text18.Text) <> 0 Then
    Text18.Text = Int(Rnd * 6) + 1
    Ruc = Ruc + 1
    End If
    
If t > 50 Then
If a > 1 Then

    Rac.RuB
    If Rac.Bla = 1 Then
a = a - 1
Bro.Caption = Val(Bro.Caption) - 1
Command1.Enabled = True
Ruc = 0
Exit Sub
End If
End If
End If
 End If
 
End Sub




Private Sub Exit_Click()
If MsgBox("Napustate YAMB?", vbYesNo, "Provera") = vbYes Then
Unload Form1
Set Form1 = Nothing
End
End If
End Sub


Private Sub Form_Load()
Dim zec As Single, uta As String, zam As Integer, Fuf As String
Dim Frr As Integer

Path = App.Path
If Right$(Path, 1) <> "\" Then
Path = Path & "\"
End If
Poruke.mSTART

uta = (Right$(Now, 2))
For zam = 0 To Val(uta)
zec = Rnd
Next
Undo = False

Fuf = Dir$(Path)
While Fuf <> ""
If Fuf = "Res.rec" Then Exit Sub
Fuf = Dir$
Wend
Frr = FreeFile
Open Path & "Res.rec" For Output As #Frr
Print #Frr, 2750
Close #Frr
End Sub



Private Sub Label10_Click()
If Label9.Caption <> "" Then
If Label10.Caption = "" Then
Kontrokor.Kont
    p = 10
    Undo = True

If k = 1 Then
    If Bro.Caption = 1 Then
    Label10.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label10.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label10.Caption = 40
    End If
    
Text120.Text = Val(Label9.Caption) + Val(Label10.Caption) + Val(Label11.Caption) + Val(Label12.Caption) + Val(Label13.Caption)

Else
Label10.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If

Else
Poruke.Red
End If

End Sub

Private Sub Label11_Click()
If Label10.Caption <> "" Then
If Label11.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis

    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label11.Caption = r + 60
Text120.Text = Val(Label9.Caption) + Val(Label10.Caption) + Val(Label11.Caption) + Val(Label12.Caption) + Val(Label13.Caption)

Else
Label11.Caption = 0
End If
    t = t + 1
p = 11
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label12_Click()
Dim r As Integer
If Label11.Caption <> "" Then
If Label12.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label12.Caption = ka + 70
    
    If ka = 0 Then
     Label12.Caption = 0
     End If
     If Ya > 1 Then
     Label12.Caption = r + 70
     End If

Text120.Text = Val(Label9.Caption) + Val(Label10.Caption) + Val(Label11.Caption) + Val(Label12.Caption) + Val(Label13.Caption)

     p = 12
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label13_Click()
If Label12.Caption <> "" Then
If Label13.Caption = "" Then
Undo = True

    Kontrokor.Kont
    Label13.Caption = Ya + 80

    If Ya = 0 Then
     Label13.Caption = 0
     End If
    Text120.Text = Val(Label9.Caption) + Val(Label10.Caption) + Val(Label11.Caption) + Val(Label12.Caption) + Val(Label13.Caption)
    t = t + 1
     p = 13
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label15_Click()
If Label15.Caption = "" Then
Undo = True

Rac.Dru
    'Upisivanje
        Label15.Caption = Rac.Re
        t = t + 1
        p = 15
    Text20.Text = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption) + Val(Label18.Caption) + Val(Label19.Caption)
End If

End Sub

Private Sub Label16_Click()
If Label16.Caption = "" Then
Undo = True

Rac.Tre
    'Upisivanje
        Label16.Caption = Rac.Re
        t = t + 1
        p = 16
    Text20.Text = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption) + Val(Label18.Caption) + Val(Label19.Caption)
End If

End Sub

Private Sub Label17_Click()
If Label17.Caption = "" Then
Undo = True

Rac.Cet
    'Upisivanje
        Label17.Caption = Rac.Re
        t = t + 1
        p = 17
    Text20.Text = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption) + Val(Label18.Caption) + Val(Label19.Caption)
End If

End Sub

Private Sub Label18_Click()
If Label18.Caption = "" Then
Undo = True

Rac.Pet
    'Upisivanje
        Label18.Caption = Rac.Re
        t = t + 1
        p = 18
    Text20.Text = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption) + Val(Label18.Caption) + Val(Label19.Caption)
End If

End Sub




Private Sub Label19_Click()
If Label19.Caption = "" Then
Undo = True

Rac.Ses
    'Upisivanje
        Label19.Caption = Rac.Re
        t = t + 1
        p = 19
    Text20.Text = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption) + Val(Label18.Caption) + Val(Label19.Caption)
End If

End Sub

Private Sub Label20_Click()
    If Label20.Caption = "" Then
    Label20.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text42.Text = Val(Label14.Caption) * (Val(Label20.Caption) - Val(Label21.Caption))
    If Label14.Caption = "" Or Label21.Caption = "" Then
    Text42.Text = 0
    End If
Undo = True

    t = t + 1
    
    'Upisivanje
        p = 20
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
End Sub

Private Sub Label21_Click()

    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label21.Caption = "" Then
    Label21.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text42.Text = Val(Label14.Caption) * (Val(Label20.Caption) - Val(Label21.Caption))
    If Label14.Caption = "" Or Label20.Caption = "" Then
    Text42.Text = 0
    End If
    Undo = True

    t = t + 1
    'Upisivanje
        p = 21
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If

End Sub

Private Sub Label22_Click()
Dim r As Integer, r1 As Integer, r2 As Integer, r3 As Integer, r4 As Integer, r5 As Integer
    If Label22.Caption = "" Then
    Undo = True

        Kontrokor.Kont
  ' Upis
        Label22.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label22.Caption = 0
        End If
        Text121.Text = Val(Label22.Caption) + Val(Label23.Caption) + Val(Label24.Caption) + Val(Label25.Caption) + Val(Label26.Caption)
        p = 22
    t = t + 1
      'Popis
    Rac.Popis
  
    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If

End Sub

Private Sub Label23_Click()
If Label23.Caption = "" Then
Undo = True

Kontrokor.Kont
    p = 23
If k = 1 Then
    If Bro.Caption = 1 Then
    Label23.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label23.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label23.Caption = 40
    End If
    
        Text121.Text = Val(Label22.Caption) + Val(Label23.Caption) + Val(Label24.Caption) + Val(Label25.Caption) + Val(Label26.Caption)

Else
Label23.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If

End Sub

Private Sub Label24_Click()
If Label24.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis

    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label24.Caption = r + 60
Text121.Text = Val(Label22.Caption) + Val(Label23.Caption) + Val(Label24.Caption) + Val(Label25.Caption) + Val(Label26.Caption)

Else
Label24.Caption = 0
End If
    t = t + 1
p = 24
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If

End Sub

Private Sub Label25_Click()
Dim r As Integer
If Label25.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label25.Caption = ka + 70
    
    If ka = 0 Then
     Label25.Caption = 0
     End If
     If Ya > 1 Then
     Label25.Caption = r + 70
     End If

Text121.Text = Val(Label22.Caption) + Val(Label23.Caption) + Val(Label24.Caption) + Val(Label25.Caption) + Val(Label26.Caption)

     p = 25
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If

End Sub

Private Sub Label26_Click()
If Label26.Caption = "" Then
Undo = True

    Kontrokor.Kont
    Label26.Caption = Ya + 80

    If Ya = 0 Then
     Label26.Caption = 0
     End If
Text121.Text = Val(Label22.Caption) + Val(Label23.Caption) + Val(Label24.Caption) + Val(Label25.Caption) + Val(Label26.Caption)
    t = t + 1
     p = 26
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If

End Sub

Private Sub Label28_Click()
If Label29.Caption <> "" Then
If Label28.Caption = "" Then
Undo = True

Rac.Dru
    'Upisivanje
        Label28.Caption = Rac.Re
        t = t + 1
        p = 28
    Text21.Text = Val(Label27.Caption) + Val(Label28.Caption) + Val(Label29.Caption) + Val(Label30.Caption) + Val(Label31.Caption) + Val(Label32.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label29_Click()
If Label30.Caption <> "" Then
If Label29.Caption = "" Then
Undo = True

Rac.Tre
    'Upisivanje
        Label29.Caption = Rac.Re
        t = t + 1
        p = 29
    Text21.Text = Val(Label27.Caption) + Val(Label28.Caption) + Val(Label29.Caption) + Val(Label30.Caption) + Val(Label31.Caption) + Val(Label32.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label3_Click()
If Label2.Caption <> "" Then
If Label3.Caption = "" Then
Undo = True

Rac.Tre
    'Upisivanje
        Label3.Caption = Rac.Re
        t = t + 1
        p = 3
    Text19.Text = Val(Label92.Caption) + Val(Label2.Caption) + Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption) + Val(Label6.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label30_Click()
If Label31.Caption <> "" Then
If Label30.Caption = "" Then
Undo = True

Rac.Cet
    'Upisivanje
        Label30.Caption = Rac.Re
        t = t + 1
        p = 30
    Text21.Text = Val(Label27.Caption) + Val(Label28.Caption) + Val(Label29.Caption) + Val(Label30.Caption) + Val(Label31.Caption) + Val(Label32.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label31_Click()
If Label32.Caption <> "" Then
If Label31.Caption = "" Then
Undo = True

Rac.Pet
    'Upisivanje
        Label31.Caption = Rac.Re
        t = t + 1
        p = 31
    Text21.Text = Val(Label27.Caption) + Val(Label28.Caption) + Val(Label29.Caption) + Val(Label30.Caption) + Val(Label31.Caption) + Val(Label32.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label32_Click()
If Label33.Caption <> "" Then
If Label31.Caption = "" Then
Undo = True

Rac.Ses
    'Upisivanje
        Label32.Caption = Rac.Re
        t = t + 1
        p = 32
    Text21.Text = Val(Label27.Caption) + Val(Label28.Caption) + Val(Label29.Caption) + Val(Label30.Caption) + Val(Label31.Caption) + Val(Label32.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label33_Click()
If Label34.Caption <> "" Then
    If Label33.Caption = "" Then
    Label33.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
Undo = True

    t = t + 1
    
    'Upisivanje
        p = 33
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Red
End If

End Sub

Private Sub Label34_Click()
If Label35.Caption <> "" Then
    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label34.Caption = "" Then
    Label34.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)

    Text43.Text = Val(Label27.Caption) * (Val(Label33.Caption) - Val(Label34.Caption))
    If Val(Label27.Caption) = 0 Then
    Text43.Text = 0
    End If
    Undo = True

    t = t + 1
    'Upisivanje
        p = 34
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If
Else
Poruke.Red
    
    
End If

End Sub

Private Sub Label35_Click()
Dim r As Integer, r1 As Integer, r2 As Integer, r3 As Integer, r4 As Integer, r5 As Integer
If Label36.Caption <> "" Then
    If Label35.Caption = "" Then
    Undo = True

        Kontrokor.Kont
  ' Upis
        Label35.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label35.Caption = 0
        End If
        Text122.Text = Val(Label35.Caption) + Val(Label36.Caption) + Val(Label37.Caption) + Val(Label38.Caption) + Val(Label39.Caption)
        p = 35
    t = t + 1
      'Popis
    Rac.Popis
  
    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Red
End If

End Sub

Private Sub Label36_Click()
If Label37.Caption <> "" Then
If Label36.Caption = "" Then
Undo = True

Kontrokor.Kont
    p = 36
If k = 1 Then
    If Bro.Caption = 1 Then
    Label36.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label36.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label36.Caption = 40
    End If
    
Text122.Text = Val(Label35.Caption) + Val(Label36.Caption) + Val(Label37.Caption) + Val(Label38.Caption) + Val(Label39.Caption)

Else
Label36.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If

Else
Poruke.Red
End If

End Sub

Private Sub Label37_Click()
If Label38.Caption <> "" Then
If Label37.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis

    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label37.Caption = r + 60
Text122.Text = Val(Label35.Caption) + Val(Label36.Caption) + Val(Label37.Caption) + Val(Label38.Caption) + Val(Label39.Caption)

Else
Label37.Caption = 0
End If
    t = t + 1
p = 37
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label38_Click()
Dim r As Integer
If Label39.Caption <> "" Then
If Label38.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label38.Caption = ka + 70
    
    If ka = 0 Then
     Label38.Caption = 0
     End If
     If Ya > 1 Then
     Label38.Caption = r + 70
     End If

Text122.Text = Val(Label35.Caption) + Val(Label36.Caption) + Val(Label37.Caption) + Val(Label38.Caption) + Val(Label39.Caption)

     p = 38
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label39_Click()
If Label39.Caption = "" Then
Undo = True

    Kontrokor.Kont
    Label39.Caption = Ya + 80

    If Ya = 0 Then
     Label39.Caption = 0
     End If
Text122.Text = Val(Label35.Caption) + Val(Label36.Caption) + Val(Label37.Caption) + Val(Label38.Caption) + Val(Label39.Caption)
    t = t + 1
     p = 39
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If

End Sub

Private Sub Label4_Click()
If Label3.Caption <> "" Then
If Label4.Caption = "" Then
Undo = True

Rac.Cet
    'Upisivanje
        Label4.Caption = Rac.Re
        t = t + 1
        p = 4
    Text19.Text = Val(Label92.Caption) + Val(Label2.Caption) + Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption) + Val(Label6.Caption)
End If
Else
Poruke.Red
End If
End Sub

Private Sub Label41_Click()
If Label42.Caption <> "" Then
If Label41.Caption = "" Then
Undo = True

Rac.Dru
    'Upisivanje
        Label41.Caption = Rac.Re
        t = t + 1
        p = 41
Text22.Text = Val(Label40.Caption) + Val(Label41.Caption) + Val(Label42.Caption) + Val(Label43.Caption) + Val(Label44.Caption) + Val(Label45.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label42_Click()
If Label43.Caption <> "" Then
If Label42.Caption = "" Then
Undo = True

Rac.Tre
    'Upisivanje
        Label42.Caption = Rac.Re
        t = t + 1
        p = 42
    Text22.Text = Val(Label40.Caption) + Val(Label41.Caption) + Val(Label42.Caption) + Val(Label43.Caption) + Val(Label44.Caption) + Val(Label45.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label43_Click()
If Label44.Caption <> "" Then
If Label43.Caption = "" Then
Undo = True

Rac.Cet
    'Upisivanje
        Label43.Caption = Rac.Re
        t = t + 1
        p = 43
    Text22.Text = Val(Label40.Caption) + Val(Label41.Caption) + Val(Label42.Caption) + Val(Label43.Caption) + Val(Label44.Caption) + Val(Label45.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label44_Click()
If Label45.Caption <> "" Then
If Label44.Caption = "" Then
Undo = True

Rac.Pet
    'Upisivanje
        Label44.Caption = Rac.Re
        t = t + 1
        p = 44
    Text22.Text = Val(Label40.Caption) + Val(Label41.Caption) + Val(Label42.Caption) + Val(Label43.Caption) + Val(Label44.Caption) + Val(Label45.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label45_Click()
If Label46.Caption <> "" Then
If Label45.Caption = "" Then
Undo = True

Rac.Ses
    'Upisivanje
        Label45.Caption = Rac.Re
        t = t + 1
        p = 45
    Text22.Text = Val(Label40.Caption) + Val(Label41.Caption) + Val(Label42.Caption) + Val(Label43.Caption) + Val(Label44.Caption) + Val(Label45.Caption)
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label46_Click()
    If Label46.Caption = "" Then
    Label46.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)

    t = t + 1
    Undo = True

    
    'Upisivanje
        p = 46
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
End Sub

Private Sub Label47_Click()
If Label46.Caption <> "" Then
    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label47.Caption = "" Then
    Label47.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text44.Text = Val(Label40.Caption) * (Val(Label46.Caption) - Val(Label47.Caption))
    If Label40.Caption = "" Then
    Text44.Text = 0
    End If
    t = t + 1
    Undo = True

    'Upisivanje
        p = 47
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If
Else
Poruke.Red
    
    
End If

End Sub

Private Sub Label48_Click()
Dim r As Integer, r1 As Integer, r2 As Integer, r3 As Integer, r4 As Integer, r5 As Integer
If Label47.Caption <> "" Then
    If Label48.Caption = "" Then
    Undo = True

        Kontrokor.Kont
  ' Upis
        Label48.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label48.Caption = 0
        End If
        Text123.Text = Val(Label48.Caption) + Val(Label49.Caption) + Val(Label50.Caption) + Val(Label51.Caption) + Val(Label52.Caption)
        p = 48
    t = t + 1
      'Popis
    Rac.Popis
  
    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Red
End If

End Sub

Private Sub Label49_Click()
If Label48.Caption <> "" Then
If Label49.Caption = "" Then
Undo = True

Kontrokor.Kont
    p = 49
If k = 1 Then
    If Bro.Caption = 1 Then
    Label49.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label49.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label49.Caption = 40
    End If
    
Text123.Text = Val(Label48.Caption) + Val(Label49.Caption) + Val(Label50.Caption) + Val(Label51.Caption) + Val(Label52.Caption)

Else
Label49.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If

Else
Poruke.Red
End If

End Sub

Private Sub Label5_Click()
If Label4.Caption <> "" Then
If Label5.Caption = "" Then
Undo = True

Rac.Pet
    'Upisivanje
        Label5.Caption = Rac.Re
        t = t + 1
        p = 5
    Text19.Text = Val(Label92.Caption) + Val(Label2.Caption) + Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption) + Val(Label6.Caption)
End If
Else
Poruke.Red
End If
End Sub

Private Sub Label50_Click()
If Label49.Caption <> "" Then
If Label50.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis

    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label50.Caption = r + 60
Text123.Text = Val(Label48.Caption) + Val(Label49.Caption) + Val(Label50.Caption) + Val(Label51.Caption) + Val(Label52.Caption)

Else
Label50.Caption = 0
End If
    t = t + 1
p = 50
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label51_Click()
Dim r As Integer
If Label50.Caption <> "" Then
If Label51.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label51.Caption = ka + 70
    
    If ka = 0 Then
     Label51.Caption = 0
     End If
     If Ya > 1 Then
     Label51.Caption = r + 70
     End If

Text123.Text = Val(Label48.Caption) + Val(Label49.Caption) + Val(Label50.Caption) + Val(Label51.Caption) + Val(Label52.Caption)

     p = 51
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label52_Click()
If Label51.Caption <> "" Then
If Label52.Caption = "" Then
Undo = True

    Kontrokor.Kont
    Label52.Caption = Ya + 80

    If Ya = 0 Then
     Label52.Caption = 0
     End If
Text123.Text = Val(Label48.Caption) + Val(Label49.Caption) + Val(Label50.Caption) + Val(Label51.Caption) + Val(Label52.Caption)
    t = t + 1
     p = 52
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Red
End If

End Sub

Private Sub Label53_Click()
If Ruc = 0 Or Ruc = 6 Then

If Label53.Caption = "" Then
Undo = True
ru.ForeColor = "&h00ffffff"
Rac.Prv
'Upisivanje
    Label53.Caption = Rac.Re
    p = 53
    t = t + 1
    Text23.Text = Val(Label53.Caption) + Val(Label54.Caption) + Val(Label55.Caption) + Val(Label56.Caption) + Val(Label57.Caption) + Val(Label58.Caption)
    Text45.Text = Val(Label53.Caption) * (Val(Label59.Caption) - Val(Label60.Caption))
    If Label59.Caption = "" Or Label60.Caption = "" Then
    Text45.Text = 0
    End If
End If

Else
Poruke.Ruc
End If




End Sub


Private Sub Label54_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label54.Caption = "" Then
Undo = True

Rac.Dru
'Upisivanje
    Label54.Caption = Rac.Re
    p = 54
    t = t + 1
    Text23.Text = Val(Label53.Caption) + Val(Label54.Caption) + Val(Label55.Caption) + Val(Label56.Caption) + Val(Label57.Caption) + Val(Label58.Caption)
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label55_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label55.Caption = "" Then
Undo = True

Rac.Tre
'Upisivanje
    Label55.Caption = Rac.Re
    p = 55
    t = t + 1
    Text23.Text = Val(Label53.Caption) + Val(Label54.Caption) + Val(Label55.Caption) + Val(Label56.Caption) + Val(Label57.Caption) + Val(Label58.Caption)
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label56_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label56.Caption = "" Then
Undo = True

Rac.Cet
'Upisivanje
    Label56.Caption = Rac.Re
    p = 56
    t = t + 1
    Text23.Text = Val(Label53.Caption) + Val(Label54.Caption) + Val(Label55.Caption) + Val(Label56.Caption) + Val(Label57.Caption) + Val(Label58.Caption)
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label57_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label57.Caption = "" Then
Undo = True

Rac.Pet
'Upisivanje
    Label57.Caption = Rac.Re
    p = 57
    t = t + 1
    Text23.Text = Val(Label53.Caption) + Val(Label54.Caption) + Val(Label55.Caption) + Val(Label56.Caption) + Val(Label57.Caption) + Val(Label58.Caption)
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label58_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label58.Caption = "" Then
Undo = True

Rac.Ses
'Upisivanje
    Label58.Caption = Rac.Re
    p = 58
    t = t + 1
    Text23.Text = Val(Label53.Caption) + Val(Label54.Caption) + Val(Label55.Caption) + Val(Label56.Caption) + Val(Label57.Caption) + Val(Label58.Caption)
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label59_Click()
If Ruc = 0 Or Ruc = 6 Then
    If Label59.Caption = "" Then

    Label59.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text45.Text = Val(Label53.Caption) * (Val(Label59.Caption) - Val(Label60.Caption))
    If Label8.Caption = "" Or Label60.Caption = "" Then
    Text45.Text = 0
    End If
    t = t + 1
    Undo = True

    'Upisivanje
        p = 59
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label6_Click()
If Label5.Caption <> "" Then
If Label6.Caption = "" Then
Undo = True

Rac.Ses
    'Upisivanje
        Label6.Caption = Rac.Re
        t = t + 1
        p = 6
    Text19.Text = Val(Label92.Caption) + Val(Label2.Caption) + Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption) + Val(Label6.Caption)
End If
Else
Poruke.Red
End If
End Sub

Private Sub Label60_Click()
If Ruc = 0 Or Ruc = 6 Then
    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label60.Caption = "" Then
    Label60.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text45.Text = Val(Label53.Caption) * (Val(Label59.Caption) - Val(Label60.Caption))
    If Label8.Caption = "" Or Label59.Caption = "" Then
    Text45.Text = 0
    End If
    t = t + 1
    Undo = True

    'Upisivanje
        p = 60
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label61_Click()
If Ruc = 0 Or Ruc = 6 Then
    If Label61.Caption = "" Then
    Undo = True

        Kontrokor.Kont
  ' Upis
        Label61.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label61.Caption = 0
        End If
Text124.Text = Val(Label61.Caption) + Val(Label62.Caption) + Val(Label63.Caption) + Val(Label64.Caption) + Val(Label65.Caption)
        p = 61
        t = t + 1
      'Popis
    Rac.Popis
  
    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label62_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label62.Caption = "" Then
Undo = True

Kontrokor.Kont
    p = 62
If k = 1 Then
    If Bro.Caption = 1 Then
    Label62.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label62.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label62.Caption = 40
    End If
    
Text124.Text = Val(Label61.Caption) + Val(Label62.Caption) + Val(Label63.Caption) + Val(Label64.Caption) + Val(Label65.Caption)

Else
Label62.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label63_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label63.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis
    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label63.Caption = r + 60
Text124.Text = Val(Label61.Caption) + Val(Label62.Caption) + Val(Label63.Caption) + Val(Label64.Caption) + Val(Label65.Caption)

Else
Label63.Caption = 0
End If
    t = t + 1
p = 63
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label64_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label64.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label64.Caption = ka + 70
    
    If ka = 0 Then
     Label64.Caption = 0
     End If
     If Ya > 1 Then
     Label64.Caption = r + 70
     End If

Text124.Text = Val(Label61.Caption) + Val(Label62.Caption) + Val(Label63.Caption) + Val(Label64.Caption) + Val(Label65.Caption)

     p = 64
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Ruc
End If

End Sub

Private Sub Label65_Click()
If Ruc = 0 Or Ruc = 6 Then
If Label65.Caption = "" Then
Undo = True

    Kontrokor.Kont
    Label65.Caption = Ya + 80
If Ya > 0 Then
Poruke.mAplauz
Image1.Visible = True
Timer2.Enabled = True
End If
    If Ya = 0 Then
     Label65.Caption = 0
    End If
Text124.Text = Val(Label61.Caption) + Val(Label62.Caption) + Val(Label63.Caption) + Val(Label64.Caption) + Val(Label65.Caption)
    t = t + 1
    p = 65
    
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Ruc
End If

End Sub


Private Sub Label67_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'---------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label67.Caption = "" Then
Undo = True

Rac.Dru
'Upisivanje
    Label67.Caption = Rac.Re
    p = 67
    t = t + 1
    q = 0
    Text24.Text = Val(Label66.Caption) + Val(Label67.Caption) + Val(Label68.Caption) + Val(Label69.Caption) + Val(Label70.Caption) + Val(Label71.Caption)
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub


Private Sub Label68_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'-------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label68.Caption = "" Then
Undo = True

Rac.Tre
'Upisivanje
    Label68.Caption = Rac.Re
    p = 68
    t = t + 1
    q = 0
    Text24.Text = Val(Label66.Caption) + Val(Label67.Caption) + Val(Label68.Caption) + Val(Label69.Caption) + Val(Label70.Caption) + Val(Label71.Caption)
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub


Private Sub Label69_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'---------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label69.Caption = "" Then
Undo = True

Rac.Cet
'Upisivanje
    Label69.Caption = Rac.Re
    p = 69
    t = t + 1
    q = 0
    Text24.Text = Val(Label66.Caption) + Val(Label67.Caption) + Val(Label68.Caption) + Val(Label69.Caption) + Val(Label70.Caption) + Val(Label71.Caption)
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label7_Click()
If Label6.Caption <> "" Then
    If Label7.Caption = "" Then
    Label7.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
Undo = True

    t = t + 1
    
    'Upisivanje
        p = 7
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Red
End If

End Sub


Private Sub Label70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'--------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label70.Caption = "" Then
Undo = True

Rac.Pet
'Upisivanje
    Label70.Caption = Rac.Re
    p = 70
    t = t + 1
    q = 0
    Text24.Text = Val(Label66.Caption) + Val(Label67.Caption) + Val(Label68.Caption) + Val(Label69.Caption) + Val(Label70.Caption) + Val(Label71.Caption)
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub


Private Sub Label71_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'--------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label71.Caption = "" Then
Undo = True

Rac.Ses
'Upisivanje
    Label71.Caption = Rac.Re
    p = 71
    t = t + 1
    q = 0
    Text24.Text = Val(Label66.Caption) + Val(Label67.Caption) + Val(Label68.Caption) + Val(Label69.Caption) + Val(Label70.Caption) + Val(Label71.Caption)
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label72_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'--------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then
    If Label72.Caption = "" Then
    Label72.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text46.Text = Val(Label66.Caption) * (Val(Label72.Caption) - Val(Label73.Caption))
If Label66.Caption = "" Or Label73.Caption = "" Then
    Text46.Text = 0
End If
    t = t + 1
    Undo = True

    'Upisivanje
        p = 72
'Popis
    Rac.A1 = Val(Text8.Text)
    Rac.A2 = Val(Text9.Text)
    Rac.A3 = Val(Text10.Text)
    Rac.A4 = Val(Text11.Text)
    Rac.A5 = Val(Text12.Text)

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
Sn.ForeColor = "&h00ffffff"
blok.deb
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label73_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'---------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then
    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label73.Caption = "" Then
    Label73.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text46.Text = Val(Label66.Caption) * (Val(Label72.Caption) - Val(Label73.Caption))
If Label66.Caption = "" Or Label72.Caption = "" Then
    Text46.Text = 0
End If
    t = t + 1
    Undo = True

    'Upisivanje
        p = 73
'Popis
    Rac.A1 = Val(Text8.Text)
    Rac.A2 = Val(Text9.Text)
    Rac.A3 = Val(Text10.Text)
    Rac.A4 = Val(Text11.Text)
    Rac.A5 = Val(Text12.Text)

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
Sn.ForeColor = "&h00ffffff"
blok.deb
        a = 0: Command1.Enabled = True
    
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If
Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label74_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'--------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label74.Caption = "" Then
Undo = True

        Kontrokor.Kont
  ' Upis
        Label74.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label74.Caption = 0
        End If
Text125.Text = Val(Label74.Caption) + Val(Label75.Caption) + Val(Label76.Caption) + Val(Label77.Caption) + Val(Label78.Caption)
    t = t + 1
        p = 74
        q = 0
      'Popis
    Rac.Popis
    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label75_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'-------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then
If Label75.Caption = "" Then
Undo = True

Kontrokor.Kont
    p = 75
If k = 1 Then
    If Bro.Caption = 1 Then
    Label75.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label75.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label75.Caption = 40
    End If
    
Text125.Text = Val(Label74.Caption) + Val(Label75.Caption) + Val(Label76.Caption) + Val(Label77.Caption) + Val(Label78.Caption)

Else
Label75.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
Sn.ForeColor = "&h00ffffff"
blok.deb
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label76_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then
If Label76.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis

    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label76.Caption = r + 60
Text125.Text = Val(Label74.Caption) + Val(Label75.Caption) + Val(Label76.Caption) + Val(Label77.Caption) + Val(Label78.Caption)

Else
Label76.Caption = 0
End If
    t = t + 1
p = 76
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
Sn.ForeColor = "&h00ffffff"
blok.deb
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label77_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'-----------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then
If Label77.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label77.Caption = ka + 70
    
    If ka = 0 Then
     Label77.Caption = 0
     End If
     If Ya > 1 Then
     Label77.Caption = r + 70
     End If

Text125.Text = Val(Label74.Caption) + Val(Label75.Caption) + Val(Label76.Caption) + Val(Label77.Caption) + Val(Label78.Caption)

     p = 77
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
Sn.ForeColor = "&h00ffffff"
blok.deb
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label78_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'---------------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then
If Label78.Caption = "" Then
    Kontrokor.Kont
    Label78.Caption = Ya + 80

    If Ya = 0 Then
     Label78.Caption = 0
     End If
Text125.Text = Val(Label74.Caption) + Val(Label75.Caption) + Val(Label76.Caption) + Val(Label77.Caption) + Val(Label78.Caption)
    t = t + 1
    Undo = True

     p = 78
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
Sn.ForeColor = "&h00ffffff"
blok.deb
    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Najava
End If
End If
End Sub


Private Sub Label79_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label79.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label79.BackColor = "&h0000ff00" Or Label79.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    nnn.ForeColor = "&h00ffffff"
    Label79.BackColor = "&h00c0ffc0"
    Exit Sub
    End If
blok.blok
blok.NBL
Label79.Enabled = True
Label79.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"
Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'*****************
If Button = vbLeftButton Then

If Label79.BackColor = "&h0000ff00" Or Label79.BackColor = "&h0000c000" Then
If Label79.Caption = "" Then
Undo = True

Rac.Prv
'Upisivanje
   Label79.Caption = Rac.Re
    p = 79
    t = t + 1
    Text25.Text = Val(Label79.Caption) + Val(Label80.Caption) + Val(Label81.Caption) + Val(Label82.Caption) + Val(Label83.Caption) + Val(Label84.Caption)
    Text47.Text = Val(Label79.Caption) * (Val(Label85.Caption) - Val(Label86.Caption))
    If Label85.Caption = "" Or Label86.Caption = "" Then
    Text47.Text = 0
    End If
    
    blok.deb
    blok.DENBL
    Label79.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub

Private Sub Label8_Click()
If Label7.Caption <> "" Then
    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label8.Caption = "" Then
    Label8.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
Undo = True

    Text41.Text = Label92.Caption * (Val(Label7.Caption) - Val(Label8.Caption))
    t = t + 1
    'Upisivanje
        p = 8
'Popis
    Rac.Popis

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If
Else
Poruke.Red
    
    
End If

End Sub


Private Sub Label80_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label80.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label80.BackColor = "&h0000ff00" Or Label80.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label80.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label80.Enabled = True
Label80.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'----------------------
If Button = vbLeftButton Then

If Label80.BackColor = "&h0000ff00" Or Label80.BackColor = "&h0000c000" Then
If Label80.Caption = "" Then
Undo = True

Rac.Dru
'Upisivanje
   Label80.Caption = Rac.Re
    p = 80
    t = t + 1
    Text25.Text = Val(Label79.Caption) + Val(Label80.Caption) + Val(Label81.Caption) + Val(Label82.Caption) + Val(Label83.Caption) + Val(Label84.Caption)
    
    blok.deb
    blok.DENBL
    Label80.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub


Private Sub Label81_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label81.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label81.BackColor = "&h0000ff00" Or Label80.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label81.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label81.Enabled = True
Label81.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'-------------------
If Button = vbLeftButton Then
If Label81.BackColor = "&h0000ff00" Or Label81.BackColor = "&h0000c000" Then
If Label81.Caption = "" Then
Undo = True

Rac.Tre
'Upisivanje
   Label81.Caption = Rac.Re
    p = 81
    t = t + 1
    Text25.Text = Val(Label79.Caption) + Val(Label80.Caption) + Val(Label81.Caption) + Val(Label82.Caption) + Val(Label83.Caption) + Val(Label84.Caption)
    
    blok.deb
    blok.DENBL
    Label81.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub


Private Sub Label82_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label82.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label82.BackColor = "&h0000ff00" Or Label82.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label82.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label82.Enabled = True
Label82.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'-------------------
If Button = vbLeftButton Then

If Label82.BackColor = "&h0000ff00" Or Label82.BackColor = "&h0000c000" Then
If Label82.Caption = "" Then
Undo = True

Rac.Cet
'Upisivanje
   Label82.Caption = Rac.Re
    p = 82
    t = t + 1
    Text25.Text = Val(Label79.Caption) + Val(Label80.Caption) + Val(Label81.Caption) + Val(Label82.Caption) + Val(Label83.Caption) + Val(Label84.Caption)
    
    blok.deb
    blok.DENBL
    Label82.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub


Private Sub Label83_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label83.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label83.BackColor = "&h0000ff00" Or Label83.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label83.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label83.Enabled = True
Label83.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'---------------
If Button = vbLeftButton Then
If Label83.BackColor = "&h0000ff00" Or Label83.BackColor = "&h0000c000" Then
If Label83.Caption = "" Then
Undo = True

Rac.Pet
'Upisivanje
   Label83.Caption = Rac.Re
    p = 83
    t = t + 1
    Text25.Text = Val(Label79.Caption) + Val(Label80.Caption) + Val(Label81.Caption) + Val(Label82.Caption) + Val(Label83.Caption) + Val(Label84.Caption)
    
    blok.deb
    blok.DENBL
    Label83.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub


Private Sub Label84_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label84.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label84.BackColor = "&h0000ff00" Or Label84.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label84.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label84.Enabled = True
Label84.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'-------------------
If Button = vbLeftButton Then
If Label84.BackColor = "&h0000ff00" Or Label84.BackColor = "&h0000c000" Then
If Label84.Caption = "" Then
Undo = True

Rac.Ses
'Upisivanje
   Label84.Caption = Rac.Re
    p = 84
    t = t + 1
    Text25.Text = Val(Label79.Caption) + Val(Label80.Caption) + Val(Label81.Caption) + Val(Label82.Caption) + Val(Label83.Caption) + Val(Label84.Caption)
    
    blok.deb
    blok.DENBL
    Label84.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub


Private Sub Label85_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label85.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label85.BackColor = "&h0000ff00" Or Label85.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label85.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label85.Enabled = True
Label85.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'------------------
If Button = vbLeftButton Then
If Label85.BackColor = "&h0000ff00" Or Label85.BackColor = "&h0000c000" Then
    If Label85.Caption = "" Then
    Label85.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text47.Text = Val(Label79.Caption) * (Val(Label85.Caption) - Val(Label86.Caption))
    If Label79.Caption = "" Or Label86.Caption = "" Then
    Text47.Text = 0
    End If

    t = t + 1
    Undo = True

    'Upisivanje
        p = 85

'Popis
    Rac.A1 = Val(Text8.Text)
    Rac.A2 = Val(Text9.Text)
    Rac.A3 = Val(Text10.Text)
    Rac.A4 = Val(Text11.Text)
    Rac.A5 = Val(Text12.Text)

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
    blok.deb
    blok.DENBL
    Label85.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
        a = 0: Command1.Enabled = True
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub

Private Sub Label86_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label86.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label86.BackColor = "&h0000ff00" Or Label86.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label86.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label86.Enabled = True
Label86.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'-----------------
If Button = vbLeftButton Then
If Label86.BackColor = "&h0000ff00" Or Label86.BackColor = "&h0000c000" Then
    If (Val(Text8.Text) * Val(Text9.Text) * Val(Text10.Text) * Val(Text11.Text) * Val(Text12.Text)) <> 0 Then
    If Label86.Caption = "" Then
    Label86.Caption = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text) + Val(Text12.Text)
    Text47.Text = Val(Label79.Caption) * (Val(Label85.Caption) - Val(Label86.Caption))
    If Label79.Caption = "" Or Label85.Caption = "" Then
    Text47.Text = 0
    End If
    t = t + 1
    Undo = True

    'Upisivanje
        p = 86
'Popis
    Rac.A1 = Val(Text8.Text)
    Rac.A2 = Val(Text9.Text)
    Rac.A3 = Val(Text10.Text)
    Rac.A4 = Val(Text11.Text)
    Rac.A5 = Val(Text12.Text)

    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
        blok.deb
    blok.DENBL
    Label86.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    
    Else
    MsgBox ("Odaberite 5 brojeva.")
    End If
    Else
Poruke.Najava
    
End If
End If
End Sub

Private Sub Label87_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label87.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label87.BackColor = "&h0000ff00" Or Label87.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label87.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label87.Enabled = True
Label87.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'----------------
If Button = vbLeftButton Then
Dim r As Integer
If Label87.BackColor = "&h0000ff00" Or Label87.BackColor = "&h0000c000" Then

If Label87.Caption = "" Then
Undo = True

        Kontrokor.Kont
  ' Upis
        Label87.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label87.Caption = 0
        End If
Text126.Text = Val(Label87.Caption) + Val(Label88.Caption) + Val(Label89.Caption) + Val(Label90.Caption) + Val(Label91.Caption)

'Popis
    Rac.Popis

'Upisivanje
    p = 87
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    
    a = 0: Command1.Enabled = True
    blok.deb
    blok.DENBL
    Label87.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    End If
    Else
Poruke.Najava

End If
End If
End Sub

Private Sub Label88_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label88.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label88.BackColor = "&h0000ff00" Or Label88.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label88.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label88.Enabled = True
Label88.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'--------------
If Button = vbLeftButton Then
If Label88.BackColor = "&h0000ff00" Or Label88.BackColor = "&h0000c000" Then

If Label88.Caption = "" Then
Undo = True

Kontrokor.Kont
    p = 88
If k = 1 Then
    If Bro.Caption = 1 Then
    Label88.Caption = 60
    End If
    
    If Bro.Caption = 2 Then
    Label88.Caption = 50
    End If
    
    If Bro.Caption = 3 Then
    Label88.Caption = 40
    End If
    
Text126.Text = Val(Label87.Caption) + Val(Label88.Caption) + Val(Label89.Caption) + Val(Label90.Caption) + Val(Label91.Caption)

Else
Label88.Caption = 0

End If
    t = t + 1
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    blok.deb
    blok.DENBL
    Label88.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    a = 0: Command1.Enabled = True
End If
    Else
Poruke.Najava

End If
End If
End Sub

Private Sub Label89_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label89.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label89.BackColor = "&h0000ff00" Or Label89.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label89.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label89.Enabled = True
Label89.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'-----------------
If Button = vbLeftButton Then
If Label89.BackColor = "&h0000ff00" Or Label89.BackColor = "&h0000c000" Then

If Label89.Caption = "" Then
Undo = True

Kontrokor.Kont

'Popis
    Rac.Popis

    r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4 + Rac.A5

If f = 1 Or Ya > 1 Then
Label89.Caption = r + 60
Text126.Text = Val(Label87.Caption) + Val(Label88.Caption) + Val(Label89.Caption) + Val(Label90.Caption) + Val(Label91.Caption)

Else
Label89.Caption = 0
End If
    t = t + 1
p = 89
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    blok.deb
    blok.DENBL
    Label89.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    a = 0: Command1.Enabled = True
End If
Else
Poruke.Najava
End If
End If
End Sub

Private Sub Label9_Click()
Dim r As Integer, r1 As Integer, r2 As Integer, r3 As Integer, r4 As Integer, r5 As Integer
If Label8.Caption <> "" Then
    If Label9.Caption = "" Then
        Kontrokor.Kont
  ' Upis
        Label9.Caption = Kontrokor.r + 40
        If Kontrokor.r = 0 Then
        Label9.Caption = 0
        End If
        Text120.Text = Val(Label9.Caption) + Val(Label10.Caption) + Val(Label11.Caption) + Val(Label12.Caption) + Val(Label13.Caption)
        p = 9
        Undo = True

    t = t + 1
      'Popis
    Rac.Popis
  
    ' Brisanje polja
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""
        Text15.Text = ""
        Text16.Text = ""
        Text17.Text = ""
        Text18.Text = ""
        
        a = 0: Command1.Enabled = True
    End If
Else
Poruke.Red
End If

End Sub

Private Sub Label90_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label90.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label90.BackColor = "&h0000ff00" Or Label90.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label90.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label90.Enabled = True
Label90.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'-------------------
If Button = vbLeftButton Then
Dim r As Integer
If Label90.BackColor = "&h0000ff00" Or Label90.BackColor = "&h0000c000" Then
If Label90.Caption = "" Then
Undo = True

    Kontrokor.Kont
'Popis
    Rac.Popis
r = Rac.A1 + Rac.A2 + Rac.A3 + Rac.A4
    Label90.Caption = ka + 70
    
    If ka = 0 Then
     Label90.Caption = 0
     End If
     If Ya > 1 Then
     Label90.Caption = r + 70
     End If

Text126.Text = Val(Label87.Caption) + Val(Label88.Caption) + Val(Label89.Caption) + Val(Label90.Caption) + Val(Label91.Caption)

     p = 90
    t = t + 1
' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    blok.deb
    blok.DENBL
    Label90.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"

    a = 0: Command1.Enabled = True
    
End If
    Else
Poruke.Najava

End If
End If
End Sub

Private Sub Label91_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
If Label91.Caption = "" Then
    If Val(Bro.Caption) <= 1 Then
    If Label91.BackColor = "&h0000ff00" Or Label91.BackColor = "&h0000c000" Then
    blok.deb
    blok.DENBL
    Label91.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"
    Exit Sub
    End If
blok.blok
blok.NBL
Label91.Enabled = True
Label91.BackColor = "&h0000ff00"
nnn.ForeColor = "&h000000ff"

Else
MsgBox ("Najava ili odjava se vrsi samo posle prvog bacanja.")
End If
End If
End If
'----------------
If Button = vbLeftButton Then
If Label91.BackColor = "&h0000ff00" Or Label91.BackColor = "&h0000c000" Then
If Label91.Caption = "" Then
Undo = True

    Kontrokor.Kont
    Label91.Caption = Ya + 80

    If Ya = 0 Then
     Label91.Caption = 0
     End If
Text126.Text = Val(Label87.Caption) + Val(Label88.Caption) + Val(Label89.Caption) + Val(Label90.Caption) + Val(Label91.Caption)
    t = t + 1
     p = 91
'Popis
    Rac.Popis

' Brisanje polja
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    blok.deb
    blok.DENBL
    Label91.BackColor = "&h00c0ffc0"
    nnn.ForeColor = "&h00ffffff"

    a = 0: Command1.Enabled = True
    
End If
Else
Poruke.Najava

End If
End If
End Sub

Private Sub Muz_Click()
Poruke.mSTART
End Sub

Private Sub Nga_Click()
If MsgBox("Prekidate aktuelnu igru?", vbOKCancel, "Potvrda") = vbOK Then
Ng.Ng
End If

End Sub

Private Sub rec_Click()
Dim sl As Integer, Rez As String
sl = FreeFile

Open Path & "Res.rec" For Input As #sl
Line Input #sl, Rez
Close #sl
MsgBox ("Trenutni rekord je:" & Rez)
End Sub

Public Sub sn_Click()
If Val(Bro.Caption) <= 1 Then

q = q + 1
If (q Mod 2) = 0 Then
Sn.ForeColor = "&h00ffffff"
blok.deb
Else
Sn.ForeColor = "&h000000ff"
blok.blok
End If

If q > 20 Then q = 1

Else
MsgBox ("Kolonu mozete najaviti, ili se predomisliti, samo posle prvog bacanja.")
End If

End Sub



Private Sub Label92_Click()
If Label92.Caption = "" Then
Undo = True

Rac.Prv
'Upisivanje
    Label92.Caption = Rac.Re
    t = t + 1
    p = 92
    Text19.Text = Val(Label92.Caption) + Val(Label2.Caption) + Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption) + Val(Label6.Caption)
End If

End Sub


Private Sub Label66_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
Form1.sn_Click
End If
'-----------------
If Button = vbLeftButton Then
If Sn.ForeColor = "&h000000ff" Or a = 1 Then

If Label66.Caption = "" Then
Undo = True

Rac.Prv
'Upisivanje
    Label66.Caption = Rac.Re
    p = 66
    t = t + 1
    q = 0
    Text24.Text = Val(Label66.Caption) + Val(Label67.Caption) + Val(Label68.Caption) + Val(Label69.Caption) + Val(Label70.Caption) + Val(Label71.Caption)
    Text46.Text = Val(Label66.Caption) * (Val(Label72.Caption) - Val(Label73.Caption))
If Label72.Caption = "" Or Label73.Caption = "" Then
    Text46.Text = 0
End If
    
Sn.ForeColor = "&h00ffffff"
blok.deb
End If

Else
Poruke.Najava
End If
End If
End Sub



Private Sub Text10_Click()

    If Text13.Text = "" Then
        Text13.Text = Text10.Text
        Text10.Text = ""
        GoTo 11
    End If
    
    If Text14.Text = "" Then
        Text14.Text = Text10.Text
        Text10.Text = ""
        GoTo 11
    End If
    
    If Text15.Text = "" Then
        Text15.Text = Text10.Text
        Text10.Text = ""
        GoTo 11
    End If
    
    If Text16.Text = "" Then
        Text16.Text = Text10.Text
        Text10.Text = ""
        GoTo 11
    End If
    
    If Text17.Text = "" Then
        Text17.Text = Text10.Text
        Text10.Text = ""
        GoTo 11
    End If
    
    If Text18.Text = "" Then
        Text18.Text = Text10.Text
        Text10.Text = ""
        GoTo 11
    End If
    
11:

End Sub



Private Sub Text11_Click()

    If Text13.Text = "" Then
        Text13.Text = Text11.Text
        Text11.Text = ""
        GoTo 11
    End If
    
    If Text14.Text = "" Then
        Text14.Text = Text11.Text
        Text11.Text = ""
        GoTo 11
    End If
    
    If Text15.Text = "" Then
        Text15.Text = Text11.Text
        Text11.Text = ""
        GoTo 11
    End If
    
    If Text16.Text = "" Then
        Text16.Text = Text11.Text
        Text11.Text = ""
        GoTo 11
    End If
    
    If Text17.Text = "" Then
        Text17.Text = Text11.Text
        Text11.Text = ""
        GoTo 11
    End If
    
    If Text18.Text = "" Then
        Text18.Text = Text11.Text
        Text11.Text = ""
        GoTo 11
    End If
    
11:

End Sub




Private Sub Text12_Click()

    If Text13.Text = "" Then
        Text13.Text = Text12.Text
        Text12.Text = ""
        GoTo 11
    End If
    
    If Text14.Text = "" Then
        Text14.Text = Text12.Text
        Text12.Text = ""
        GoTo 11
    End If
    
    If Text15.Text = "" Then
        Text15.Text = Text12.Text
        Text12.Text = ""
        GoTo 11
    End If
    
    If Text16.Text = "" Then
        Text16.Text = Text12.Text
        Text12.Text = ""
        GoTo 11
    End If
    
    If Text17.Text = "" Then
        Text17.Text = Text12.Text
        Text12.Text = ""
        GoTo 11
    End If
    
    If Text18.Text = "" Then
        Text18.Text = Text12.Text
        Text12.Text = ""
        GoTo 11
    End If
    
11:

End Sub

Private Sub Text120_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)
End Sub

Private Sub Text121_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)

End Sub

Private Sub Text122_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)

End Sub

Private Sub Text123_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)

End Sub

Private Sub Text124_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)

End Sub

Private Sub Text125_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)

End Sub

Private Sub Text126_Change()
Text127.Text = Val(Text120.Text) + Val(Text121.Text) + Val(Text122.Text) + Val(Text123.Text) + Val(Text124.Text) + Val(Text125.Text) + Val(Text126.Text)

End Sub

Private Sub Text127_Change()
Text119.Text = Val(Text48.Text) + Val(Text26.Text) + Val(Text127.Text)

End Sub

Private Sub Text128_Change()
If Label79.BackColor <> "&h00c0ffc0" Then
Label79.BackColor = (Text128.BackColor)
Label79.BackColor = (Text128.BackColor)

GoTo 2
End If
If Label80.BackColor <> "&h00c0ffc0" Then
Label80.BackColor = Text128.BackColor
GoTo 2
End If
If Label81.BackColor <> "&h00c0ffc0" Then
Label81.BackColor = Text128.BackColor
GoTo 2
End If
If Label82.BackColor <> "&h00c0ffc0" Then
Label82.BackColor = Text128.BackColor
GoTo 2
End If
If Label83.BackColor <> "&h00c0ffc0" Then
Label83.BackColor = Text128.BackColor
GoTo 2
End If
If Label84.BackColor <> "&h00c0ffc0" Then
Label84.BackColor = Text128.BackColor
GoTo 2
End If
If Label85.BackColor <> "&h00c0ffc0" Then
Label85.BackColor = Text128.BackColor
GoTo 2
End If

If Label86.BackColor <> "&h00c0ffc0" Then
Label86.BackColor = Text128.BackColor
GoTo 2
End If

If Label87.BackColor <> "&h00c0ffc0" Then
Label87.BackColor = Text128.BackColor
GoTo 2
End If
If Label88.BackColor <> "&h00c0ffc0" Then
Label88.BackColor = Text128.BackColor
GoTo 2
End If
If Label89.BackColor <> "&h00c0ffc0" Then
Label89.BackColor = Text128.BackColor
GoTo 2
End If
If Label90.BackColor <> "&h00c0ffc0" Then
Label90.BackColor = Text128.BackColor
GoTo 2
End If
If Label91.BackColor <> "&h00c0ffc0" Then
Label91.BackColor = Text128.BackColor
GoTo 2
End If
2:
End Sub

Private Sub Text13_Click()


    If Text8.Text = "" Then
        Text8.Text = Text13.Text
        Text13.Text = ""
        GoTo 11
    End If
    
    If Text9.Text = "" Then
        Text9.Text = Text13.Text
        Text13.Text = ""
        GoTo 11
    End If
    
    If Text10.Text = "" Then
        Text10.Text = Text13.Text
        Text13.Text = ""
        GoTo 11
    End If
    
    If Text11.Text = "" Then
        Text11.Text = Text13.Text
        Text13.Text = ""
        GoTo 11
    End If
    
    If Text12.Text = "" Then
        Text12.Text = Text13.Text
        Text13.Text = ""
        GoTo 11
    End If
    
       
11:
End Sub


Private Sub Text14_Click()


    If Text8.Text = "" Then
        Text8.Text = Text14.Text
        Text14.Text = ""
        GoTo 11
    End If
    
    If Text9.Text = "" Then
        Text9.Text = Text14.Text
        Text14.Text = ""
        GoTo 11
    End If
    
    If Text10.Text = "" Then
        Text10.Text = Text14.Text
        Text14.Text = ""
        GoTo 11
    End If
    
    If Text11.Text = "" Then
        Text11.Text = Text14.Text
        Text14.Text = ""
        GoTo 11
    End If
    
    If Text12.Text = "" Then
        Text12.Text = Text14.Text
        Text14.Text = ""
        GoTo 11
    End If
    
       
11:

End Sub


Private Sub Text15_Click()

    If Text8.Text = "" Then
        Text8.Text = Text15.Text
        Text15.Text = ""
        GoTo 11
    End If
    
    If Text9.Text = "" Then
        Text9.Text = Text15.Text
        Text15.Text = ""
        GoTo 11
    End If
    
    If Text10.Text = "" Then
        Text10.Text = Text15.Text
        Text15.Text = ""
        GoTo 11
    End If
    
    If Text11.Text = "" Then
        Text11.Text = Text15.Text
        Text15.Text = ""
        GoTo 11
    End If
    
    If Text12.Text = "" Then
        Text12.Text = Text15.Text
        Text15.Text = ""
        GoTo 11
    End If
    
       
11:

End Sub

Private Sub Text16_Click()

    If Text8.Text = "" Then
        Text8.Text = Text16.Text
        Text16.Text = ""
        GoTo 11
    End If
    
    If Text9.Text = "" Then
        Text9.Text = Text16.Text
        Text16.Text = ""
        GoTo 11
    End If
    
    If Text10.Text = "" Then
        Text10.Text = Text16.Text
        Text16.Text = ""
        GoTo 11
    End If
    
    If Text11.Text = "" Then
        Text11.Text = Text16.Text
        Text16.Text = ""
        GoTo 11
    End If
    
    If Text12.Text = "" Then
        Text12.Text = Text16.Text
        Text16.Text = ""
        GoTo 11
    End If
    
       
11:

End Sub

Private Sub Text17_Click()

    If Text8.Text = "" Then
        Text8.Text = Text17.Text
        Text17.Text = ""
        GoTo 11
    End If
    
    If Text9.Text = "" Then
        Text9.Text = Text17.Text
        Text17.Text = ""
        GoTo 11
    End If
    
    If Text10.Text = "" Then
        Text10.Text = Text17.Text
        Text17.Text = ""
        GoTo 11
    End If
    
    If Text11.Text = "" Then
        Text11.Text = Text17.Text
        Text17.Text = ""
        GoTo 11
    End If
    
    If Text12.Text = "" Then
        Text12.Text = Text17.Text
        Text17.Text = ""
        GoTo 11
    End If
    
       
11:

End Sub


Private Sub Text18_Click()

    If Text8.Text = "" Then
        Text8.Text = Text18.Text
        Text18.Text = ""
        GoTo 11
    End If
    
    If Text9.Text = "" Then
        Text9.Text = Text18.Text
        Text18.Text = ""
        GoTo 11
    End If
    
    If Text10.Text = "" Then
        Text10.Text = Text18.Text
        Text18.Text = ""
        GoTo 11
    End If
    
    If Text11.Text = "" Then
        Text11.Text = Text18.Text
        Text18.Text = ""
        GoTo 11
    End If
    
    If Text12.Text = "" Then
        Text12.Text = Text18.Text
        Text18.Text = ""
        GoTo 11
    End If
    
       
11:

End Sub

Private Sub Text19_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)
End Sub


Private Sub Label14_Click()
If Label14.Caption = "" Then
Undo = True

Rac.Prv
'Upisivanje
    Label14.Caption = Rac.Re
    p = 14
    t = t + 1
    Text20.Text = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label17.Caption) + Val(Label18.Caption) + Val(Label19.Caption)
    Text42.Text = Val(Label14.Caption) * (Val(Label20.Caption) - Val(Label21.Caption))
    If Val(Label20.Caption) = 0 Or Val(Label21.Caption) = 0 Then
    Text42.Text = 0
    End If
End If

End Sub

Private Sub Text20_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)

End Sub

Private Sub Text21_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)

End Sub

Private Sub Text22_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)

End Sub

Private Sub Text23_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)

End Sub

Private Sub Text24_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)

End Sub

Private Sub Text25_Change()
Text26.Text = Val(Text19.Text) + Val(Text20.Text) + Val(Text21.Text) + Val(Text22.Text) + Val(Text23.Text) + Val(Text24.Text) + Val(Text25.Text)

End Sub

Private Sub Text26_Change()
Text119.Text = Val(Text48.Text) + Val(Text26.Text) + Val(Text127.Text)
End Sub



Private Sub Label27_Click()
If Label28.Caption <> "" Then
If Label27.Caption = "" Then
Undo = True

 Rac.Prv
    'Upisivanje
        Label27.Caption = Rac.Re
        p = 27
        t = t + 1

    Text21.Text = Val(Label27.Caption) + Val(Label28.Caption) + Val(Label29.Caption) + Val(Label30.Caption) + Val(Label31.Caption) + Val(Label32.Caption)
    Text43.Text = Val(Label27.Caption) * (Val(Label33.Caption) - Val(Label34.Caption))

    End If
Else
Poruke.Red
End If

End Sub



Private Sub Label40_Click()
If Label41.Caption <> "" Then
If Label40.Caption = "" Then
Undo = True

Rac.Prv
    'Upisivanje
       Label40.Caption = Rac.Re
        p = 40
        t = t + 1
        Text22.Text = Val(Label40.Caption) + Val(Label41.Caption) + Val(Label42.Caption) + Val(Label43.Caption) + Val(Label44.Caption) + Val(Label45.Caption)
        Text44.Text = Val(Label40.Caption) * (Val(Label46.Caption) - Val(Label47.Caption))

        If Val(Label47.Caption) = 0 Then
            Text44.Text = 0
        End If
        End If
Else
Poruke.Red
    
End If

End Sub


Private Sub Text41_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)
End Sub

Private Sub Text42_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)

End Sub

Private Sub Text43_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)

End Sub

Private Sub Text44_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)

End Sub

Private Sub Text45_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)

End Sub

Private Sub Text46_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)

End Sub

Private Sub Text47_Change()
Text48.Text = Val(Text41.Text) + Val(Text42.Text) + Val(Text43.Text) + Val(Text44.Text) + Val(Text45.Text) + Val(Text46.Text) + Val(Text47.Text)

End Sub

Private Sub Text48_Change()
Text119.Text = Val(Text48.Text) + Val(Text26.Text) + Val(Text127.Text)

End Sub

Private Sub Label2_Click()
If Label92.Caption <> "" Then
If Label2.Caption = "" Then
Undo = True

Rac.Dru
    'Upisivanje
        Label2.Caption = Rac.Re
        t = t + 1
        p = 2
    Text19.Text = Val(Label92.Caption) + Val(Label2.Caption) + Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption) + Val(Label6.Caption)
End If
Else
Poruke.Red
End If
End Sub








Private Sub Text8_Click()


    If Text13.Text = "" Then
        Text13.Text = Text8.Text
        Text8.Text = ""
        GoTo 11
    End If
    
    If Text14.Text = "" Then
        Text14.Text = Text8.Text
        Text8.Text = ""
        GoTo 11
    End If
    
    If Text15.Text = "" Then
        Text15.Text = Text8.Text
        Text8.Text = ""
        GoTo 11
    End If
    
    If Text16.Text = "" Then
        Text16.Text = Text8.Text
        Text8.Text = ""
        GoTo 11
    End If
    
    If Text17.Text = "" Then
        Text17.Text = Text8.Text
        Text8.Text = ""
        GoTo 11
    End If
    
    If Text18.Text = "" Then
        Text18.Text = Text8.Text
        Text8.Text = ""
        GoTo 11
    End If
    
11:

End Sub




Private Sub Text9_Click()

    If Text13.Text = "" Then
        Text13.Text = Text9.Text
        Text9.Text = ""
        GoTo 11
    End If
    
    If Text14.Text = "" Then
        Text14.Text = Text9.Text
        Text9.Text = ""
        GoTo 11
    End If
    
    If Text15.Text = "" Then
        Text15.Text = Text9.Text
        Text9.Text = ""
        GoTo 11
    End If
    
    If Text16.Text = "" Then
        Text16.Text = Text9.Text
        Text9.Text = ""
        GoTo 11
    End If
    
    If Text17.Text = "" Then
        Text17.Text = Text9.Text
        Text9.Text = ""
        GoTo 11
    End If
    
    If Text18.Text = "" Then
        Text18.Text = Text9.Text
        Text9.Text = ""
        GoTo 11
    End If
    
11:

End Sub

Private Sub Timer1_Timer()
vre = vre + 1
If vre > 1 Then
Text128.BackColor = "&h0000ff00"
Text128.Text = ""
Else
Text128.BackColor = "&h0000c000"
Text128.Text = 91 - t
End If
If vre = 2 Then vre = 0
End Sub


Private Sub Timer2_Timer()
ccc1 = ccc1 + 1
If ccc1 > 10 Then
Image1.Visible = False
Timer2.Enabled = False
End If
End Sub

Private Sub Undo_Click()
Und.Und
End Sub


Private Sub Uputstva_Click()
Load Informacije
Informacije.Show

End Sub

Private Sub Yamb1_Click()
Load Yamb
Yamb.Show
End Sub

Public Sub Boy()
    Label79.BackColor = "&h00c0ffc0"
    Label80.BackColor = "&h00c0ffc0"
    Label81.BackColor = "&h00c0ffc0"
    Label82.BackColor = "&h00c0ffc0"
    Label83.BackColor = "&h00c0ffc0"
    Label84.BackColor = "&h00c0ffc0"
    Label85.BackColor = "&h00c0ffc0"
    Label86.BackColor = "&h00c0ffc0"
    Label87.BackColor = "&h00c0ffc0"
    Label88.BackColor = "&h00c0ffc0"
    Label89.BackColor = "&h00c0ffc0"
    Label90.BackColor = "&h00c0ffc0"
    Label91.BackColor = "&h00c0ffc0"

End Sub

