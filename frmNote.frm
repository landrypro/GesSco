VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmNote 
   BackColor       =   &H80000013&
   Caption         =   "Saisie Des Notes"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   13335
   Icon            =   "frmNote.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10605
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Caption         =   "Informations sur la Classe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   600
      TabIndex        =   104
      Top             =   8880
      Width           =   3375
      Begin VB.TextBox txtEffectif 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "  "
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Effectifs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Informations sur l'Elève"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   480
      TabIndex        =   95
      Top             =   5640
      Width           =   3375
      Begin VB.TextBox txtStatut 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtSexe 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtPrenom 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtNom 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Statut"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sexe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   101
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Prénom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   600
         Width           =   855
      End
   End
   Begin MSDataListLib.DataList dlistEleve 
      Height          =   4155
      Left            =   480
      TabIndex        =   94
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7329
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483629
      ForeColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcmbClasse 
      Height          =   375
      Left            =   1440
      TabIndex        =   93
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   12582912
      Text            =   "Choisissez Votre Classe"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   11055
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   130
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   129
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   128
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   127
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   126
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   125
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox txtRedac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   124
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox txtRedac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   123
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox txtRedac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   122
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox txtRedac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   121
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox txtRedac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   120
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox txtRedac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   119
         Top             =   9000
         Width           =   615
      End
      Begin VB.TextBox txtPCT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   118
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtPCT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   117
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtPCT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   116
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtPCT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   115
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtPCT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   114
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtPCT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   113
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtLV2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   112
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox txtLV2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   111
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox txtLV2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   110
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox txtLV2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   109
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox txtLV2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   108
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox txtLV2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   107
         Top             =   7920
         Width           =   615
      End
      Begin VB.TextBox txtDictee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   92
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtDictee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   91
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtDictee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   90
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtDictee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   89
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtDictee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   88
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtEPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   87
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtEPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   86
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtEPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   85
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtEPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   84
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtEPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   83
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtTM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   82
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox txtTM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   81
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox txtTM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   80
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox txtTM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   79
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox txtTM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   78
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox txtEtudText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   77
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtEtudText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   76
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtEtudText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   75
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtEtudText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   74
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtEtudText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   73
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtLitterature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   9120
         TabIndex        =   72
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtLitterature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   7800
         TabIndex        =   71
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtLitterature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   6480
         TabIndex        =   70
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtLitterature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5160
         TabIndex        =   69
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtLitterature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3960
         TabIndex        =   68
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtLangue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   67
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtLangue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   66
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtLangue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   65
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtLangue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   64
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtLangue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   63
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtPhilo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   9120
         TabIndex        =   62
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtPhilo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   7800
         TabIndex        =   61
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtPhilo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   6480
         TabIndex        =   60
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtPhilo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5160
         TabIndex        =   59
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtPhilo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3960
         TabIndex        =   58
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtHistGeo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   9120
         TabIndex        =   57
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtHistGeo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   7800
         TabIndex        =   56
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtHistGeo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   6480
         TabIndex        =   55
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtHistGeo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5160
         TabIndex        =   54
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtHistGeo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3960
         TabIndex        =   53
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtEC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   52
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtEC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   51
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtEC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   50
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtEC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   49
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtEC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   48
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtAnglais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   9120
         TabIndex        =   47
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtAnglais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   7800
         TabIndex        =   46
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtAnglais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   6480
         TabIndex        =   45
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtAnglais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5160
         TabIndex        =   44
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtAnglais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3960
         TabIndex        =   43
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtEPS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   42
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtTM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   41
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox txtEtudText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   40
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtDictee 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   39
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox txtLitterature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2640
         TabIndex        =   38
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtLangue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   37
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtPhilo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2640
         TabIndex        =   36
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtHistGeo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2640
         TabIndex        =   35
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtEC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   34
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtAnglais 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2640
         TabIndex        =   33
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtChimie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   32
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtChimie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   31
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtChimie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   30
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtChimie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   29
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtChimie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   28
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtChimie 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   27
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtPhysiq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   26
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtPhysiq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   25
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtPhysiq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   24
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtPhysiq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   23
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtPhysiq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   22
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtPhysiq 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   21
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtSVT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   20
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtSVT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   19
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtSVT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtSVT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtSVT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtSVT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtMath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMath 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Mathématiques"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   149
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "SVT"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   148
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Physiques"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   147
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Chimie"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   146
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Anglais"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   145
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "EC"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   144
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Hist/Géo"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   143
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Philosophie"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   142
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Langue"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   141
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Littérature"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   140
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Dictée"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   139
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Etude de Texte"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   138
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "TM"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   137
         Top             =   6960
         Width           =   615
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "EPS"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   136
         Top             =   7440
         Width           =   615
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "2eLangue"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   135
         Top             =   7920
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "(All/Esp)"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   8160
         Width           =   855
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "PCT"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   133
         Top             =   8640
         Width           =   615
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Rédaction"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   132
         Top             =   9120
         Width           =   1215
      End
      Begin VB.Label lblMatiere 
         BackStyle       =   0  'Transparent
         Caption         =   "Informatique"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   131
         Top             =   9600
         Width           =   1575
      End
      Begin VB.Label lblseq 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   8880
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblseq 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   7560
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblseq 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblseq 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblseq 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblseq 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence 1 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label lblEleve 
      BackStyle       =   0  'Transparent
      Caption         =   "Liste des Elèves"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblClasse 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnx As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rstEleve As New ADODB.Recordset
Dim rstClasse As New ADODB.Recordset
Dim rstNote As New ADODB.Recordset
Dim rstSave As New ADODB.Recordset
Dim strSave As String
Dim strNote As String


Dim strInfo As String
Dim rstInfo As New ADODB.Recordset

Dim rstSaisie As New ADODB.Recordset
Dim TabCode(1000) As String
Dim Note As String
Dim NoteConvert As Double
'Dim strNote, strSaisies As String
Dim i As Integer



Private Sub dcmbClasse_Change()
Dim strEleve As String
Dim rstCount As New ADODB.Recordset

 dlistEleve.ReFill
 strEleve = "Select Matricule,Nom,Prenom From Eleve WHERE Classe= '" & dcmbClasse.BoundText & "'"
 strEleve = strEleve + " ORDER By Nom ASC "
 ExecReq strEleve, cnx, rstEleve, adOpenKeyset, adLockOptimistic, adCmdText
 
 Set dlistEleve.DataSource = rstEleve
 Set dlistEleve.RowSource = rstEleve
 dlistEleve.BoundColumn = "Matricule"
 dlistEleve.ListField = "Nom"
 txtEffectif.Text = rstEleve.RecordCount

 txtNom.Text = ""
 txtPrenom.Text = ""
 txtSexe.Text = ""
 txtStatut.Text = ""
      

      
End Sub

Private Sub dlistEleve_Click()
Dim i, val, j, k As Integer
Dim trouve As Boolean
    For k = 0 To 5
      txtMath(k) = ""
      txtPhilo(k) = ""
      txtAnglais(k) = ""
      txtChimie(k) = ""
      txtDictee(k) = ""
      txtEC(k) = ""
      txtEPS(k) = ""
      txtEtudText(k) = ""
      txtHistGeo(k) = ""
      txtLangue(k) = ""
      txtLitterature(k) = ""
      txtLangue(k) = ""
      txtLV2(k) = ""
      txtPCT(k) = ""
      txtPhysiq(k) = ""
      txtSVT(k) = ""
      txtTM(k) = ""
      txtRedac(k) = ""
      txtInfo(k) = ""
     Next k
    i = 1
    j = 1
    trouve = False
    strNote = "SELECT Note.Matricule AS Matricule ,Matiere.Libellé as Libelle,Note,Sequence,[Note].IDMatiere as CodeMatiere From Eleve, Matiere, Sequence, Classe, [Note]"
    strNote = strNote + " WHERE (((Eleve.Matricule)=[Note].[Matricule]) AND ((Matiere.NumeroMatiere)=[Note].[IDMatiere]) "
    strNote = strNote + "AND ((Sequence.NoSequence)=[Note].[Sequence]) AND ((Eleve.Classe)=[Classe].[NomClasse]) "
    strNote = strNote + " AND ((Classe.NomClasse)=[Matiere].[Classe]) "
    strNote = strNote + "AND Note.Matricule= '" + dlistEleve.BoundText + "' )"

    
    
    If ExecReq(strNote, cnx, rstNote, adOpenKeyset, adLockOptimistic, adCmdText) = True Then
    
    While Not rstNote.EOF
                 
          If (rstNote!Libelle = "Mathématiques") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtMath(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Philosophie")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtPhilo(rstNote!Sequence - 1).Text = rstNote!Note
          End If
           If (rstNote!Libelle = "Physiques") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtPhysiq(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Chimie")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtChimie(rstNote!Sequence - 1).Text = rstNote!Note
          End If
            If (rstNote!Libelle = "Anglais") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtAnglais(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Hist/Géo")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtHistGeo(rstNote!Sequence - 1).Text = rstNote!Note
          End If
           If (rstNote!Libelle = "Langue") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtLangue(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Littérature")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtLitterature(rstNote!Sequence - 1).Text = rstNote!Note
          End If
            If (rstNote!Libelle = "Dictée") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtDictee(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Etude de Texte")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtEtudText(rstNote!Sequence - 1).Text = rstNote!Note
          End If
           If (rstNote!Libelle = "TM") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtTM(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "EPS")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtEPS(rstNote!Sequence - 1).Text = rstNote!Note
          End If
           If (rstNote!Libelle = "SVT") Then
             TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
             txtSVT(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "EC")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtEC(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "2eLangue")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtLV2(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "PCT")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtPCT(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Rédaction")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtRedac(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          If ((rstNote!Libelle = "Informatique")) Then
            TabCode(rstNote!CodeMatiere) = rstNote!CodeMatiere
            txtInfo(rstNote!Sequence - 1).Text = rstNote!Note
          End If
          
         rstNote.MoveNext
    Wend
   End If
   
   ExecReq "Select * From Eleve where Matricule= '" & dlistEleve.BoundText & "'", cnx, rstInfo, adOpenKeyset, adLockOptimistic, adCmdText
     If rstInfo!Nom <> "" Then txtNom.Text = rstInfo!Nom
     
     If rstInfo!Prenom <> "" Then txtPrenom.Text = rstInfo!Prenom

     If rstInfo!Sexe <> "" Then txtSexe.Text = rstInfo!Sexe
    
      If rstInfo!Redoublant <> "" Then
       If rstInfo!Redoublant = "OUI" Then txtStatut.Text = "Redoublant"
       
       If rstInfo!Redoublant = "NON" Then txtStatut.Text = "Non Redoublant"
      
     End If
End Sub

Private Sub Form_Load()

    Taille Me, 10000, 13500
    Connexion cnx, rst
    ExecReq "SELECT * FROM Classe ", cnx, rstClasse, adOpenKeyset, adLockOptimistic, adCmdText
    'rstClasse.Open " SELECT * FROM Classe ", cnx, adOpenKeyset, adLockOptimistic, adCmdText
   
    Set dcmbClasse.DataSource = rstClasse
    Set dcmbClasse.RowSource = rstClasse
    dcmbClasse.BoundColumn = "NomClasse"
    dcmbClasse.ListField = "NomClasse"
End Sub

Function Match(ByRef Matiere As String) As Integer

rstNote.MoveFirst
While Not rstNote.EOF
If rstNote!Libelle = Matiere Then
    Match = rstNote!CodeMatiere
End If
rstNote.MoveNext
Wend

End Function
Sub MAJNote(ByRef Mat As String, seq As Integer)
Dim strCoef As String
Dim rstCoef As New ADODB.Recordset
Dim noteCoef As Double
Dim LaNote As Double

strCoef = " Select Coefficient From Matiere Where NumeroMatiere=" & Match(Mat) & ""
ExecReq strCoef, cnx, rstCoef, adOpenKeyset, adLockOptimistic, adCmdText

strSaisie = " Update [Note] Set [Note].Note= " & Note & " where Matricule = '" & dlistEleve.BoundText & "'"
strSaisie = strSaisie + " and IDMatiere=   " & Match(Mat) & " and Sequence= " & seq

ExecReq strSaisie, cnx, rstSaisie, adOpenKeyset, adLockOptimistic, adCmdText

rstCoef.Close
End Sub

Private Sub Form_Terminate()

strSave = " SELECT * FROM [Note] "
rstSave.Open strSave, cnx, adOpenKeyset, adLockOptimistic, adCmdText
rstSave.Update
rstSave.Close
'cnx.Close
End Sub

Private Sub txtAnglais_Click(Index As Integer)


Dim code As Integer
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")

 Call MAJNote("Anglais", Index + 1)
 dlistEleve_Click
End Sub

Private Sub txtChimie_Click(Index As Integer)
    

Dim code As Integer
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")

 Call MAJNote("Chimie", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtDictee_Click(Index As Integer)


Dim code As Integer

 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Dictée", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtEC_Click(Index As Integer)
    

Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")

 Call MAJNote("EC", Index + 1)
dlistEleve_Click
 

End Sub

Private Sub txtEPS_Click(Index As Integer)


Dim code As Integer
 
Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("EPS", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtEtudText_Click(Index As Integer)


Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Etude de Texte", Index + 1)

dlistEleve_Click
End Sub

Private Sub txtHistGeo_Click(Index As Integer)
    

Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Hist/Géo", Index + 1)

dlistEleve_Click
End Sub

Private Sub txtInfo_Click(Index As Integer)

Dim code As Integer
 
Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Informatique", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtLangue_Click(Index As Integer)


Dim code As Integer
 
Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Langue", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtLitterature_Click(Index As Integer)


Dim code As Integer

 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Littérature", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtLV2_Click(Index As Integer)

Dim code As Integer

Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("2e Langue", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtMath_Click(Index As Integer)

Dim code As Integer

Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
Call MAJNote("Mathématiques", Index + 1)

dlistEleve_Click
   
End Sub


Private Sub txtPCT_Click(Index As Integer)

 Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("PCT", Index + 1)
 dlistEleve_Click
End Sub

Private Sub txtPhilo_Click(Index As Integer)


Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Philosophie", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtPhysiq_Click(Index As Integer)


Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Physiques", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtRedac_Click(Index As Integer)

 Dim code As Integer
 
Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("Rédaction", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtSVT_Click(Index As Integer)


Dim code As Integer
 
 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("SVT", Index + 1)
dlistEleve_Click
End Sub

Private Sub txtTM_Click(Index As Integer)


Dim code As Integer

 Note = InputBox("ENTREZ LA NOTE", "Ges_Sco")
 Call MAJNote("TM", Index + 1)
dlistEleve_Click
End Sub
