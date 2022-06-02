VERSION 5.00
Begin VB.Form frmMatiere 
   Caption         =   "Paramétrages des Matières"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   7755
   Begin VB.CommandButton cmdOK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Picture         =   "frmMatiere.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnnuler 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Picture         =   "frmMatiere.frx":33AC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Année"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Groupe de la Matière"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Classe"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Coefficient"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Libellé"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Code Matiere"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmMatiere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
