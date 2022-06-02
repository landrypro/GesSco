VERSION 5.00
Begin VB.Form frmBulSience 
   BackColor       =   &H80000009&
   Caption         =   "BULLETIN"
   ClientHeight    =   10710
   ClientLeft      =   285
   ClientTop       =   -4650
   ClientWidth     =   13980
   Icon            =   "frmBulScience.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   13980
   Begin VB.Label lblClasse 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   207
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblMatricule 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   206
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Informatique"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   205
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   2040
      TabIndex        =   204
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   203
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4200
      TabIndex        =   202
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   6120
      TabIndex        =   201
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   7080
      TabIndex        =   200
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   8040
      TabIndex        =   199
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5160
      TabIndex        =   198
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   9720
      TabIndex        =   197
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblNom 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   196
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblPrenom 
      BackStyle       =   0  'Transparent
      Caption         =   "Prenom"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   195
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblDateNaiss 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   194
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblSexe 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   193
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblLieu 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieu"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   192
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblEffectif 
      BackStyle       =   0  'Transparent
      Caption         =   "Effectif"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   191
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUPE 3"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   190
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUPE 2"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   189
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUPE 1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   188
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   187
      Top             =   14040
      Width           =   975
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclusion"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   186
      Top             =   14040
      Width           =   735
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   185
      Top             =   13680
      Width           =   1095
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Blame"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   184
      Top             =   13680
      Width           =   735
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   183
      Top             =   13320
      Width           =   975
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   182
      Top             =   13320
      Width           =   1335
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   181
      Top             =   12960
      Width           =   1095
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Consignes"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   180
      Top             =   12960
      Width           =   1215
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   179
      Top             =   12600
      Width           =   975
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Abscences"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   178
      Top             =   12600
      Width           =   1215
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   177
      Top             =   14040
      Width           =   975
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Blame Travail"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   176
      Top             =   14040
      Width           =   1215
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   175
      Top             =   13680
      Width           =   1095
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   174
      Top             =   13680
      Width           =   1455
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   173
      Top             =   11880
      Width           =   1095
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   172
      Top             =   11880
      Width           =   735
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   171
      Top             =   11520
      Width           =   975
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Coef"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   170
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   169
      Top             =   11160
      Width           =   1095
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   168
      Top             =   11160
      Width           =   735
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "DISCIPLINE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   167
      Top             =   12240
      Width           =   3735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "TRAVAIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   166
      Top             =   12240
      Width           =   2775
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   165
      Top             =   13320
      Width           =   975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Félicitation"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   164
      Top             =   13320
      Width           =   1455
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   163
      Top             =   12960
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Encouragement"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   162
      Top             =   12960
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   161
      Top             =   12600
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Tab Honneur"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   160
      Top             =   12600
      Width           =   1575
   End
   Begin VB.Label lblMoyClasse 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   159
      Top             =   11880
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Moy Classe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   158
      Top             =   11880
      Width           =   975
   End
   Begin VB.Label lblRangEleve 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   157
      Top             =   11520
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Rang"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   156
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label lblMoyenne 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   155
      Top             =   11160
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Moyenne"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   154
      Top             =   11160
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   153
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Coef"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   152
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   151
      Top             =   10440
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   150
      Top             =   10440
      Width           =   735
   End
   Begin VB.Label lblTotauxCoef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   149
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Coef"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   148
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label lblTotauFinal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   147
      Top             =   10440
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   146
      Top             =   10440
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rappels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   145
      Top             =   10080
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Recapitulatif"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   144
      Top             =   10080
      Width           =   2775
   End
   Begin VB.Label lblTotauxG3 
      BackStyle       =   0  'Transparent
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
      Left            =   7080
      TabIndex        =   143
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label lblCoefG3 
      BackStyle       =   0  'Transparent
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
      Left            =   6120
      TabIndex        =   142
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label lblTotalG3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   141
      Top             =   9600
      Width           =   2775
   End
   Begin VB.Label lblTotauxG2 
      BackStyle       =   0  'Transparent
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
      Left            =   7080
      TabIndex        =   140
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label lblCoefG2 
      BackStyle       =   0  'Transparent
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
      Left            =   6120
      TabIndex        =   139
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label lblTotalG2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   138
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Label lblTotauxG1 
      BackStyle       =   0  'Transparent
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
      Left            =   7080
      TabIndex        =   137
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblCoefG1 
      BackStyle       =   0  'Transparent
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
      Left            =   6120
      TabIndex        =   136
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblTotalG1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total             "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   135
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   9720
      TabIndex        =   134
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   9720
      TabIndex        =   133
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   9720
      TabIndex        =   132
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   9720
      TabIndex        =   131
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   9720
      TabIndex        =   130
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   9720
      TabIndex        =   129
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   9720
      TabIndex        =   128
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9720
      TabIndex        =   127
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9720
      TabIndex        =   126
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9720
      TabIndex        =   125
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9720
      TabIndex        =   124
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9720
      TabIndex        =   123
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9720
      TabIndex        =   122
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9720
      TabIndex        =   121
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblNomProf 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Professeur"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   120
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5160
      TabIndex        =   119
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5160
      TabIndex        =   118
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5160
      TabIndex        =   117
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5160
      TabIndex        =   116
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5160
      TabIndex        =   115
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   114
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   113
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   112
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   111
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   110
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   109
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   108
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   107
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   106
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblRang 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Rang"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   105
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   104
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   8040
      TabIndex        =   103
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   8040
      TabIndex        =   102
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   8040
      TabIndex        =   101
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   8040
      TabIndex        =   100
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   99
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   8040
      TabIndex        =   98
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   8040
      TabIndex        =   97
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8040
      TabIndex        =   96
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   95
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8040
      TabIndex        =   94
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8040
      TabIndex        =   93
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   92
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   91
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblAppreciation 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Appr."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   90
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   7080
      TabIndex        =   89
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   7080
      TabIndex        =   88
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7080
      TabIndex        =   87
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7080
      TabIndex        =   86
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   7080
      TabIndex        =   85
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7080
      TabIndex        =   84
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   83
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   82
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   7080
      TabIndex        =   81
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   80
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   79
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   78
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   77
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   76
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblTotaux 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Totaux"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   75
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   6120
      TabIndex        =   74
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   6120
      TabIndex        =   73
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   72
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   71
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   6120
      TabIndex        =   70
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   69
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   68
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   67
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6120
      TabIndex        =   66
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   65
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   64
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   63
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   62
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   61
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblCoeff 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Coef"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   60
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4200
      TabIndex        =   59
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4200
      TabIndex        =   58
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4200
      TabIndex        =   57
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4200
      TabIndex        =   56
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   55
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   54
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   53
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   52
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   51
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   50
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   49
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   48
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   47
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   46
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblMoy 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Moy"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   45
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   44
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   43
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   42
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   41
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   40
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   39
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   38
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   37
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   36
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   35
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   34
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   33
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   32
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   31
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2040
      TabIndex        =   30
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   29
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   28
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   27
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   26
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   25
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   24
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   22
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   20
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   18
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "EPS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   17
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   16
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   15
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   14
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Philosophie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   13
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Hist/Géo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   12
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "EC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Langue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Littérature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Anglais"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "SVT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Chimie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Physiques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblSeq2 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSeq2"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblSeq1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSeq1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "       MATIERE "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Mathématiques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Imprimer"
   End
End
Attribute VB_Name = "frmBulSience"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nom As String
Dim rstNote As New Adodb.Recordset

Dim rstEleve As New Adodb.Recordset
Dim rstMatiere As New Adodb.Recordset
Dim rstClasse As New Adodb.Recordset
Dim rstCount As New Adodb.Recordset
Dim cnx As New Adodb.Connection

Dim strClasse As String
Dim strEleve As String
Dim strMatiere As String

Dim Moyenne As Double
Dim i, Note1, Note2 As Integer
 
'Procedure qui affiche la Note de pas Sequence d'un trimestre

Sub AffichNote(ByVal pos As Integer)
                If leTrimestre = 1 Then
                    If rstNote!Sequence = 1 Then
                     lblMat(pos).Caption = rstNote!Libelle
                     lblNoteSeq1(pos).Caption = rstNote!Note
                     Coef(pos).Caption = rstNote!Coef
                     lblSeq1.Caption = "Seq 1"
                     lblProf(pos).Caption = rstNote!NomProf
                    End If
                    If rstNote!Sequence = 2 Then
                      lblMat(pos).Caption = rstNote!Libelle
                      lblNoteSeq2(pos).Caption = rstNote!Note
                      lblSeq2.Caption = "Seq 2"
                      Coef(pos).Caption = rstNote!Coef
                      lblProf(pos).Caption = rstNote!NomProf
                    End If
                End If
                If leTrimestre = 2 Then
                    If rstNote!Sequence = 3 Then
                      lblMat(pos).Caption = rstNote!Libelle
                     lblNoteSeq1(pos).Caption = rstNote!Note
                     lblSeq1.Caption = "Seq 3"
                     Coef(pos).Caption = rstNote!Coef
                     lblProf(pos).Caption = rstNote!NomProf
                    End If
                    If rstNote!Sequence = 4 Then
                       lblMat(pos).Caption = rstNote!Libelle
                      lblNoteSeq2(pos).Caption = rstNote!Note
                      lblSeq2.Caption = "Seq 4"
                      Coef(pos).Caption = rstNote!Coef
                      lblProf(pos).Caption = rstNote!NomProf
                    End If
                End If
                If leTrimestre = 3 Then
                    If rstNote!Sequence = 5 Then
                     lblMat(pos).Caption = rstNote!Libelle
                     lblNoteSeq1(pos).Caption = Note
                     lblSeq1.Caption = "Seq 5"
                     Coef(pos).Caption = rstNote!Coef
                     lblProf(pos).Caption = rstNote!NomProf
                    End If
                    If rstNote!Sequence = 6 Then
                      lblMat(pos).Caption = rstNote!Libelle
                      lblNoteSeq2(pos).Caption = rstNote!Note
                      lblSeq2.Caption = "Seq 6"
                      Coef(pos).Caption = rstNote!Coef
                      lblProf(pos).Caption = rstNote!NomProf
                    End If
                End If
End Sub


'Procedure qui calcule la Moyenne de la Note des Sequences

Public Function MoyNote(ByVal Note1 As Double, ByVal Note2 As Double) As Double
    MoyNote = CStr((Note1 + Note2) / 2)

End Function

Private Sub Form_Load()
Taille Me, 14010, 12510
Connexion cnx, rstNote
Dim i, val, j As Integer
Dim cf, cf1, cf2, cf3 As Double
Dim tcoef, tgene, Moy As Double
Dim trouve As Boolean
     

strEleve = " SELECT Matricule,Nom,Prenom,DateNaiss,Lieu ,"
strEleve = strEleve + " Redoublant,Sexe,Classe FROM Eleve Where Matricule ='" & leMatricule & "'"

ExecReq strEleve, cnx, rstEleve

strMatiere = " SELECT Matiere.NumeroMatiere, Matiere.Libellé as Libelle, "
strMatiere = strMatiere + " Matiere.IDGroupe as Groupe , Matiere.Coefficient as Coef FROM Matiere where Classe = '" & rstEleve!Classe & "'"
ExecReq strMatiere, cnx, rstMatiere

strClasse = "Select NomClasse From Classe Where NomClasse= '" & rstEleve!Classe & "'"
ExecReq strClasse, cnx, rstClasse

strCount = " SELECT Count(Matricule) as Nbre FROM ELEVE WHERE Classe='" & rstEleve!Classe & "'"
ExecReq strCount, cnx, rstCount

'Infos d'Entete

    lblEffectif.Caption = rstCount!Nbre
    lblLieu.Caption = rstEleve!Lieu
    lblDateNaiss.Caption = rstEleve!DateNaiss
    lblMatricule.Caption = rstEleve!Matricule
    lblNom.Caption = rstEleve!Nom
    lblPrenom.Caption = rstEleve!Prenom
    lblClasse.Caption = rstEleve!Classe
    lblSexe.Caption = rstEleve!Sexe
    

    strNote = "SELECT Note.Matricule AS Matricule ,Matiere.Libellé as Libelle,Note,Sequence,[Note].IDMatiere as CodeMatiere ,Groupe.IDGroupe as Groupe,Matiere.Coefficient as Coef,NomProf FROM Groupe,Eleve, Matiere, Sequence, Classe, [Note],Trimestre, Professeur,Dispense "
    strNote = strNote + " WHERE (((Eleve.Matricule)=[Note].[Matricule]) AND ((Matiere.NumeroMatiere)=[Note].[IDMatiere]) "
    strNote = strNote + "AND ((Sequence.NoSequence)=[Note].[Sequence]) AND ((Eleve.Classe)=[Classe].[NomClasse]) "
    strNote = strNote + " AND ((Classe.NomClasse)=[Matiere].[Classe]) "
    strNote = strNote + " AND Note is not null "
    strNote = strNote + " AND ((Trimestre.IDTrimestre=Sequence.Trimestre)) "
    strNote = strNote + " AND Matiere.IDGroupe=Groupe.IDGroupe "
    strNote = strNote + " AND Matiere.NumeroMatiere=Dispense.IDMatiere "
    strNote = strNote + " AND Dispense.IDProf=Professeur.Matricule "
    strNote = strNote + " AND Note.Matricule='" + leMatricule & "'"
    strNote = strNote + " AND IDTrimestre=" & leTrimestre & " ) ORDER BY Groupe.IDGroupe,Matiere.Libellé,Sequence.NoSequence ASC "
       ' 1=(1+2)  2=(3+4) 3=5
     MsgBox strNote
    ExecReq strNote, cnx, rstNote
    
   While Not rstNote.EOF
   
   If rstNote!Libelle = "Mathématiques" Then
    AffichNote (0)
   End If
    If rstNote!Libelle = "Physiques" Then
    AffichNote (1)
   End If
   If rstNote!Libelle = "Chimie" Then
    AffichNote (2)
   End If
    If rstNote!Libelle = "SVT" Then
    AffichNote (3)
   End If
   
   If rstNote!Libelle = "Anglais" Then
    AffichNote (4)
   End If
    If rstNote!Libelle = "Littérature" Then
    AffichNote (5)
   End If
   If rstNote!Libelle = "Langue" Then
    AffichNote (6)
   End If
    If rstNote!Libelle = "EC" Then
    AffichNote (7)
   End If
   If rstNote!Libelle = "Hist/Géo" Then
    AffichNote (8)
   End If
    If rstNote!Libelle = "Philosophie" Then
    AffichNote (9)
   End If
    If rstNote!Libelle = "TM" Then
    AffichNote (12)
   End If
    If rstNote!Libelle = "EPS" Then
    AffichNote (13)
   End If
   rstNote.MoveNext
   Wend
    For i = 0 To 13
     If lblNoteSeq1(i).Caption <> "" And lblNoteSeq2(i).Caption <> "" Then
       Note1 = CDbl(lblNoteSeq1(i).Caption)
       Note2 = CDbl(lblNoteSeq2(i).Caption)
       Moyenne = MoyNote(Note1, Note2)
       lblNoteMoy(i).Caption = Moyenne
       lblTotal(i).Caption = CDbl(lblNoteMoy(i).Caption * Coef(i).Caption)
       lblAppr(i).Caption = ShowAppreciation(lblNoteMoy(i).Caption)
     End If
    Next i
    
    Dim TabCoef1(3) As Double
     Dim TabCoef2(11) As Double
     Dim TabCoef3(1) As Double
     
     For i = 0 To 3
       If Coef(i).Caption <> "" Then
         TabCoef1(i) = CInt(Coef(i).Caption)
       End If
     Next i
     
     For i = 4 To 11
       If Coef(i).Caption <> "" Then
         TabCoef2(i) = CInt(Coef(i).Caption)
       End If
     Next i
     
     For i = 0 To 1
        If Coef(i).Caption <> "" Then
         TabCoef3(i) = CInt(Coef(i).Caption)
       End If
     Next
     
     cf1 = 0
     cf2 = 0
     cf3 = 0
     For i = 0 To 3
       cf1 = cf1 + TabCoef1(i)
     Next i
     
     For i = 0 To 11
       cf2 = cf2 + TabCoef2(i)
     Next i
     
     For i = 0 To 1
       cf3 = cf3 + TabCoef2(i)
     Next i
     
     lblCoefG1.Caption = cf1
     lblCoefG2.Caption = cf2
     lblCoefG3.Caption = cf3
     
      tcoef = cf1 + cf2 + cf3
     lblTotauxCoef = tcoef
     
     
     For i = 0 To 3
       If lblTotal(i).Caption <> "" Then
         TabCoef1(i) = CDbl(lblTotal(i).Caption)
       End If
     Next i
     
     For i = 4 To 11
       If lblTotal(i).Caption <> "" Then
         TabCoef2(i) = CDbl(lblTotal(i).Caption)
       End If
     Next i
     
     For i = 0 To 1
        If lblTotal(i).Caption <> "" Then
         TabCoef3(i) = CDbl(lblTotal(i).Caption)
       End If
     Next
     
     cf1 = 0
     cf2 = 0
     cf3 = 0
     For i = 0 To 3
       cf1 = cf1 + TabCoef1(i)
     Next i
     
     For i = 0 To 11
       cf2 = cf2 + TabCoef2(i)
     Next i
     
     For i = 0 To 1
       cf3 = cf3 + TabCoef2(i)
     Next i
     
     lblTotauxG1.Caption = cf1
     lblTotauxG2.Caption = cf2
     lblTotauxG3.Caption = cf3
     
      tgene = cf1 + cf2 + cf3
     lblTotauFinal.Caption = tgene
     If tcoef <> 0 Then
      Moy = tgene / tcoef
     End If
     lblMoyenne.Caption = Moy
End Sub
Private Sub lblMat6_Click()

End Sub

Private Sub mnPrintOne_Click()
 
 
 
 
End Sub

Private Sub mnTerAll_Click()

If mnTerAll.Checked = False Then mnuPrint.Enabled = False
mnTEsp.Checked = False
mnTC.Checked = False
mnTerAll.Checked = True
If mnTerAll.Checked = True Then mnuPrint.Enabled = True

End Sub

Private Sub mnTC_Click()

If mnTC.Checked = True Then mnTC.Checked = False
If mnTC.Checked = False Then mnuPrint.Enabled = False

mnTerAll.Checked = False
mnTEsp.Checked = False
mnTC.Checked = True
If mnTC.Checked = True Then mnuPrint.Enabled = True
End Sub

Private Sub lblRg_Click(Index As Integer)

End Sub

Private Sub mnuPrint_Click()
PrintForm
End Sub
