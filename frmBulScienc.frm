VERSION 5.00
Begin VB.Form frmBulScienc 
   BackColor       =   &H80000009&
   Caption         =   "BULLETIN"
   ClientHeight    =   10710
   ClientLeft      =   285
   ClientTop       =   -14520
   ClientWidth     =   14355
   Icon            =   "frmBulScienc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   14355
   Begin VB.Line Line27 
      X1              =   120
      X2              =   11400
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line26 
      X1              =   120
      X2              =   11400
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line10 
      X1              =   3120
      X2              =   3120
      Y1              =   10680
      Y2              =   14040
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   7800
      Y1              =   12720
      Y2              =   12720
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   7800
      Y1              =   12480
      Y2              =   12480
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   7800
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   11400
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   11400
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11400
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Noms et prenoms de l'élève:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   200
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Date et lieu de naissance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   199
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   198
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   197
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieu:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   196
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   195
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Effectif:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   194
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBulScienc.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   193
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBulScienc.frx":04F0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8040
      TabIndex        =   192
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "BULLETIN DE NOTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   191
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Recor Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   190
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   11400
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   11400
      Y1              =   10680
      Y2              =   10680
   End
   Begin VB.Line Line11 
      X1              =   2160
      X2              =   2160
      Y1              =   3360
      Y2              =   5400
   End
   Begin VB.Line Line12 
      X1              =   2160
      X2              =   2160
      Y1              =   5880
      Y2              =   8280
   End
   Begin VB.Line Line13 
      X1              =   9360
      X2              =   9360
      Y1              =   3360
      Y2              =   10320
   End
   Begin VB.Line Line14 
      X1              =   7800
      X2              =   7800
      Y1              =   3360
      Y2              =   14040
   End
   Begin VB.Line Line15 
      X1              =   6480
      X2              =   6480
      Y1              =   3360
      Y2              =   10320
   End
   Begin VB.Line Line16 
      X1              =   5640
      X2              =   5640
      Y1              =   3360
      Y2              =   10320
   End
   Begin VB.Line Line17 
      X1              =   4800
      X2              =   4800
      Y1              =   3360
      Y2              =   10320
   End
   Begin VB.Line Line18 
      X1              =   3720
      X2              =   3720
      Y1              =   3360
      Y2              =   5400
   End
   Begin VB.Line Line19 
      X1              =   3720
      X2              =   3720
      Y1              =   5880
      Y2              =   8280
   End
   Begin VB.Line Line20 
      X1              =   2040
      X2              =   2040
      Y1              =   8760
      Y2              =   10320
   End
   Begin VB.Line Line21 
      X1              =   3720
      X2              =   3720
      Y1              =   8760
      Y2              =   10320
   End
   Begin VB.Line Line22 
      X1              =   11400
      X2              =   11400
      Y1              =   3360
      Y2              =   10320
   End
   Begin VB.Line Line23 
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   14040
   End
   Begin VB.Line Line24 
      X1              =   11400
      X2              =   11400
      Y1              =   10680
      Y2              =   14040
   End
   Begin VB.Line Line25 
      X1              =   120
      X2              =   11400
      Y1              =   14040
      Y2              =   14040
   End
   Begin VB.Line Line28 
      X1              =   7800
      X2              =   11400
      Y1              =   11040
      Y2              =   11040
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Observations et visa du principal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   189
      Top             =   10680
      Width           =   3255
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Travail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   188
      Top             =   12480
      Width           =   1455
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Discipline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   187
      Top             =   12480
      Width           =   975
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Tableau d'honneur"
      Height          =   255
      Left            =   360
      TabIndex        =   186
      Top             =   12840
      Width           =   1575
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Encouragement"
      Height          =   255
      Left            =   360
      TabIndex        =   185
      Top             =   13080
      Width           =   1575
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Félécitations"
      Height          =   255
      Left            =   360
      TabIndex        =   184
      Top             =   13320
      Width           =   1575
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement"
      Height          =   255
      Left            =   360
      TabIndex        =   183
      Top             =   13560
      Width           =   1575
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Absences"
      Height          =   255
      Left            =   3960
      TabIndex        =   182
      Top             =   12840
      Width           =   1095
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Consignes"
      Height          =   255
      Left            =   3960
      TabIndex        =   181
      Top             =   13080
      Width           =   1095
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissements"
      Height          =   255
      Left            =   3960
      TabIndex        =   180
      Top             =   13320
      Width           =   1095
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Blame"
      Height          =   255
      Left            =   3960
      TabIndex        =   179
      Top             =   13560
      Width           =   1095
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Exclusion"
      Height          =   255
      Left            =   3960
      TabIndex        =   178
      Top             =   13800
      Width           =   1095
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Blame travail"
      Height          =   255
      Left            =   360
      TabIndex        =   177
      Top             =   13800
      Width           =   1575
   End
   Begin VB.Line Line29 
      X1              =   1800
      X2              =   1800
      Y1              =   11160
      Y2              =   12480
   End
   Begin VB.Line Line30 
      X1              =   5760
      X2              =   5760
      Y1              =   11160
      Y2              =   12480
   End
   Begin VB.Line Line33 
      X1              =   120
      X2              =   11400
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line34 
      X1              =   0
      X2              =   11400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line35 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Line Line36 
      X1              =   11400
      X2              =   11400
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Recapitulatif"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   176
      Top             =   10800
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rappels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   175
      Top             =   10800
      Width           =   1335
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
      TabIndex        =   174
      Top             =   11160
      Width           =   735
   End
   Begin VB.Label lblTotauFinal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   173
      Top             =   11280
      Width           =   735
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
      TabIndex        =   172
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label lblTotauxCoef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   171
      Top             =   11640
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Trimestre 1"
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
      Left            =   4320
      TabIndex        =   170
      Top             =   11160
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   169
      Top             =   11280
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Trimestre 2"
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
      Left            =   4320
      TabIndex        =   168
      Top             =   11520
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   167
      Top             =   11640
      Width           =   1095
   End
   Begin VB.Label lblMoyenne 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   166
      Top             =   12240
      Width           =   735
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
      TabIndex        =   165
      Top             =   11880
      Width           =   735
   End
   Begin VB.Label lblRangEleve 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   164
      Top             =   11880
      Width           =   735
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
      TabIndex        =   163
      Top             =   12240
      Width           =   1455
   End
   Begin VB.Label lblMoyClasse 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   162
      Top             =   13560
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Trimestre 3"
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
      Left            =   4320
      TabIndex        =   161
      Top             =   11880
      Width           =   1215
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   160
      Top             =   12000
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Moyenne An."
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
      Left            =   4320
      TabIndex        =   159
      Top             =   12240
      Width           =   1335
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   158
      Top             =   12360
      Width           =   1095
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
      Left            =   6960
      TabIndex        =   157
      Top             =   13560
      Width           =   1095
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
      Left            =   1560
      TabIndex        =   156
      Top             =   720
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
      Left            =   1560
      TabIndex        =   155
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   154
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   153
      Top             =   7920
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
      Left            =   3600
      TabIndex        =   152
      Top             =   480
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
      Left            =   3600
      TabIndex        =   151
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   150
      Top             =   7920
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
      Left            =   4560
      TabIndex        =   149
      Top             =   480
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
      Left            =   4560
      TabIndex        =   148
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   6480
      TabIndex        =   147
      Top             =   7920
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
      Left            =   6480
      TabIndex        =   146
      Top             =   480
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
      Left            =   6480
      TabIndex        =   145
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   7440
      TabIndex        =   144
      Top             =   7920
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
      Left            =   7440
      TabIndex        =   143
      Top             =   480
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
      Left            =   7440
      TabIndex        =   142
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   8400
      TabIndex        =   141
      Top             =   7920
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
      Left            =   8400
      TabIndex        =   140
      Top             =   480
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
      Left            =   8400
      TabIndex        =   139
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblRg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5520
      TabIndex        =   138
      Top             =   7920
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
      Left            =   5520
      TabIndex        =   137
      Top             =   480
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
      Left            =   5520
      TabIndex        =   136
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   10080
      TabIndex        =   135
      Top             =   7920
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
      Left            =   10080
      TabIndex        =   134
      Top             =   480
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
      Left            =   10080
      TabIndex        =   133
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Informatique"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   132
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   131
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   130
      Top             =   9720
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4920
      TabIndex        =   129
      Top             =   9720
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5760
      TabIndex        =   128
      Top             =   9720
      Width           =   615
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   6720
      TabIndex        =   127
      Top             =   9720
      Width           =   735
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   7800
      TabIndex        =   126
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   9840
      TabIndex        =   125
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label lblClasse 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   124
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblMatricule 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   123
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblEffectif 
      BackStyle       =   0  'Transparent
      Caption         =   "Effectif"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   122
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblLieu 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieu"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   121
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblSexe 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   120
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblDateNaiss 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   119
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblPrenom 
      BackStyle       =   0  'Transparent
      Caption         =   "Prenom"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   118
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblNom 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   117
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUPE 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   116
      Top             =   10320
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUPE 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   115
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUPE 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   114
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
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
      Left            =   6000
      TabIndex        =   113
      Top             =   15240
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
      Left            =   4200
      TabIndex        =   112
      Top             =   15240
      Width           =   1095
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
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
      Left            =   6000
      TabIndex        =   111
      Top             =   14880
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
      Left            =   4200
      TabIndex        =   110
      Top             =   14880
      Width           =   1215
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   109
      Top             =   15240
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
      Left            =   360
      TabIndex        =   108
      Top             =   15240
      Width           =   1215
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   107
      Top             =   14880
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
      Left            =   360
      TabIndex        =   106
      Top             =   14880
      Width           =   1455
   End
   Begin VB.Label lblTotauxG3 
      BackStyle       =   0  'Transparent
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
      Left            =   6600
      TabIndex        =   105
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblCoefG3 
      BackStyle       =   0  'Transparent
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
      Left            =   5760
      TabIndex        =   104
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblTotalG3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   103
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Label lblTotauxG2 
      BackStyle       =   0  'Transparent
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
      Left            =   6720
      TabIndex        =   102
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblCoefG2 
      BackStyle       =   0  'Transparent
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
      Left            =   5760
      TabIndex        =   101
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblTotalG2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   100
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Label lblTotauxG1 
      BackStyle       =   0  'Transparent
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
      Left            =   6840
      TabIndex        =   99
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblCoefG1 
      BackStyle       =   0  'Transparent
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
      Left            =   5880
      TabIndex        =   98
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label lblTotalG1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total             "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   97
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   9840
      TabIndex        =   96
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   9840
      TabIndex        =   95
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   94
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   93
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   92
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   91
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   90
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9960
      TabIndex        =   89
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9960
      TabIndex        =   88
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9960
      TabIndex        =   87
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   86
      Top             =   3960
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
      Left            =   9960
      TabIndex        =   85
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   7800
      TabIndex        =   84
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   7800
      TabIndex        =   83
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7680
      TabIndex        =   82
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7680
      TabIndex        =   81
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   7680
      TabIndex        =   80
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   7680
      TabIndex        =   79
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7680
      TabIndex        =   78
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   77
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   76
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   75
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblAppr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   74
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblAppreciation 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Appreciation"
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
      Left            =   7920
      TabIndex        =   73
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   6720
      TabIndex        =   72
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   6720
      TabIndex        =   71
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   70
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   69
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   68
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   67
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   66
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   65
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   64
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   63
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   62
      Top             =   3960
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
      Left            =   6840
      TabIndex        =   61
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5760
      TabIndex        =   60
      Top             =   9360
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   59
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   58
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   57
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   56
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   55
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   54
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   53
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   52
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   51
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Coef 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   50
      Top             =   3960
      Width           =   615
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
      Left            =   5880
      TabIndex        =   49
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   48
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   47
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   46
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   45
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   44
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   43
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   42
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   41
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   40
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   39
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblNoteMoy 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   38
      Top             =   3960
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
      Left            =   5040
      TabIndex        =   37
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   36
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3960
      TabIndex        =   35
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   34
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   33
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   32
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   31
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   30
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   29
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   28
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   27
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2760
      TabIndex        =   26
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2760
      TabIndex        =   25
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   24
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   23
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   22
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   21
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   20
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   18
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblNoteSeq2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   16
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "EPS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   15
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   14
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Philosophie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   13
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Hist/Géo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   12
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "EC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   11
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Langue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Littérature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Anglais"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "SVT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Chimie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Physiques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblNoteSeq1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   3960
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
      Left            =   4080
      TabIndex        =   3
      Top             =   3480
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
      Left            =   2880
      TabIndex        =   2
      Top             =   3480
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
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "Mathématiques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Imprimer"
   End
End
Attribute VB_Name = "frmBulScienc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nom As String

Dim cnx As New ADODB.Connection

Dim rstNote As New ADODB.Recordset
Dim rstEleve As New ADODB.Recordset
Dim rstMatiere As New ADODB.Recordset
Dim rstClasse As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset




Dim strClasse As String
Dim strEleve As String
Dim strMatiere As String
Dim strCount As String

Dim Moyenne As String
Dim i, Note1, Note2 As Double
 
'Procedure qui affiche la Note de pas Sequence d'un trimestre

 Sub AffichNote(ByVal pos As Integer)
              
                If leTrimestre = 1 Then
                    If rstNote!Sequence = 1 Then
                     lblMat(pos).Caption = rstNote!Libelle
                     lblNoteSeq1(pos).Caption = rstNote!Note
                     Coef(pos).Caption = rstNote!Note
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
                     lblNoteSeq1(pos).Caption = rstNote!Note
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
 Dim cf, cf1, cf2, cf3, cf4 As Double
Dim i, val, j, k As Integer
Dim tcoef, tgene, Moy As Double
Dim trouve As Boolean
Dim totalcoef As Integer
Dim Note As Integer

Taille Me, 15000, 5000


strEleve = " SELECT Matricule,Nom,Prenom,DateNaiss,Lieu ,"
strEleve = strEleve + " Redoublant,Sexe,Classe FROM Eleve Where Matricule ='" & leMatricule & "'"

ExecReq strEleve, cnx, rstEleve, adOpenKeyset, adLockOptimistic, adCmdText


strMatiere = " SELECT Matiere.NumeroMatiere, Matiere.Libellé as Libelle, "
strMatiere = strMatiere + " Matiere.IDGroupe as Groupe , Matiere.Coefficient as Coef FROM Matiere where Classe = '" & rstEleve!Classe & "'"
ExecReq strMatiere, cnx, rstMatiere, adOpenKeyset, adLockOptimistic, adCmdText

strClasse = "Select NomClasse From Classe Where NomClasse= '" & rstEleve!Classe & "'"
ExecReq strClasse, cnx, rstClasse, adOpenKeyset, adLockOptimistic, adCmdText

strCount = " SELECT Count(Matricule) as Nbre FROM ELEVE WHERE Classe='" & rstEleve!Classe & "'"
ExecReq strCount, cnx, rstCount, adOpenKeyset, adLockOptimistic, adCmdText

'Infos d'Entete

    lblEffectif.Caption = rstCount!Nbre
    lblLieu.Caption = rstEleve!Lieu
    lblDateNaiss.Caption = rstEleve!DateNaiss
    lblMatricule.Caption = rstEleve!Matricule
    lblNom.Caption = rstEleve!Nom
    lblPrenom.Caption = rstEleve!Prenom
    lblClasse.Caption = rstEleve!Classe
    lblSexe.Caption = rstEleve!Sexe
    

    strNote = "SELECT Note.Matricule AS Matricule ,Matiere.Libellé as Libelle,Note,NoteCoef as NoteCoefficié,Sequence,[Note].IDMatiere as CodeMatiere ,Groupe.IDGroupe as Groupe,Matiere.Coefficient as Coef,NomProf FROM Groupe,Eleve, Matiere, Sequence, Classe, [Note],Trimestre, Professeur,Dispense "
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
  'MsgBox strNote
       ' 1=(1+2)  2=(3+4) 3=5
     
  
   ExecReq strNote, cnx, rstNote, adOpenKeyset, adLockOptimistic, adCmdText
     
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
   
     For i = 0 To 14
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
     Dim TabCoef3(15) As Double
     
     'Affiche Coefficient
     
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
     
     For i = 12 To 14
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
     
     For i = 12 To 14
       cf3 = cf3 + TabCoef3(i)
     Next i
     
     lblCoefG1.Caption = cf1
     lblCoefG2.Caption = cf2
     lblCoefG3.Caption = cf3
     
     tcoef = cf1 + cf2 + cf3
     lblTotauxCoef.Caption = tcoef
     
       'Affiche Note Total
     
     
     cf1 = 0
     cf2 = 0
     cf3 = 0
     
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
     
     For i = 12 To 14
        If lblTotal(i).Caption <> "" Then
         TabCoef3(i) = CDbl(lblTotal(i).Caption)
       End If
     Next
     
   
     For i = 0 To 3
       cf1 = cf1 + TabCoef1(i)
     Next i
     
     For i = 0 To 11
       cf2 = cf2 + TabCoef2(i)
     Next i
     
     For i = 12 To 14
       cf3 = cf3 + TabCoef3(i)
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
     range (lblTotauFinal.Caption)
     
   
     
End Sub

'Determine le Rang a l'aide d'une moyenne

Sub range(ByRef Moy As Double)
Dim rstMoyClasse As New ADODB.Recordset
Dim strMoyenneClasse As String
Dim rang As Integer
Dim NbreEleve As Integer


strMoyenneClasse = " SELECT Eleve.Classe as Classe ,Note.Matricule AS Matricule , "
strMoyenneClasse = strMoyenneClasse + " Sum(NoteCoef)/2 as TotalNote"
strMoyenneClasse = strMoyenneClasse + " From Groupe, Eleve, Matiere, Sequence, Classe, [Note], Trimestre, Professeur, Dispense"
strMoyenneClasse = strMoyenneClasse + " Where Eleve.Matricule = Note.Matricule And Matiere.NumeroMatiere=Note.IDMatiere"
strMoyenneClasse = strMoyenneClasse + " And Sequence.NoSequence=Note.Sequence And Eleve.Classe=Classe.NomClasse"
strMoyenneClasse = strMoyenneClasse + " And Classe.NomClasse=Matiere.Classe And Trimestre.IDTrimestre=Sequence.Trimestre"
strMoyenneClasse = strMoyenneClasse + " And Matiere.IDGroupe=Groupe.IDGroupe And Matiere.NumeroMatiere=Dispense.IDMatiere"
strMoyenneClasse = strMoyenneClasse + " And Dispense.IDProf=Professeur.Matricule AND Trimestre.IDTrimestre=" & leTrimestre
strMoyenneClasse = strMoyenneClasse + " AND Eleve.Classe='" & rstEleve!Classe & "' "
strMoyenneClasse = strMoyenneClasse + " Group By Eleve.Classe,Note.Matricule HAVING Sum(NoteCoef)"
strMoyenneClasse = strMoyenneClasse + " Order By Sum(NoteCoef)/2  DESC"
ExecReq strMoyenneClasse, cnx, rstMoyClasse, adOpenKeyset, adLockOptimistic, adCmdText
rang = 1
MoyenClass = 0
NbreEleve = 0
While Not rstMoyClasse.EOF
 NbreEleve = NbreEleve + 1
 MoyenClass = MoyenClass + rstMoyClasse!TotalNote
 If Moy < rstMoyClasse!TotalNote Then
     rang = rang + 1
 End If
 rstMoyClasse.MoveNext
Wend
If NbreEleve <> 0 Then
MoyenClass = (MoyenClass / NbreEleve)
End If
MoyenneClasse lblTotauxCoef
lblRangEleve.Caption = CStr(rang) + "e/" + lblEffectif.Caption

End Sub

 Sub MoyenneClasse(ByVal Coef As Integer)
   If Coef <> 0 Then
   MoyenClass = MoyenClass / Coef
   End If
   lblMoyClasse.Caption = MoyenClass
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

Private Sub Form_Terminate()
cnx.Close
End Sub


Private Sub mnuPrint_Click()
PrintForm
Printer.Font = "Arial"
Printer.FontSize = 10
Printer.FontBold = True

'Printer.PaperSize = VB
Printer.Height = 13000
Printer.Width = 14940

End Sub

