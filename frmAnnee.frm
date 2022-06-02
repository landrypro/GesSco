VERSION 5.00
Begin VB.Form frmAnnee 
   BackColor       =   &H80000009&
   Caption         =   "Ajouter des Années"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   Icon            =   "frmAnnee.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAnnee.frx":0442
   ScaleHeight     =   2670
   ScaleWidth      =   6330
   WindowState     =   2  'Maximized
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
      Left            =   1080
      Picture         =   "frmAnnee.frx":2F73
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
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
      Left            =   3480
      Picture         =   "frmAnnee.frx":631F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtAnnee 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Année"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmAnnee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstAnnee As ADODB.Recordset

Dim cnx As ADODB.Connection
Dim strAnne, strValue, strAnnee  As String

Private Sub cmdOK_Click()
strInsert = " Insert Into Annee ( Annee) "
strValue = " Values ( " + txtAnnee.Text + " )"

strAnnee = strInsert + strValue
MsgBox strAnnee
ExecReq strAnnee, cnx, rstAnnee, adOpenKeyset, adLockOptimistic, adCmdText
'rstAnnee.Open strAnnee, cnx

End Sub

Private Sub Form_Load()
'CenterForm frmAjoutEleve
Taille frmAnnee, 4695, 6330
Connexion cnx, rstAnnee
End Sub


