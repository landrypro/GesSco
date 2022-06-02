VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmChoixBulletin 
   BackColor       =   &H80000013&
   Caption         =   "CHOISISSEZ !!!"
   ClientHeight    =   6870
   ClientLeft      =   3600
   ClientTop       =   3030
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Période"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
      Begin VB.OptionButton optAnnuel 
         BackColor       =   &H80000013&
         Caption         =   "Annuel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optTrimestre 
         BackColor       =   &H80000013&
         Caption         =   "Trimestriel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdAffich 
      BackColor       =   &H80000000&
      Caption         =   "Afficher"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Option "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   2775
      Begin VB.OptionButton optEleve 
         BackColor       =   &H80000013&
         Caption         =   "Par Eleve"
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
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optClasse 
         BackColor       =   &H80000013&
         Caption         =   "Par Classe"
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSDataListLib.DataList dlstEleve 
      Height          =   7305
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   12885
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList dlstClasse 
      Height          =   3840
      Left            =   360
      TabIndex        =   0
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6773
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   10980
      Left            =   0
      Picture         =   "frmChoixBulletin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19080
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "Choisissez le Trimestre"
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
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Elève"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmChoixBulletin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Dim cnx As New ADODB.Connection
Dim rstCycle As New ADODB.Recordset
Dim rstClasse As New ADODB.Recordset

Dim Lit(6) As String
Dim Sc(6) As String

Private Sub cmdAffich_Click()

estde1 = False
estde2 = False
leTrimestre = lstNum.Text
leMatricule = dlstEleve.BoundText



If rstCycle!Cycle = 1 Then
  frmBul1erCycle.Show
End If

estde1 = dlstClasse.BoundText = "2nd C" Or dlstClasse.BoundText = "2nd C1"
estde1 = estde1 Or dlstClasse.BoundText = "2nd C2" Or dlstClasse.BoundText = "P C"

estde1 = estde1 Or dlstClasse.BoundText = "P D" Or dlstClasse.BoundText = "P D1" Or dlstClasse.BoundText = "Tle C"
estde1 = estde1 Or dlstClasse.BoundText = "Tle D" Or dlstClasse.BoundText = "Tle D1"

If estde1 = True Then
  frmBulScienc.Show
End If


estde2 = dlstClasse.BoundText = "2nd A4 Esp" Or dlstClasse.BoundText = "2nd A4 Esp1"
estde2 = estde2 Or dlstClasse.BoundText = "2nd A4 All" Or dlstClasse.BoundText = "2nd A4 All1"

estde2 = estde2 Or dlstClasse.BoundText = "Tle A4 All" Or dlstClasse.BoundText = "Tle A4 Esp"
estde2 = estde2 Or dlstClasse.BoundText = "P A4 All" Or dlstClasse.BoundText = "P A4 Esp"

If estde2 = True Then
  frmBulLitter.Show
End If
'frmBulTC.Show
End Sub

Private Sub dlstClasse_Click()



strEleve = "Select Matricule,Nom from Eleve Where Classe='" & dlstClasse.BoundText & "'"
strEleve = strEleve + " Order By Nom,Prenom ASC "
ExecReq strEleve, cnx, rstClasse, adOpenKeyset, adLockOptimistic, adCmdText

Set dlstEleve.DataSource = rstClasse
Set dlstEleve.RowSource = rstClasse

dlstEleve.BoundColumn = "Matricule"
dlstEleve.ListField = "Nom"

strClasse = "Select NomClasse,Classe.IDCycle as Cycle "
strClasse = strClasse + " from Classe,Cycle Where Classe.IDCycle=Cycle.IDCycle "
strClasse = strClasse + " AND NomClasse='" & dlstClasse.BoundText & "'"

ExecReq strClasse, cnx, rstCycle, adOpenKeyset, adLockOptimistic, adCmdText


'Set dlstClasse.DataSource = rst
'Set dlstClasse.RowSource = rst
'dlstClasse.BoundColumn = "NomClasse"
'dlstClasse.ListField = "NomClasse"


End Sub

Private Sub dlstClasse_KeyUp(KeyCode As Integer, Shift As Integer)
rstClasse.Close
End Sub

Private Sub dlstEleve_Click()
leMatricule = dlstEleve.BoundText
End Sub

Private Sub Form_Load()

Connexion cnx, rst
strClasse = "Select NomClasse,Classe.IDCycle as Cycle from Classe,Cycle Where Classe.IDCycle=Cycle.IDCycle"
ExecReq strClasse, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
Set dlstClasse.DataSource = rst
Set dlstClasse.RowSource = rst
dlstClasse.BoundColumn = "NomClasse"
dlstClasse.ListField = "NomClasse"
optEleve.Value = True
End Sub

Private Sub Form_Terminate()
cnx.Close
End Sub

Private Sub lstNum_Click()
leTrimestre = lstNum.Text

End Sub

Private Sub optAnnuel_Click()
lstNum.Enabled = False
End Sub

Private Sub optClasse_Click()
dlstEleve.Enabled = False
End Sub

Private Sub optEleve_Click()
dlstEleve.Enabled = True
End Sub


Private Sub optTrimestre_Click()
lstNum.Clear
lstNum.Enabled = True
lstNum.AddItem ("1")
lstNum.AddItem ("2")
lstNum.AddItem ("3")
End Sub


