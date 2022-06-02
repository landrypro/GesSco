VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAjoutEleve 
   BackColor       =   &H80000013&
   Caption         =   "Ajout d'Elèves"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   Icon            =   "frmAjoutEleve.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Parcourir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   39
      Top             =   3840
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   15600
      Picture         =   "frmAjoutEleve.frx":0442
      ScaleHeight     =   2595
      ScaleWidth      =   3195
      TabIndex        =   38
      Top             =   960
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2280
      TabIndex        =   34
      Top             =   3360
      Width           =   1575
      Begin VB.OptionButton optM 
         BackColor       =   &H80000013&
         Caption         =   "M"
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
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optF 
         BackColor       =   &H80000013&
         Caption         =   "F"
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
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "OK"
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H80000009&
      Caption         =   "Annuler"
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdModifEleve 
      BackColor       =   &H80000009&
      Caption         =   "Modifier"
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10560
      TabIndex        =   28
      Top             =   5880
      Width           =   2535
      Begin VB.OptionButton optNon 
         BackColor       =   &H80000013&
         Caption         =   "Non"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton optOui 
         BackColor       =   &H80000013&
         Caption         =   "Oui"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdListeEleve 
      BackColor       =   &H000000FF&
      Caption         =   "Liste des Eleve"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6600
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo dcmbClasse 
      Height          =   360
      Left            =   10560
      TabIndex        =   26
      Top             =   6600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Choisissez La Classe"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2655
      Left            =   1080
      TabIndex        =   25
      Top             =   5040
      Width           =   4575
      _Version        =   524288
      _ExtentX        =   8070
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   65280
      Year            =   2008
      Month           =   10
      Day             =   17
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtNomMere 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   10440
      TabIndex        =   24
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtRedoublant 
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
      Left            =   10560
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtAdrParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   10440
      TabIndex        =   9
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox txtNomPere 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Left            =   10440
      TabIndex        =   8
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtReligion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Left            =   10440
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtNation 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Left            =   10440
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtLieu 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtDateNaiss 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txtSexe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPrenom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtNom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtMatricule 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Height          =   3135
      Left            =   15720
      TabIndex        =   37
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
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
      Left            =   8640
      TabIndex        =   23
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Redoublant"
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
      Left            =   8520
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse Parents"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de la Mere"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du Pere"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Réligion"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationalité"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieu"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date de Naissance"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prenom"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
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
      Left            =   1080
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
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
      Left            =   960
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmAjoutEleve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEleve As ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim strStatut, strSexe As String
Dim rstMatiere As New ADODB.Recordset

Dim cnx As ADODB.Connection
Dim strInsert, strValue, strEleve, sqlelev As String

Private Sub Calendar1_Click()
txtDateNaiss.Text = CStr(Calendar1.Day) + "/" + CStr(Calendar1.Month) + "/" + CStr(Calendar1.Year)
End Sub

Private Sub cmdListe_Click()

End Sub

Private Sub cmdListeEleve_Click()
Me.Hide
Unload frmAjoutEleve

frmListeEleve.Show

End Sub

Private Sub cmdModifEleve_Click()
frmModifEleve.Show
End Sub

Private Sub cmdOK_Click()

strInsert = " Insert Into Eleve ( Matricule,Nom,Prenom,Sexe,DateNaiss,Lieu, "
strInsert = strInsert + " Nationalite,Religion,NomPere,NomMere,AdressParents,Redoublant,Classe  )"
strValue = " Values (  '" + txtMatricule.Text + "','" + txtNom.Text + "','" + txtPrenom.Text + "','" + strSexe + "','"
strValue = strValue + txtDateNaiss.Text + "',' " + txtLieu.Text + "','" + txtNation.Text + "','" + txtReligion.Text + " ' "
strValue = strValue + ",'" + txtNomPere.Text + "','" + txtNomMere.Text + "'"
strValue = strValue + ",'" + txtAdrParent.Text + "','" + strStatut + "','" + dcmbClasse.BoundText + "'  )"
strEleve = strInsert + strValue
'MsgBox strEleve
'Text1.Text = strEleve
'rstEleve.Open strEleve, cnx
'rstMatiere.Open strEleve, cnx, adOpenKeyset, adLockOptimistic, adCmdText
ExecReq strEleve, cnx, rstMatiere, adOpenKeyset, adLockOptimistic, adCmdText
sqlelev = " INSERT INTO [Note] SELECT Eleve.matricule AS matricule, Matiere.NumeroMatiere AS idmatiere, Sequence.NoSequence AS Sequence, AnneeScolaire.NoAnne AS NoAnnee "
sqlelev = sqlelev + " From Eleve, Matiere, Sequence, AnneeScolaire "
sqlelev = sqlelev + "  WHERE (Eleve.Classe=Matiere.Classe) AND Eleve.matricule not in (Select Matricule FROM [Note]); "
'MsgBox sqlelev
ExecReq sqlelev, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
'rst.Open sqlelev, cnx

strClasse = "Select NomClasse from Classe"
ExecReq strClasse, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
'rst.Open strClass, cnx

Set dcmbClasse.DataSource = rst
Set dcmbClasse.RowSource = rst
dcmbClasse.BoundColumn = "NomClasse"
dcmbClasse.ListField = "NomClasse"


End Sub



Private Sub Form_Load()
'CenterForm frmAjoutEleve
Connexion cnx, rstEleve
strClasse = "Select NomClasse from Classe"
ExecReq strClasse, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
If rst.RecordCount = 0 Then
 MsgBox "Vous n'avez pas encore enregistrer des Classes", vbCritical, "Ges_Sco"
End If
If rst.RecordCount <> 0 Then
  Taille frmAjoutEleve, 10000, 12255
  Set dcmbClasse.DataSource = rst
  Set dcmbClasse.RowSource = rst
  dcmbClasse.BoundColumn = "NomClasse"
  dcmbClasse.ListField = "NomClasse"
  strStatut = ""
  strSexe = ""
End If

End Sub

Private Sub Form_Terminate()
cnx.Close
End Sub

Private Sub optF_Click()
strSexe = "F"
End Sub

Private Sub optM_Click()
strSexe = "M"
End Sub

Private Sub optNon_Click()
strStatut = "NON"
End Sub

Private Sub optOui_Click()
strStatut = "OUI"
End Sub
