VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmModifEleve 
   BackColor       =   &H80000013&
   Caption         =   "Modifier un Eléve"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13860
   BeginProperty Font 
      Name            =   "Script MT Bold"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModifEleve.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   13860
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtMatricule 
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
      Left            =   2160
      TabIndex        =   32
      Top             =   1080
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo dcmbClasse1 
      Height          =   390
      Left            =   12120
      TabIndex        =   31
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcmbClasse2 
      Height          =   390
      Left            =   7560
      TabIndex        =   30
      Top             =   5520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdConfir 
      BackColor       =   &H80000013&
      Caption         =   "Confirmer la modification"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   24
      Top             =   4200
      Width           =   2295
      Begin VB.OptionButton optOui 
         BackColor       =   &H80000013&
         Caption         =   "OUI"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optNon 
         BackColor       =   &H80000013&
         Caption         =   "NON"
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
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtMere 
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
      Left            =   7200
      TabIndex        =   21
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtPere 
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
      Left            =   7200
      TabIndex        =   19
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
      Begin VB.OptionButton optM 
         BackColor       =   &H80000013&
         Caption         =   "M"
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
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optF 
         BackColor       =   &H80000013&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtReligion 
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
      Left            =   7080
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtNation 
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
      Left            =   2280
      TabIndex        =   6
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtAdrP 
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
      Left            =   7560
      TabIndex        =   5
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txtLieu 
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
      Left            =   2280
      TabIndex        =   4
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtDOB 
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
      Left            =   2280
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtPrenom 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtNom 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin MSDataListLib.DataList dlstEleve 
      Height          =   3435
      Left            =   12240
      TabIndex        =   0
      Top             =   1920
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6059
      _Version        =   393216
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Matricule"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   33
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Classes"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   29
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Redoublant"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de la Mere"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du Pere"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Réligion"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationalité"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label txtAdrParent 
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse Parents"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieu"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date de Naissance"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Prenom"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "frmModifEleve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnx As ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rstClasse As New ADODB.Recordset
Dim rstEleve As New ADODB.Recordset
Dim rstUpdate As New ADODB.Recordset
Dim strEleve As String

Private Sub dcmdClasse_Change()
 strEleve = "Select Matricule,Nom,Prenom From Eleve WHERE Classe='" & dcmbClasse1.BoundText & "'"
 ExecReq strEleve, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
 'MsgBox strEleve dlstEleve
Set dlstEleve.DataSource = rst
Set dlstEleve.RowSource = rst
 dlstEleve.BoundColumn = "Matricule"
dlstEleve.ListField = "Nom"
End Sub

Private Sub cmdConfir_Click()
  Dim strUpdate As String
  
   
   strUpdate = " Update Eleve set Matricule='" & txtMatricule & "'"
   strUpdate = strUpdate + " , Nom='" & txtNom & "'"
   strUpdate = strUpdate + " , Prenom= '" & txtPrenom & "'"
   
   If optF.Value = True Then
       strUpdate = strUpdate + " ,  Sexe= 'F'"
   End If
   If optM.Value = True Then
       strUpdate = strUpdate + " , Sexe= 'M'"
   End If
   strUpdate = strUpdate + " , DateNaiss= '" & txtDOB & "'"
   strUpdate = strUpdate + " ,  Lieu= '" & txtLieu & "'"
   strUpdate = strUpdate + " ,  Nationalite= '" & txtNation & "'"
   strUpdate = strUpdate + ", Religion = '" & txtReligion & "'"
   strUpdate = strUpdate + ", NomPere= '" & txtPere & "'"
   strUpdate = strUpdate + " , NomMere = '" & txtMere & "'"
   strUpdate = strUpdate + " , AdressParents = '" & txtAdrP & "'"
   If optOui.Value = True Then
       strUpdate = strUpdate + " , Redoublant =  'OUI'"
   End If
   If optNon.Value = True Then
       strUpdate = strUpdate + " , Redoublant =  'NON'"
   End If
   strUpdate = strUpdate + " , Classe='" & dcmbClasse2.Text & " ' "
   strUpdate = strUpdate + " Where Matricule='" & dlstEleve.BoundText & "'"
   MsgBox strUpdate
   ExecReq strUpdate, cnx, rstUpdate, adOpenKeyset, adLockOptimistic, adCmdText

  'Remplissage du Datalistbox
   strEleve = "Select Matricule,Nom,Prenom From Eleve WHERE Classe='" & dcmbClasse1.BoundText & "'"
   ExecReq strEleve, cnx, rstEleve, adOpenKeyset, adLockOptimistic, adCmdText
   'MsgBox strEleve
   Set dlstEleve.DataSource = rstEleve
   Set dlstEleve.RowSource = rstEleve
   dlstEleve.BoundColumn = "Matricule"
   dlstEleve.ListField = "Nom"
   dlstEleve.Refresh
End Sub

Private Sub dcmbClasse1_Change()
  strEleve = "Select Matricule,Nom,Prenom From Eleve WHERE Classe= '" & dcmbClasse1.BoundText & "'"
  ExecReq strEleve, cnx, rstEleve, adOpenKeyset, adLockOptimistic, adCmdText
 'MsgBox strEleve
 Set dlstEleve.DataSource = rstEleve
 Set dlstEleve.RowSource = rstEleve
 dlstEleve.BoundColumn = "Matricule"
 dlstEleve.ListField = "Nom"
End Sub

Private Sub dlstEleve_Click()
  
  strEleve = "Select * From Eleve Where Matricule= '" & dlstEleve.BoundText & "'"
  ExecReq strEleve, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
   txtMatricule = rst!Matricule
   txtNom = rst!Nom
   txtPrenom = rst!Prenom
   
   If rst!Sexe = "F" Then optF.Value = Checked
   If rst!Sexe = "M" Then optM.Value = Checked
   
   txtDOB = rst!DateNaiss
   txtLieu = rst!Lieu
   txtNation = rst!Nationalite
   txtReligion = rst!Religion
   txtPere = rst!NomPere
   txtMere = rst!NomMere
   txtAdrP = rst!AdressParents
   '------Classe de l'eleve-------
   'dcmbClasse2.Text=rst!
   If rst!Redoublant = "OUI" Then optOui.Value = Checked
   If rst!Redoublant = "NON" Then optNon.Value = Checked
   dcmbClasse2.Text = rst!Classe
  'Nom Prenom Sexe DateNaiss Lieu  Nationalite Religion
  'NomPere NomMere AdressParents Redoublant  Classe
End Sub

Private Sub Form_Load()
'CenterForm frmAjoutEleve
'Taille Me, 10905, 6180
Taille Me, 6180, 10905
Connexion cnx, rstClasse

strClasse = "Select NomClasse from Classe"

ExecReq strClasse, cnx, rstClasse, adOpenKeyset, adLockOptimistic, adCmdText
 Set dcmbClasse1.DataSource = rstClasse
 Set dcmbClasse1.RowSource = rstClasse
 dcmbClasse1.BoundColumn = "NomClasse"
 dcmbClasse1.ListField = "NomClasse"
  Set dcmbClasse2.DataSource = rstClasse
 Set dcmbClasse2.RowSource = rstClasse
 dcmbClasse2.BoundColumn = "NomClasse"
 dcmbClasse2.ListField = "NomClasse"

End Sub

Private Sub Form_Terminate()
cnx.Close
End Sub
