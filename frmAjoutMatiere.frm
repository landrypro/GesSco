VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAjoutMatiere 
   BackColor       =   &H80000009&
   Caption         =   "Ajout Des Matières"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmAjoutMatiere.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAjoutMatiere.frx":0442
   ScaleHeight     =   10245
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdModif 
      BackColor       =   &H80000009&
      Caption         =   "Modifier"
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox txtCodeMatiere 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cbxLibelle 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   14
      Top             =   3600
      Width           =   2055
   End
   Begin VB.ComboBox cbxGroupe 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtAnFin 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   6480
      Width           =   855
   End
   Begin MSDataListLib.DataCombo dcmbClasse 
      Height          =   360
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
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
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "Valider"
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H80000009&
      Caption         =   "Annuler"
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtAnDebut 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txtCoefficient 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4680
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dgClasse 
      Height          =   5775
      Left            =   6360
      TabIndex        =   18
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   15600
      TabIndex        =   19
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   70
      X1              =   0
      X2              =   19080
      Y1              =   10560
      Y2              =   10560
   End
   Begin VB.Label lblTitre 
      BackStyle       =   0  'Transparent
      Caption         =   "Liste des Matieres de  "
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
      Left            =   6480
      TabIndex        =   16
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      Caption         =   "/"
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
      Left            =   3960
      TabIndex        =   12
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Coefficient"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Année Scolaire"
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
      TabIndex        =   6
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Libellé"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Groupe"
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
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
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
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "CodeMatiere"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmAjoutMatiere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMatiere As New ADODB.Recordset
Dim rstClasse  As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim rstAffich As New ADODB.Recordset

Dim cnx As ADODB.Connection

Dim strInsert, strValue, strMatiere, strClasse  As String

Private Sub cmdOK_Click()
Dim i, annee As Integer
Dim sqlmat As String
For i = 0 To 11
 If txtAnDebut.Text = 2008 + i And txtAnFin.Text = 2008 + (i + 1) Then
   annee = i + 1
 End If
Next i
strInsert = " Insert Into Matiere (Libellé,Coefficient,Classe, "
strInsert = strInsert + " IDGroupe,Année )"
strValue = " Values ('" + cbxLibelle.Text + "'," + txtCoefficient.Text + ","
strValue = strValue + "'" + dcmbClasse.BoundText + "'," + cbxGroupe.Text + "," + CStr(annee) + "  )"
strMatiere = strInsert + strValue

ExecReq strMatiere, cnx, rstMatiere, adOpenKeyset, adLockOptimistic, adCmdText
'rstMatiere.Open strMatiere, cnx
dcmbClasse_Change
End Sub


Private Sub dcmbClasse_Change()
Dim str As String
lblTitre.Caption = "Liste des Matieres de " & dcmbClasse.BoundText
str = " Select NumeroMatiere as Numero,Libellé, "
str = str + " Coefficient,IDGroupe as Groupe From Matiere Where Classe= '" & dcmbClasse.Text & "'"
str = str + "ORDER BY Libellé "
ExecReq str, cnx, rstAffich, adOpenKeyset, adLockOptimistic, adCmdText
'rstAffich.Open str, cnx
Set dgClasse.DataSource = rstAffich
End Sub



Private Sub Form_Load()
Taille frmAjoutMatiere, 9000, 10000

'Connexion cnx, rstMatiere
Connexion cnx, rst
'Connexion cnx, rstAffich
'Connexion cnx, rstClasse

strClasse = "Select NomClasse from Classe"
ExecReq strClasse, cnx, rstClasse, adOpenKeyset, adLockOptimistic, adCmdText
'rstClasse.Open strClasse, cnx
If rstClasse.RecordCount = 0 Then
 MsgBox "Vous n'avez pas encore enregistrer des Classes", vbCritical, "Ges_Sco"
End If
Set dcmbClasse.DataSource = rstClasse
Set dcmbClasse.RowSource = rstClasse

dcmbClasse.BoundColumn = "NomClasse"
dcmbClasse.ListField = "NomClasse"

cbxGroupe.AddItem ("1")
cbxGroupe.AddItem ("2")
cbxGroupe.AddItem ("3")
For i = 0 To 17
cbxLibelle.AddItem (frmNote.lblMatiere(i).Caption)
Next i
frmNote.Hide
Unload frmNote

End Sub

Private Sub Form_Terminate()

sqlmat = " INSERT INTO [Note] SELECT Eleve.matricule AS matricule, Matiere.NumeroMatiere AS iDMatiere, Sequence.NoSequence AS Sequence, AnneeScolaire.NoAnne AS NoAnnee"
sqlmat = sqlmat + " From Eleve, Matiere, Sequence, AnneeScolaire"
sqlmat = sqlmat + "  WHERE (Eleve.Classe=Matiere.Classe) AND Matiere.NumeroMatiere  not in (Select Distinct IDMatiere FROM [Note]) "
ExecReq sqlmat, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
 'rst.Open sqlmat, cnx

cnx.Close
End Sub
