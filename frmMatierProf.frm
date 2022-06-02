VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMatierProf 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attribution des Professeurs au Matieres"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   Icon            =   "frmMatierProf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMatierProf.frx":0442
   ScaleHeight     =   4800
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
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
      Left            =   3720
      TabIndex        =   9
      Top             =   720
      Width           =   855
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
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin MSDataListLib.DataCombo dcmbProf 
      Height          =   390
      Left            =   2880
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo dcmbClasse 
      Height          =   390
      Left            =   2880
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcmbMatiere 
      Height          =   390
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   6600
      TabIndex        =   12
      Top             =   2160
      Width           =   2535
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
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   720
      Width           =   1815
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
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Matiere"
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
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Matiere 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblProf 
      BackStyle       =   0  'Transparent
      Caption         =   "Professeur"
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
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frmMatierProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnx As New ADODB.Connection

Private Sub cmdOK_Click()
Dim i, annee As Integer
For i = 0 To 11
 If txtAnDebut.Text = 2008 + i And txtAnFin.Text = 2008 + (i + 1) Then
   annee = i + 1
 End If
Next i


Dim str As String
Dim rstDisp As New ADODB.Recordset
str = "Insert into Dispense (IDMatiere,IDProf,Annee) Values (" & dcmbMatiere.BoundText & ","
str = str + dcmbProf.BoundText & "," & annee & " )"
MsgBox str
ExecReq str, cnx, rstDisp, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

Private Sub dcmbClasse_Change()
Dim rst As New ADODB.Recordset
Dim str As String
Connexion cnx, rst
str = "Select * from Matiere Where Classe='" & dcmbClasse.BoundText & "'"
ExecReq str, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText

Set dcmbMatiere.DataSource = rst
Set dcmbMatiere.RowSource = rst
dcmbMatiere.BoundColumn = "NumeroMatiere"
dcmbMatiere.ListField = "Libellé"
End Sub

Private Sub dcmbClasse_Click(Area As Integer)
dcmbMatiere.Text = ""
End Sub

Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim cnx As New ADODB.Connection
Dim rstProf As New ADODB.Recordset

Dim strProf As String
'CenterForm frmAjoutEleve

Taille frmMatierProf, 5000, 8000
Connexion cnx, rst

strProf = "Select * from Professeur"

ExecReq strProf, cnx, rstProf, adOpenKeyset, adLockOptimistic, adCmdText

Set dcmbProf.DataSource = rstProf
Set dcmbProf.RowSource = rstProf
dcmbProf.BoundColumn = "Matricule"
dcmbProf.ListField = "NomProf"



strProf = "Select * from Classe"
ExecReq strProf, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
Set dcmbClasse.DataSource = rst
Set dcmbClasse.RowSource = rst
dcmbClasse.BoundColumn = "NomClasse"
dcmbClasse.ListField = "NomClasse"

End Sub

Private Sub txtMatiere_Change()

End Sub

Private Sub Form_Terminate()
cnx.Close
End Sub

Private Sub txtAnDebut_Change()
' Controle txtAnDebut
End Sub
