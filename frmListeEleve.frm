VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListeEleve 
   Caption         =   "Liste des Eleves"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   7170
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdListe 
      Caption         =   "Liste"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo dcmbClasse 
      Height          =   390
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmListeEleve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim cnx As New ADODB.Connection
Dim rstProf As New ADODB.Recordset

Taille frmMatierProf, 5000, 8000
Connexion cnx, rst

strProf = "Select * from Classe"
ExecReq strProf, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
Set dcmbClasse.DataSource = rst
Set dcmbClasse.RowSource = rst
dcmbClasse.BoundColumn = "IDClasse"
dcmbClasse.ListField = "NomClasse"
End Sub
