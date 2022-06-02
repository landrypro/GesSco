VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAjoutProf 
   BackColor       =   &H80000009&
   Caption         =   "Ajoute de Professeur"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   Icon            =   "frmAjoutProf.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   9060
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dgProf 
      Height          =   2655
      Left            =   4200
      TabIndex        =   9
      Top             =   5760
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.CommandButton cmdAttrib 
      BackColor       =   &H80000005&
      Caption         =   "Attribuer des Matieres au Prof"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000005&
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtMatricule 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6480
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtPrenom 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6480
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdValider 
      BackColor       =   &H80000005&
      Caption         =   "Ajouter"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtNom 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6480
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Prenom "
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
      Left            =   4440
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Nom "
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
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmAjoutProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstProf As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim cnx As New ADODB.Connection
Dim strProf As String

Private Sub cmdAttrib_Click()
frmMatierProf.Show , 1
End Sub

Private Sub cmdValider_Click()
strProf = "Insert into Professeur (Matricule,NomProf,PrenomProf)"
strProf = strProf + " Values ('" + txtMatricule + "','" + txtNom + "','" + txtPrenom + "' )"
If ExecReq(strProf, cnx, rstProf, adOpenKeyset, adLockOptimistic, adCmdText) = True Then
 MsgBox "Ajout réussi"
End If
cnx.Close
ExecReq "Select IDProf as Numeror ,Matricule,NomProf,PrenomProf FROM Professeur", cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
Set dgProf.DataSource = rst
End Sub

Private Sub Form_Load()

Connexion cnx, rstProf
Connexion cnx, rst
ExecReq "Select IDProf as Numeror ,Matricule,NomProf,PrenomProf FROM Professeur", cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
Set dgProf.DataSource = rst
End Sub

Private Sub Form_Terminate()
cnx.Close
End Sub
