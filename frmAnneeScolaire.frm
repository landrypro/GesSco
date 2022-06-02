VERSION 5.00
Begin VB.Form frmAnneeScolaire 
   BackColor       =   &H80000009&
   Caption         =   "Ajouter des Années Scolaires"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   Icon            =   "frmAnneeScolaire.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAnneeScolaire.frx":0442
   ScaleHeight     =   3945
   ScaleWidth      =   6300
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
      Left            =   960
      Picture         =   "frmAnneeScolaire.frx":2F73
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
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
      Left            =   3360
      Picture         =   "frmAnneeScolaire.frx":631F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtFin 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtDebut 
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
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Année de Fin"
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
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Année de Debut"
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
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmAnneeScolaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Dim rst As New ADODB.Recordset
Dim cnx As New ADODB.Connection
Dim sql As String
sql = "Insert into AnneeScolaire (Debut,Fin)"
sql = sql + "Value ( " + txtDebut + "," + txtFin + ")"
Connexion cnx, rst
rst.Open sql, cnx
ExecReq sql, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

