VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClasse 
   BackColor       =   &H80000009&
   Caption         =   "Ajout et Configuration des Classes"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   Icon            =   "frmClasse.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmClasse.frx":0442
   ScaleHeight     =   8940
   ScaleWidth      =   9105
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txtClasse 
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
      Left            =   3000
      TabIndex        =   10
      Text            =   "Selectionner "
      Top             =   1440
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid dgClasse 
      Height          =   6375
      Left            =   6480
      TabIndex        =   9
      Top             =   1560
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
      _Version        =   393216
      Enabled         =   -1  'True
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   9
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
   Begin VB.ComboBox cbCycle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtLibelle 
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
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H8000000E&
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   1440
      TabIndex        =   11
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Exemple :Terminale C,Terminale D,Seconde C, Sixieme M1,Premiere C,Quatrieme Allemande"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Exemple : Tle C,Tle D,2nd C,P C ,6eM1,4e All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle"
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
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Classe"
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
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "frmClasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstClasse As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim cnx As ADODB.Connection
Dim strInsert, strValue, strClass As String

Private Sub cmdListeClasse_Click()

End Sub

Private Sub cmdOK_Click()
strInsert = " Insert Into Classe ( NomClasse,Libellé,IDCycle ) "
strValue = " Values ( '" + txtClasse + "','" + txtLibelle + "'," + cbCycle + " )"
strClasse = strInsert + strValue
ExecReq strClasse, cnx, rstClasse, adOpenKeyset, adLockOptimistic, adCmdText
ExecReq "Select NomClasse,Libellé,IDCycle as Cycle FROM Classe Order BY IDCycle ASC", cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
Set dgClasse.DataSource = rst
End Sub

Private Sub Form_Load()
'CenterForm frmAjoutEleve
Taille frmClasse, 9000, 10000
Connexion cnx, rstClasse
Connexion cnx, rst
ExecReq "Select NomClasse,Libellé,IDCycle as Cycle FROM Classe Order BY IDCycle ASC", cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText

cbCycle.AddItem "1"
cbCycle.AddItem "2"

Set dgClasse.DataSource = rst
txtClasse.AddItem "6e M"
txtClasse.AddItem "6e M1"
txtClasse.AddItem "6e M2"

txtClasse.AddItem "5e M"
txtClasse.AddItem "5e M1"
txtClasse.AddItem "5e M2"

txtClasse.AddItem "4e All"
txtClasse.AddItem "4e All1"
txtClasse.AddItem "4e All2"
txtClasse.AddItem "4e Esp"
txtClasse.AddItem "4e Esp1"
txtClasse.AddItem "4e Esp2"

txtClasse.AddItem "3e M1 All"
txtClasse.AddItem "3e M2 All1"
'txtClasse.AddItem "3e All2"
txtClasse.AddItem "3e M1 Esp"
txtClasse.AddItem "3e M2 Esp"
'txtClasse.AddItem "3e Esp2"

txtClasse.AddItem "2nd C"
txtClasse.AddItem "2nd C1"
txtClasse.AddItem "2nd C2"

txtClasse.AddItem "2nd A4 Esp"
txtClasse.AddItem "2nd A4 Esp1"
txtClasse.AddItem "2nd A4 All"
txtClasse.AddItem "2nd A4 All1"


txtClasse.AddItem "P C"
txtClasse.AddItem "P D"
txtClasse.AddItem "P D1"
txtClasse.AddItem "P A4 All"
txtClasse.AddItem "P A4 Esp"

txtClasse.AddItem "Tle C"
txtClasse.AddItem "Tle D"
txtClasse.AddItem "Tle D1"
txtClasse.AddItem "Tle A4 All"
txtClasse.AddItem "Tle A4 Esp"


End Sub

Private Sub Label9_Click()

End Sub

Private Sub Form_Terminate()
cnx.Close
End Sub
