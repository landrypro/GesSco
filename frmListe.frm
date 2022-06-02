VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListe 
   BackColor       =   &H80000013&
   Caption         =   "Formulaire des Listing des Eleves"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdListeEleve 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo dcmbClasse 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnx As New ADODB.Connection
Dim rstEleve As New ADODB.Recordset
Dim rst As New ADODB.Recordset

Private Sub cmdListeEleve_Click()
'If deGestBulletin.rscmdEleve.State = adStateOpen Then
      '  deGestBulletin.rscmdEleve.Close
 'End If

 'deGestBulletin.cmdEleve dcmbClasse.Text
' rptEleve.Show
End Sub

Private Sub Form_Load()
  Connexion cnx, rstEleve
 strClasse = "Select NomClasse from Classe"
 ExecReq strClasse, cnx, rst, adOpenKeyset, adLockOptimistic, adCmdText
 If rst.RecordCount = 0 Then
 MsgBox "Vous n'avez pas encore enregistrer des Classes", vbCritical, "Ges_Sco"
 End If
 If rst.RecordCount <> 0 Then
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
