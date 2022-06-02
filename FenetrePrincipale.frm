VERSION 5.00
Begin VB.MDIForm FenetrePrincipale 
   BackColor       =   &H80000009&
   Caption         =   "GEST_SCO"
   ClientHeight    =   8475
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11565
   Icon            =   "FenetrePrincipale.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "FenetrePrincipale.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11565
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuAjoutEleve 
         Caption         =   "&Enregistrer un nouveau élève"
      End
      Begin VB.Menu mnuSaisieNote 
         Caption         =   "&Saisie des notes"
      End
      Begin VB.Menu mnuImpressBull 
         Caption         =   "&Impression des bulletins"
         Enabled         =   0   'False
      End
      Begin VB.Menu smnuQuit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuGestElev 
      Caption         =   "Gestion des &Elèves"
      Begin VB.Menu smnuAjoutElev 
         Caption         =   "&Ajouter Un Elève"
      End
      Begin VB.Menu smnMAJElev 
         Caption         =   "&Modifier un Elève"
      End
   End
   Begin VB.Menu mnuGestNote 
      Caption         =   "Gestion des &Notes"
      Begin VB.Menu smnuAjoutNote 
         Caption         =   "&Ajoute des Notes"
         Begin VB.Menu ssmnAjoutClass 
            Caption         =   "Par &Classe"
         End
      End
   End
   Begin VB.Menu mnuParam 
      Caption         =   "&Paramétrer"
      Begin VB.Menu smnuMatiere 
         Caption         =   "&Matieres"
      End
      Begin VB.Menu smnuClass 
         Caption         =   "&Classe"
      End
      Begin VB.Menu smnuProf 
         Caption         =   "&Professeur"
         Begin VB.Menu smnuAjoutProf 
            Caption         =   "Ajout de Prof"
         End
         Begin VB.Menu smnuAttrib 
            Caption         =   "Attribution Prof--Matieres"
         End
      End
      Begin VB.Menu smnuAn 
         Caption         =   "A&nnée"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEtats 
      Caption         =   "&Etats"
      Begin VB.Menu cmdListeEleve 
         Caption         =   "Liste des &Elèves"
      End
      Begin VB.Menu smnuBulletin 
         Caption         =   "Bulletin"
      End
      Begin VB.Menu mnuNote 
         Caption         =   "Releve de Notes"
      End
   End
   Begin VB.Menu mnuAide 
      Caption         =   "&Aide"
      Begin VB.Menu smnuAide 
         Caption         =   "&?"
      End
      Begin VB.Menu smnuApropos 
         Caption         =   "A &Propos"
      End
   End
End
Attribute VB_Name = "FenetrePrincipale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strNoteCoef As String
Dim rstNoteCoef As New ADODB.Recordset
Dim cnx As New ADODB.Connection

Private Sub cmdListeEleve_Click()
frmListe.Show
End Sub

Private Sub MDIForm_Load()



Connexion cnx, rstNoteCoef

'strNoteCoef = " SELECT IDMatiere, Coefficient, Note, NoteCoef From Matiere, [Note]"
'strNoteCoef = strNoteCoef + " Where Matiere.NumeroMatiere = [Note].IDMatiere"
'rstNoteCoef.Open strNoteCoef, cnx, adOpenKeyset, adLockOptimistic, adCmdText
'rstNoteCoef.MoveFirst
'While Not rstNoteCoef.EOF
   ' rstNoteCoef!noteCoef = rstNoteCoef!Note * rstNoteCoef!Coefficient
   ' rstNoteCoef.Update
   ' rstNoteCoef.MoveNext
'Wend

'MsgBox "NB:" & vbCrLf & "-Enregistrer d'abord les classes" & vbCrLf & "-Ensuite les matieres" & vbCrLf & "-Et pour finir les Elèves avant de faire autre chose", vbInformation
End Sub

Private Sub mnuAjoutEleve_Click()
frmAjoutEleve.Show
End Sub

Private Sub mnuNote_Click()
rptRelevNote.Show
End Sub

Private Sub mnuSaisieNote_Click()
frmNote.Show
End Sub

Private Sub smnMAJElev_Click()
frmModifEleve.Show
End Sub

Private Sub smnuAjoutElev_Click()
Unload frmAjoutEleve
frmAjoutEleve.Show
End Sub

Private Sub smnuAjoutProf_Click()
frmAjoutProf.Show
End Sub

Private Sub smnuAn_Click()
frmAnnee.Show
End Sub

Private Sub smnuApropos_Click()
frmApropos.Show
End Sub

Private Sub smnuAttrib_Click()
frmMatierProf.Show
End Sub

Private Sub smnuBulletin_Click()
frmChoixBulletin.Show
End Sub

Private Sub smnuClass_Click()
frmClasse.Show
End Sub

Private Sub smnuMatiere_Click()
frmAjoutMatiere.Show
End Sub



Private Sub smnuQuit_Click()
End
End Sub

Private Sub ssmnAjoutClass_Click()
frmNote.Show
End Sub

