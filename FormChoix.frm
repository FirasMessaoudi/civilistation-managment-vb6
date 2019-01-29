VERSION 5.00
Begin VB.Form Frmchoix 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmodifycouple 
      Caption         =   "MODIFICATION D'UN COUPLE "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   4
      Top             =   6360
      Width           =   4335
   End
   Begin VB.CommandButton cmdaddcouple 
      Caption         =   "AJOUTER UN  NOUVEAU COUPLE "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   3
      Top             =   4800
      Width           =   4335
   End
   Begin VB.CommandButton cmdsearchperson 
      Caption         =   "RECHERCHE D'UNE PERSONNE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
      Width           =   4335
   End
   Begin VB.CommandButton cmdactnaissadd 
      Caption         =   "AJOUTER UN NOUVEL ACTE DE NAISSANCE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4680
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VEUILLEZ CHOISIR UNE OPERATION A EFFECTUER"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   11325
   End
   Begin VB.Menu mnumenu 
      Caption         =   "Menu"
      Begin VB.Menu mnulogout 
         Caption         =   "Se déconnecter"
      End
      Begin VB.Menu mnuquitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "Administration"
      Begin VB.Menu mnuadduser 
         Caption         =   "Ajout D'un nouvel utilisateur"
      End
      Begin VB.Menu mnudeleteuser 
         Caption         =   "Suppression D'un utilisateur"
      End
      Begin VB.Menu mnumodifyuser 
         Caption         =   "Modifier un mot de passe d'un utilisateur"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Frmchoix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
End Sub

Private Sub Menuquitter_Click()
End
End Sub

Private Sub cmdactnaissadd_Click()
naissacte.Show
Me.Hide
End Sub

Private Sub cmdaddcouple_Click()
mnumenu.Enabled = False
frmajoutcouple.Show
Me.Hide
End Sub

Private Sub cmdmodifycouple_Click()
Frmetat.Show
Me.Hide
End Sub

Private Sub cmdsearchperson_Click()
frmrecherche.Show
Me.Hide
End Sub

Private Sub mnuabout_Click()
frmabout.Show
End Sub

Private Sub mnuadduser_Click()
frmajoutuser.Show
Me.Hide
End Sub

Private Sub mnulogout_Click()
Frmlogin1.Show
Me.Hide
End Sub

Private Sub mnumodifyuser_Click()
frmmodifypass.Show
Me.Hide
End Sub

Private Sub mnuquitter_Click()
End
End Sub
