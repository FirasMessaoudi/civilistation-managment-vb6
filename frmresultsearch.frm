VERSION 5.00
Begin VB.Form frmresultsearch 
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdrevenir 
      Caption         =   "Revenir au menu"
      Height          =   1815
      Left            =   2880
      TabIndex        =   23
      Top             =   5280
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1680
      TabIndex        =   11
      Top             =   2160
      Width           =   960
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   2640
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   3600
      TabIndex        =   9
      Top             =   2160
      Width           =   1185
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   4800
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.ListBox List6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   5640
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   6600
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   7560
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   8520
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List10 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   9480
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List11 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   10440
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label Label11 
      Caption         =   "Officier de l'etat civil"
      Height          =   495
      Left            =   10440
      TabIndex        =   22
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Nom de l'annonceur"
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Date de l'annonce "
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Nom du mère"
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Nom du pére"
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Etat Civil"
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Lieu de naissance"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Date de naissance"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Sexe"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Prénom"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nom"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblresultat 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RESULTAT DE RECHERCHE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmresultsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdrevenir_Click()
Frmchoix.Show
Me.Hide
End Sub
