VERSION 5.00
Begin VB.Form Frmetat 
   Caption         =   "Modification d'un Couple"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmenu 
      Caption         =   "&REVENIR AU MENU "
      Height          =   1215
      Left            =   8040
      Picture         =   "Formetat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "VEUILLEZ CHOISIR LES DEUX PARTIES:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   13935
      Begin VB.ListBox List18 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   10200
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.ListBox List17 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   9120
         TabIndex        =   13
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ListBox List16 
         Height          =   1230
         Left            =   6120
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ListBox List15 
         Height          =   1230
         Left            =   8160
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List14 
         Height          =   1230
         Left            =   7200
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List13 
         Height          =   1230
         Left            =   5160
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List12 
         Height          =   1230
         Left            =   4200
         TabIndex        =   8
         Top             =   1560
         Width           =   960
      End
      Begin VB.ComboBox cbetat 
         Height          =   315
         ItemData        =   "Formetat.frx":06EA
         Left            =   4440
         List            =   "Formetat.frx":06F1
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4200
         Width           =   6375
      End
      Begin VB.Label Label5 
         Caption         =   "Date de mariage"
         Height          =   495
         Left            =   10200
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Prénom de l'épouse"
         Height          =   375
         Left            =   8160
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Nom de l'épouse"
         Height          =   375
         Left            =   7200
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Date de naissance de l'epouse"
         Height          =   735
         Left            =   9120
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Date de naissance du mari"
         Height          =   735
         Left            =   6120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Prénom du mari"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Nom      du mari"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VEUILLEZ CHOISIR LES DEUX PARTIES"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Veuillez choisir le nouvel état"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   5
         Top             =   3960
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&QUITTER"
      Height          =   1215
      Left            =   11280
      TabIndex        =   2
      Top             =   7560
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&MODIFIER L'ETAT"
      Height          =   1215
      Left            =   4800
      TabIndex        =   1
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MODIFICATION D'UN ETAT CIVIL "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "Frmetat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Remplissage_liste()
'Dim t As person
'Dim k As person
'Dim i As Integer
'i = 1
'Dim n As Integer
'n = FreeFile
'Open "couple.txt" For Random As #n Len = Len(t)
'While Not EOF(n)
'Get n, i, t
'List1.AddItem (t.nom)
'List2.AddItem (t.prenom)
'List3.AddItem (t.sexe)
'List4.AddItem (t.Lieunaiss)
'List5.AddItem (t.etatciv)
'List6.AddItem (t.Datenaiss)
'List7.AddItem (t.Dateannonce)
'List8.AddItem (t.nompere)
'List9.AddItem (t.nommere)
'List10.AddItem (t.nomannonce)
'List11.AddItem (t.nomofficier)
''List1(i).AddItem (t.prenom)
'i = i + 1
'Wend
''End If
'Close #n
Dim cp As couple

End Sub

Private Sub Form_Load()
Call Remplissage_liste
cbetat.List(0) = "Célibataire"
cbetat.List(1) = "Marié"
cbetat.List(2) = "Veuf"
cbetat.List(3) = "Divorcé"
End Sub
