VERSION 5.00
Begin VB.Form frmajoutcouple 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   15495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdretourmnu 
      Caption         =   "Retour au Menu Principal"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      MousePointer    =   4  'Icon
      Picture         =   "frmajoutcouple.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdajoutercpl 
      Caption         =   "&AJOUTER LE NOUVEAU COUPLE"
      Height          =   975
      Left            =   3960
      TabIndex        =   38
      Top             =   7440
      Width           =   3255
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
      Height          =   5655
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   13935
      Begin VB.ComboBox cbetat 
         Height          =   315
         ItemData        =   "frmajoutcouple.frx":06EA
         Left            =   3360
         List            =   "frmajoutcouple.frx":06F1
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4560
         Width           =   6375
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Left            =   3840
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List3 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   4680
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List4 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   6480
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List5 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   7440
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List6 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   5520
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List7 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   8400
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List8 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   9360
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List9 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   10200
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List10 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   11160
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox List11 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   12120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   2880
         TabIndex        =   12
         Top             =   720
         Width           =   960
      End
      Begin VB.ListBox List12 
         Height          =   1230
         Left            =   2880
         TabIndex        =   11
         Top             =   2520
         Width           =   960
      End
      Begin VB.ListBox List13 
         Height          =   1230
         Left            =   3840
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin VB.ListBox List14 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   4800
         TabIndex        =   9
         Top             =   2520
         Width           =   735
      End
      Begin VB.ListBox List15 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   6480
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox List16 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   7560
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin VB.ListBox List17 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   5520
         TabIndex        =   6
         Top             =   2520
         Width           =   975
      End
      Begin VB.ListBox List18 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   8400
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.ListBox List19 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   9360
         TabIndex        =   4
         Top             =   2520
         Width           =   975
      End
      Begin VB.ListBox List20 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   10320
         TabIndex        =   3
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox List21 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   11400
         TabIndex        =   2
         Top             =   2520
         Width           =   855
      End
      Begin VB.ListBox List22 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   12240
         TabIndex        =   1
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Veuillez choisir l'opération"
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
         Left            =   240
         TabIndex        =   36
         Top             =   4560
         Width           =   2535
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
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Officier de l'etat civil"
         Height          =   495
         Left            =   12120
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Nom de l'annonceur"
         Height          =   375
         Left            =   11160
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Date de l'annonce "
         Height          =   375
         Left            =   8400
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Nom du mère"
         Height          =   375
         Left            =   10320
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Nom du pére"
         Height          =   375
         Left            =   9360
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Etat Civil"
         Height          =   375
         Left            =   7560
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Lieu de naissance"
         Height          =   375
         Left            =   6600
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Date de naissance"
         Height          =   375
         Left            =   5640
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Sexe"
         Height          =   375
         Left            =   4800
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Prénom"
         Height          =   375
         Left            =   3840
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Nom"
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AJOUT D'UN NOUVEAU COUPLE"
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
      Left            =   3840
      TabIndex        =   37
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmajoutcouple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub remplissage_liste_nvcouple()
Dim t As person
Dim n As Integer
Dim i As Integer
Dim m As String * 8
Dim f As String * 7
Dim male As String * 8
male = "Masculin"
Dim celib As String * 11
celib = "Célibataire"
Dim ch3 As String * 11
Dim ch4 As String * 7
Dim ch5 As String * 8
Dim ch6 As String * 5
Dim ch7 As String * 4
Dim female As String * 7
female = "Féminin"
Dim divorm As String * 7
Dim divorf As String * 8
Dim vfm As String * 4
Dim vff As String * 5
divorm = "Divorcé"
divorf = "Divorcée"
vfm = "veuf"
vff = "veuve"
n = FreeFile
i = 1
Open "Persons1.txt" For Random As #n Len = Len(t)
While Not EOF(n)
Get n, i, t
m = t.sexe
f = t.sexe
ch3 = t.etatciv
ch4 = t.etatciv
ch5 = t.etatciv
ch6 = t.etatciv

If StrComp(male, m) = 0 And StrComp(t.etatciv, celib) = 0 Then
'And StrComp(ch3, "Célibataire") = 0 Then 'Or StrComp(ch4, divorm) = 0 Or StrComp(ch7, vfm) = 0) Then 'And (t.etatciv = "Célibataire" Or t.etatciv = "Veuf" Or t.etatciv = "Divorcé") Then
List1.AddItem (t.nom)
List2.AddItem (t.prenom)
List3.AddItem (t.sexe)
List4.AddItem (t.Lieunaiss)
List5.AddItem (t.etatciv)
List6.AddItem (t.Datenaiss)
List7.AddItem (t.Dateannonce)
List8.AddItem (t.nompere)
List9.AddItem (t.nommere)
List10.AddItem (t.nomannonce)
List11.AddItem (t.nomofficier)
ElseIf StrComp(f, female) = 0 And (StrComp(t.etatciv, celib) = 0 Or StrComp(ch6, "Veuve") = 0 Or ch5 = "Divorcée") Then
List12.AddItem (t.nom)
List13.AddItem (t.prenom)
List14.AddItem (t.sexe)
List15.AddItem (t.Lieunaiss)
List16.AddItem (t.etatciv)
List17.AddItem (t.Datenaiss)
List18.AddItem (t.Dateannonce)
List19.AddItem (t.nompere)
List20.AddItem (t.nommere)
List21.AddItem (t.nomannonce)
List22.AddItem (t.nomofficier)
End If
i = i + 1
Wend
Close #n
End Sub

Private Sub cmdajoutercpl_Click()
Dim cp As couple
Dim m As Integer
Dim n, i As Integer
m = FreeFile
i = 1
If List1.ListIndex < 0 Or List2.ListIndex < 0 Or List12.ListIndex < 0 Or List13.ListIndex < 0 Then
MsgBox "Veuillez selectionner les 2 parties"
ElseIf cbetat.ListIndex < 0 Then
MsgBox " Veuillez selectionner l'opération "
End If
cpm% = List1.ListIndex
cpf% = List12.ListIndex
Open "Couple.txt" For Random As #m Len = Len(cp)
n = LOF(m) / Len(cp)
cp.nommari = List1.List(cpm)
cp.nomepouse = List12.List(cpf)
cp.prenommari = List2.List(cpm)
cp.prenomepouse = List13.List(cpf)
cp.Date_marriage = DateValue(Now)
cp.Datenaissmari = List6.List(cpm)
cp.datenaissepouse = List17.List(cpf)
Put #m, n + 1, cp
Close #m
End Sub

Private Sub Form_Load()
Call remplissage_liste_nvcouple
Dim i As Integer
'For i = 0 To List1.ListCount

End Sub
