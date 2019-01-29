VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form naissacte 
   BackColor       =   &H80000016&
   Caption         =   "Ajout d'une acte de naissance"
   ClientHeight    =   9690
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14430
   FillColor       =   &H80000003&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdretourmnu 
      BackColor       =   &H80000000&
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
      Height          =   1215
      Left            =   7440
      MousePointer    =   4  'Icon
      Picture         =   "naissacte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8040
      Width           =   2295
   End
   Begin VB.ComboBox combocite 
      Height          =   315
      ItemData        =   "naissacte.frx":06EA
      Left            =   7440
      List            =   "naissacte.frx":0A39
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox cbetatciv 
      Height          =   315
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3960
      Width           =   2595
   End
   Begin VB.CommandButton cmdannuler 
      BackColor       =   &H80000000&
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10080
      MousePointer    =   4  'Icon
      Picture         =   "naissacte.frx":166E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdajouter 
      BackColor       =   &H80000000&
      Caption         =   "Ajouter"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      MousePointer    =   4  'Icon
      Picture         =   "naissacte.frx":1D58
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Frame framsex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7440
      TabIndex        =   18
      Top             =   2520
      Width           =   2535
      Begin VB.OptionButton optmale 
         BackColor       =   &H80000016&
         Caption         =   "Masculin"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MouseIcon       =   "naissacte.frx":2442
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optfemme 
         BackColor       =   &H80000016&
         Caption         =   "Féminin"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MouseIcon       =   "naissacte.frx":330C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtofficier 
      Height          =   735
      Left            =   7440
      TabIndex        =   16
      Top             =   7080
      Width           =   2535
   End
   Begin VB.TextBox txtannonc 
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox txtmere 
      Height          =   375
      Left            =   7440
      TabIndex        =   14
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox txtpere 
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtprenom 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtnom 
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   960
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker Datenaiss 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   151846913
      CurrentDate     =   42479
   End
   Begin MSComCtl2.DTPicker Dateannonce 
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   5880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   151846913
      CurrentDate     =   42479
   End
   Begin VB.Label etatjur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Etat civil"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5295
      TabIndex        =   24
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label lbltitre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ajout d'un nouvel acte de naissance"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3840
      TabIndex        =   21
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date de naissance"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4920
      TabIndex        =   17
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblagent 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "L'officier d'état civil "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de l'annonceur"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date d'annonce"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblnommere 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du mère"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblnompere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du père"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5640
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbllieunaiss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Lieu de naissance"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblprenom 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Prénom"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblNom 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000013&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   2295
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu ajoutact 
         Caption         =   "Ajouter un nouvel acte de naissance"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "naissacte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
Me.Hide
frmabout.Show
End Sub



Private Sub cmdajouter_Click()
Dim nv As person
Dim n As Integer
Dim m As Integer
m = FreeFile
Open "Persons1.txt" For Random As #4 Len = Len(nv)
n = LOF(4) / Len(nv)
Seek #4, n + 1

nv.nom = txtnom.Text 'vbCrLf &
nv.prenom = txtprenom.Text
'nv.Lieunaiss = txtlieunaiss.Text
nv.Lieunaiss = combocite.Text
If optmale.Value = True Then
nv.sexe = "Masculin"
Else
nv.sexe = "Féminin"
End If
nv.Datenaiss = Datenaiss.Value
nv.etatciv = cbetatciv.Text '"Célibataire" '
nv.nompere = txtpere.Text
nv.nommere = txtmere.Text
nv.Dateannonce = Dateannonce.Value
nv.nomannonce = txtannonc.Text
nv.nomofficier = txtofficier.Text
Put #4, , nv

Close #4
End Sub

Private Sub cmdannuler_Click()
txtnom.Text = ""
txtprenom.Text = ""
txtpere.Text = ""
txtmere.Text = ""
txtannonc.Text = ""
txtofficier.Text = ""
End Sub




Private Sub cmdretourmnu_Click()
Frmchoix.Show
Me.Hide
End Sub

Private Sub Dateannonce_Change()
'If Dateannonce.Day < Datenaiss.Day Or Dateannonce.Month < Datenaiss.Month Or Dateannonce.Year < Datenaiss.Year Then
If Dateannonce.Value < Datenaiss.Value Then
MsgBox "La date doit etre supérieure ou égale à la date de naissance"
Dateannonce.Value = Datenaiss.Value
End If
End Sub



Private Sub Form_Load()
'cbetatciv.Enabled = False
'cbetatciv.List(0) = "Célibataire"

Dim etat1 As String
Dim etat2 As String
Dim etat3 As String
Dim etat4 As String
Dim etat() As String

txtnom.Text = ""
txtprenom.Text = ""
txtpere.Text = ""
txtmere.Text = ""
txtannonc.Text = ""
txtofficier.Text = ""



End Sub

Private Sub optfemme_Click()
cbetatciv.List(0) = "Célibataire"
cbetatciv.List(1) = "Mariée"
cbetatciv.List(2) = "Veuve"
cbetatciv.List(3) = "Divorcée"

End Sub

Private Sub optmale_Click()

cbetatciv.List(0) = "Célibataire"
cbetatciv.List(1) = "Marié"
cbetatciv.List(2) = "Veuf"
cbetatciv.List(3) = "Divorcé"
End Sub
