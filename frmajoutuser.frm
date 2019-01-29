VERSION 5.00
Begin VB.Form frmajoutuser 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   13050
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
      Left            =   6720
      MousePointer    =   4  'Icon
      Picture         =   "frmajoutuser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ajout d'un nouvel utilisateur"
      Height          =   2775
      Left            =   4080
      TabIndex        =   1
      Top             =   1800
      Width           =   5295
      Begin VB.TextBox txtlogin 
         Height          =   765
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtpass 
         Height          =   525
         Left            =   2640
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblpass 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOT DE PASSE"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lbluser 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UTILISATEUR A AJOUTER"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdajouter 
      Caption         =   "AJOUTER "
      Height          =   975
      Left            =   4680
      Picture         =   "frmajoutuser.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lbltitre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AJOUT D'UN NOUVEL UTILISATEUR"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   5220
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6015
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
      Left            =   5640
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmajoutuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdajouter_Click()
If txtlogin.Text = "" Then 'Or
MsgBox "Veuillez saisir le nom d'utilisateur"
ElseIf txtpass.Text = "" Then
MsgBox "Veuillez saisir le mot de passe"
Else
Dim user As logger
Dim n As Integer
Open "users.txt" For Random As #2 Len = Len(user)
n = LOF(2) / Len(user)
user.login = txtlogin.Text
user.pass = txtpass.Text
If exist_nom(user) <> -1 Then
MsgBox "Cet utilisateur existe deja dans la base"
Else
Put #2, n + 1, user
End If
Close #2
End If
End Sub

Private Sub cmdretourmnu_Click()
Frmchoix.Show
Me.Hide
End Sub
