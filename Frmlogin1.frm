VERSION 5.00
Begin VB.Form Frmlogin1 
   Caption         =   "LOGIN"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitter 
      Cancel          =   -1  'True
      Caption         =   "Quitter "
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
      Left            =   3120
      MousePointer    =   4  'Icon
      Picture         =   "Frmlogin1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdloginuser 
      Caption         =   "Se connecter que tant d'utilisateur"
      Default         =   -1  'True
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
      Left            =   5640
      MousePointer    =   4  'Icon
      Picture         =   "Frmlogin1.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtlogin 
      Height          =   525
      Left            =   6840
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtpass 
      Height          =   525
      Left            =   6840
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdloginadmin 
      Caption         =   "Se connecter que tant d'Administrateur"
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
      Left            =   8160
      MousePointer    =   4  'Icon
      Picture         =   "Frmlogin1.frx":0DD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GESTION DE L'ETAT CIVIL"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label lbllogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USERNAME"
      DragIcon        =   "Frmlogin1.frx":14BE
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   540
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   1965
   End
   Begin VB.Label lblpassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   540
      Left            =   4560
      TabIndex        =   3
      Top             =   2760
      Width           =   1965
   End
End
Attribute VB_Name = "Frmlogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Sub cmdloginadmin_Click()
Dim user As logger
user.login = txtlogin.Text
user.pass = txtpass.Text

If txtlogin.Text = "" Then
MsgBox "VEUILLEZ SAISIR Votre nom d'utilisateur"
ElseIf txtpass.Text = "" Then
MsgBox "VEUILLEZ SAISIR Votre mot de passe"
ElseIf (exist(user)) <> -1 Then

Frmchoix.Show
Me.Hide
Else: MsgBox "THE USER DOES NOT EXIST IN THE DATABASE"

End If

End Sub

Private Sub cmdloginuser_Click()
Dim user1 As logger
user1.login = txtlogin.Text
user1.pass = txtpass.Text

If txtlogin.Text = "" Then
MsgBox "VEUILLEZ SAISIR Votre nom d'utilisateur"
ElseIf txtpass.Text = "" Then
MsgBox "VEUILLEZ SAISIR Votre mot de passe"
ElseIf (exist2(user1)) <> -1 Then

Frmchoix.Show
Frmlogin1.Hide
Frmchoix.mnuadmin.Enabled = False
Else: MsgBox "THE USER DOES NOT EXIST IN THE DATABASE"

txtlogin.Text = ""
txtpass.Text = ""
txtlogin.SetFocus
End If

End Sub

Private Sub cmdquitter_Click()
Dim Response As Integer
Response = MsgBox("Etes vous sur que vous voulez quitter l'application?", vbYesNo, "Exit")
If Response = vbYes Then
End
Else
Me.Show
End If
End Sub

    Private Sub Form_Load()
    
    'MsgBox "Bienvenue a l'application de gestion d'état civil.Veuillez se connecter pour utiliser l'application"
    
    End Sub


