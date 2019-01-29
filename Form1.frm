VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   1230
      Left            =   2760
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   3600
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List4 
      Height          =   1230
      Left            =   4440
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List5 
      Height          =   1230
      Left            =   5400
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List6 
      Height          =   1230
      Left            =   6360
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List7 
      Height          =   1230
      Left            =   7320
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List8 
      Height          =   1230
      Left            =   8280
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List9 
      Height          =   1230
      Left            =   9120
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List10 
      Height          =   1230
      Left            =   9960
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List11 
      Height          =   1230
      Left            =   10920
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   1800
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1695
      Left            =   9000
      TabIndex        =   10
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1455
      Left            =   9240
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   2055
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdajout 
      Caption         =   "Ajout "
      Height          =   855
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdquitter 
      Caption         =   "QUITTER"
      Height          =   675
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtpass 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtlogin 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Officier de l'etat civil"
      Height          =   495
      Left            =   10920
      TabIndex        =   32
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Nom de l'annonceur"
      Height          =   375
      Left            =   9960
      TabIndex        =   31
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Date de l'annonce "
      Height          =   375
      Left            =   7320
      TabIndex        =   30
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Nom du mère"
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Nom du pére"
      Height          =   375
      Left            =   8280
      TabIndex        =   28
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Etat Civil"
      Height          =   375
      Left            =   5400
      TabIndex        =   27
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Lieu de naissance"
      Height          =   375
      Left            =   4440
      TabIndex        =   26
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Date de naissance"
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Sexe"
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Prénom"
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "Nom"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lblpassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lbllogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarations
Dim i As Integer


Private Sub cmdajout_Click()
Dim user As logger
Dim n As Integer
Open "pass.txt" For Random As #2 Len = Len(user)
n = LOF(2) / Len(user)
'Seek #2, n + 1
user.login = txtlogin.Text
user.pass = txtpass.Text
Put #2, n + 1, user
'vbCrLf &

Close #2
End Sub

Private Sub cmdlogin_Click()
Dim user As logger
Open "login.txt" For Random As #1 Len = Len(user)
'Open "C:\Documents and Settings\KEN\Bureau\PROJET VB6\pass.txt" For Random As #2 Len = Len(user)
Get #1, , user

MsgBox user.login & " " & user.pass

i = i + 1
Close #1
End Sub

Private Sub Command1_Click()
'Frmlogin1.Show
frmLogin.Hide
naissacte.Show
End Sub

Private Sub Command2_Click()
Frmlogin1.Show
frmLogin.Hide
End Sub

Private Sub Command3_Click()
Frmchoix.Show
Me.Hide
End Sub

Private Sub Command4_Click()
frmmodifypass.Show
Me.Hide
End Sub
