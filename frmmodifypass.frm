VERSION 5.00
Begin VB.Form frmmodifypass 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   3000
         TabIndex        =   10
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtpassmodif 
         Height          =   855
         Left            =   2880
         TabIndex        =   7
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Veuillez choisir le type de l'utilisateur"
         Height          =   975
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   3615
         Begin VB.OptionButton optadmin 
            Caption         =   "Administrateur"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optuser 
            Caption         =   "Utilisateur normal"
            Height          =   255
            Left            =   1680
            TabIndex        =   5
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Label lblpass 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LE NOUVEAU MOT DE PASSE"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   9
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lbluser 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UTILISATEUR A MODIFIER"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   720
         TabIndex        =   8
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdquitter 
      Caption         =   "&Quitter"
      Height          =   1215
      Left            =   8760
      TabIndex        =   1
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdmodifier 
      Caption         =   "&Modifier"
      Height          =   1215
      Left            =   4560
      TabIndex        =   0
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Veuillez Choisir l'utilisateur pour modifier son mot de passe"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmmodifypass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmodifier_Click()

'Call Remplissage_liste
'Call modify_pass
Call modifypass

End Sub

 Sub Form_Load()
 
End Sub

Sub modifypass()
Dim user1 As logger
Dim user As logger
ch = List1.List((List1.ListIndex))
user.login = ch 'txtuseramodif.Text
user.pass = txtpassmodif.Text
Dim k As Integer
Dim j As Integer
Dim Done As Boolean
If optadmin.Value = True Then
Open "pass.txt" For Random As #3 Len = Len(user1)
j = 1
Done = False
While Not EOF(3) And Not trouve
Get #3, j, user1
If user1.login = user.login Then
Print
Put #3, j, user
Done = True
End If
j = j + 1
Wend
If Not Done Then
MsgBox "L'utlisateur recherché n'est pas trouvé"
Else: MsgBox "Done"
End If
txtpassmodif.Text = ""
Close #3
ElseIf optuser.Value = True Then
k = FreeFile
Open "users.txt" For Random As #k Len = Len(user1)
j = 1
Done = False
While Not EOF(k) And Not trouve
Get #k, j, user1
If user1.login = user.login Then
Put #k, j, user
Done = True
End If
j = j + 1
Wend
If Not Done Then
MsgBox "L'utlisateur recherché n'est pas trouvé"
txtuseramodif.Text = ""
txtpassmodif.Text = ""
End If
End If
End Sub
Public Sub Remplissage_liste()
Dim t As logger
Dim k As logger
Dim i As Integer
i = 1
Dim n As Integer
n = FreeFile
If optadmin.Value = True Then
Open "pass.txt" For Random As #n Len = Len(t)
While Not EOF(n)
Get n, i, t
k.login = t.login
List1.AddItem (k.login)
i = i + 1
Wend
ElseIf optuser.Value = True Then
Open "users.txt" For Random As #n Len = Len(t)
While Not EOF(n)
Get n, i, t
List1.AddItem (t.login)
i = i + 1
Wend
End If
Close #n
End Sub

'Sub modify_pass()
'Dim newuser As logger
'Dim temp As logger
'Dim j As Integer
'newuser.login = List1.List(j)
'j = List1.ListIndex
'newuser.pass = txtpassmodif.Text
'm = FreeFile
'Open "users.txt" For Random As #m Len = Len(newuser)
''While Not EOF(n)
'If j = 0 Then
'j = j + 1
'Seek m, j
'Put m, j, newuser.pass
'Else
'Put m, j + 2, newuser.pass
'End If
''Wend
'End Sub



Private Sub optadmin_Click()
List1.Clear
Call Remplissage_liste

End Sub

Private Sub optuser_Click()
List1.Clear
Call Remplissage_liste
End Sub
