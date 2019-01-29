VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrecherche 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12945
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
      Left            =   6360
      MousePointer    =   4  'Icon
      Picture         =   "frmrecherche.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox txtprenom 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&QUITTER"
      Height          =   975
      Left            =   8880
      TabIndex        =   7
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RECHERCHER"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   6495
      Begin MSComCtl2.DTPicker dtpnaiss 
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   3000
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   42487
      End
      Begin VB.TextBox txtnom 
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chkdate 
         Caption         =   "Date de naissance"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox chknom 
         Caption         =   "Nom"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkprenom 
         Caption         =   "Prénom"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Veuillez préciser le/les critère(s) de recherche:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RECHERCHE D'UNE PERSONNE"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   7575
   End
End
Attribute VB_Name = "frmrecherche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkdate_Click()
'dtpnaiss.Visible = True

If chkdate.Value = False Then
dtpnaiss.Visible = False
Else
dtpnaiss.Visible = True
End If
End Sub

Private Sub chknom_Click()
'txtnom.Visible = True
If chknom.Value = False Then
txtnom.Visible = False
Else
txtnom.Visible = True
End If
End Sub

Private Sub chkprenom_Click()
'txtprenom.Visible = True
If chkprenom.Value = False Then
txtprenom.Visible = False
Else
txtprenom.Visible = True
End If

End Sub

    

'Private Sub Command1_Click()
'Dim name As person
'Dim name1 As person
'Dim j As Integer
'If chknom.Value = True Then
'name.nom = txtnom.Text
'Open "persons.txt" For Random As #1 Len = Len(name)
'j = 1
''trouve = False
'While Not EOF(1) 'And Not trouve
'Get #1, j, name1
'If InStr(name1.nom, name.nom) = 1 Then
'frmresultsearch.List1.AddItem (name1.nom)
''trouve = True
'j = j + 1
'End If
'Wend
''If Not trouve Then
''MsgBox "Pas de personne possédant ce nom"
''End If
'Close #1
'End If
'frmresultsearch.Show
'
'End Sub
 Public Sub Remplissage_list()
List1.Clear
Dim t As person
Dim k As person
Dim i As Integer
i = 1
Dim n As Integer
Dim trouve As Boolean
trouve = False
f = FreeFile
Open "Persons1.txt" For Random As #f Len = Len(t)
'Recherche selon le nom seulement
If chknom.Value = 1 Then
k.nom = txtnom.Text
While Not EOF(f) 'And trouve =True
Get f, i, t
If k.nom = t.nom Then

frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
'trouve = True
End If
i = i + 1
Wend
'frmresultsearch.Show
'Recherche selon le prénom seulement
ElseIf chkprenom.Value = 1 Then
k.prenom = txtprenom.Text
While Not EOF(f)
Get f, i, t
If k.prenom = t.prenom Then
frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
'trouve = True
End If
i = i + 1
Wend
'frmresultsearch.Show
'Recherche selon La date de naissance seulement
ElseIf chkdate.Value = 1 Then
k.Datenaiss = dtpnaiss.Value
While Not EOF(f)
Get f, i, t
If k.prenom = t.prenom Then
frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
'trouve = True
End If
i = i + 1
Wend
'frmresultsearch.Show
'If frmresultsearch.List1.ListCount = 1 Then
'MsgBox "La personne recherché n'existe pas"
'End If
'Recherche selon le nom et le prénom
ElseIf chknom.Value = 1 And chkprenom.Value = 1 Then
While Not EOF(f)
Get f, i, t
k.prenom = txtprenom.Text
k.nom = txtnom.Text
If k.prenom = t.prenom And k.nom = t.nom Then
frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
'trouve = True
End If
i = i + 1
Wend
'frmresultsearch.Show
'Recherche selon le nom et la date de naissance
ElseIf chknom.Value = 1 And chkdate.Value = 1 Then
k.nom = txtnom.Text
k.Datenaiss = dtpnaiss.Value
While Not EOF(f)
Get f, i, t
If k.nom = t.nom And k.Datenaiss = t.Datenaiss Then
frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
End If
i = i + 1
Wend
'frmresultsearch.Show
'Recherche selon le prénom et la date de naissance
ElseIf chkprenom.Value = 1 And chkdate.Value = 1 Then
k.prenom = txtprenom.Text
k.Datenaiss = dtpnaiss.Value
While Not EOF(f)
Get f, i, t
If k.prenom = t.prenom And k.Datenaiss = t.Datenaiss Then
frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
End If
i = i + 1
Wend
'frmresultsearch.Show
'Recherche selon les 3 criteres Nom,prénom et date de naissance
ElseIf chkprenom.Value = 1 And chknom.Value = 1 And chkdate.Value = 1 Then
k.nom = txtnom.Text
k.prenom = txtprenom.Text
k.Datenaiss = dtpnaiss.Value
While Not EOF(f)
Get f, i, t
If k.prenom = t.prenom And k.nom = t.nom And k.Datenaiss = t.Datenaiss Then
frmresultsearch.List1.AddItem (t.nom)
frmresultsearch.List2.AddItem (t.prenom)
frmresultsearch.List3.AddItem (t.sexe)
frmresultsearch.List4.AddItem (t.Datenaiss)
frmresultsearch.List5.AddItem (t.Lieunaiss)
frmresultsearch.List6.AddItem (t.etatciv)
frmresultsearch.List7.AddItem (t.nompere)
frmresultsearch.List8.AddItem (t.nommere)
frmresultsearch.List9.AddItem (t.Dateannonce)
frmresultsearch.List10.AddItem (t.nomannonce)
frmresultsearch.List11.AddItem (t.nomofficier)
End If
i = i + 1
Wend
'frmresultsearch.Show
Else: MsgBox "veuillez choisir au moins un critère"
End If
Close #f

If frmresultsearch.List1.ListCount = 0 And frmresultsearch.List4.ListCount = 0 Then
MsgBox "La personne recherché n'existe pas"
Else
frmresultsearch.Show
End If
End Sub

Private Sub cmdretourmnu_Click()
Frmchoix.Show
Me.Hide
End Sub

Private Sub Command1_Click()
Call Remplissage_list
'Dim name As person
'Dim name1 As person
'Dim j As Integer
'If chknom.Value = True Then
'name.nom = txtnom.Text
'Open "persons.txt" For Random As #1 Len = Len(name)
'j = 1
''trouve = False
'While Not EOF(1) 'And Not trouve
'Get #1, j, name1
'If InStr(name1.nom, name.nom) = 1 Then
'frmresultsearch.List1.AddItem (name1.nom)
''trouve = True
'j = j + 1
'End If
'Wend
''If Not trouve Then
''MsgBox "Pas de personne possédant ce nom"
''End If
'Close #1
'End If
'frmresultsearch.Show

End Sub

'Control de saisie
'
'Private Sub txtnom_KeyPress(KeyAscii As Integer)
'Dim Temp As String
' Dim i As Integer
'    Dim c As String * 1
'    Dim name As String
'    name = txtnom.Text
'    i = 0
'        c = Mid$(name, i + 1, 1)
'     If (Asc(c) >= Asc("A") And Asc(c) <= Asc("Z")) Or (Asc(c) >= Asc("a") And Asc(c) <= Asc("z")) Then
'    i = i + 1
'    Else
'    MsgBox "Nom Invalide"
'    txtnom.Text = ""
'    End If
'
'End Sub
'
'Private Sub txtnom_LostFocus()
' Dim name As String
' Dim s As String
'        name = txtnom.Text
'        Dim sp As Integer
'        Dim k As Integer
'        For i = 2 To Len(name) + 1
'         's =
'            If (Asc(Mid("name", i, 1)) >= Asc("A") And Asc(Mid("name", i, 1)) <= Asc("Z")) Or (Asc(Mid("name", i, 1)) >= Asc("a") And Asc(Mid("name", i, 1)) >= Asc("z")) Then
'                 k = k + 1
'            ElseIf Mid("name", i, 1) = " " And Mid("name", i - 1, 1) <> " " Then
'                sp = sp + 1
'            Else
'                sp = 2
'            End If
'        Next
'        If sp = 1 Then
'         MsgBox "Nom valide"  'Valid Name
'        Else
'          MsgBox "Nom Invalide"  'Invalid Name
'        End If
'End Sub
