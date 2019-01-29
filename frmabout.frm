VERSION 5.00
Begin VB.Form frmabout 
   Caption         =   "Form2"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5550
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   6855
      Begin VB.Label Label2 
         Caption         =   "COPYRIGHT ©2016 HUIJI HAYTHEM ET MESSAOUDI FIRAS                            TOUS LES DROITS SONT RÉSERVÉS "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   6015
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   240
      Picture         =   "frmabout.frx":000C
      ScaleHeight     =   615
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label txtabout 
      Caption         =   "Application de Gestion D'etat Civil "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = "COPYRIGHT ©2016 HUIJI HAYTHEM ET MESSAOUDI FIRAS                            TOUS LES DROITS SONT RÉSERVÉS. Pour contacter: huiji.haythem@ outlook.com"
End Sub

