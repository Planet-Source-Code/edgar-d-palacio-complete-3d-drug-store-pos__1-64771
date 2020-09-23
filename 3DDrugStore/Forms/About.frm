VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "About the System"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4965
      TabIndex        =   7
      Top             =   3375
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   855
      HideSelection   =   0   'False
      Left            =   1725
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "About.frx":058A
      Top             =   2070
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -105
      TabIndex        =   1
      Top             =   3015
      Width           =   6555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.00 Copyright 2006"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   3195
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1725
      TabIndex        =   4
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is license to: Edgar D. Palacio"
      Height          =   375
      Left            =   1725
      TabIndex        =   3
      Top             =   1215
      Width           =   2445
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Designed and Developed by:Edgar D. Palacio. You can contact me at neojohn05@yahoo.com"
      Height          =   975
      Left            =   1725
      TabIndex        =   2
      Top             =   500
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "3D Drug Store  Point On Sale System."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1725
      TabIndex        =   0
      Top             =   195
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   60
      Picture         =   "About.frx":0677
      Stretch         =   -1  'True
      Top             =   900
      Width           =   1590
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
'center the form
Private Sub Form_Activate()
    Call CenterForm(frmAbout)
End Sub

