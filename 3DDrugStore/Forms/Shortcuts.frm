VERSION 5.00
Begin VB.Form frmShortcuts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Shortcuts"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Shortcuts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   113
      TabIndex        =   0
      Top             =   45
      Width           =   4185
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "F11"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   2745
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   21
         Top             =   2745
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Exit Application"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   20
         Top             =   3030
         Width           =   1305
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   19
         Top             =   2465
         Width           =   660
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "View Products"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   18
         Top             =   2185
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   17
         Top             =   1905
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Product  (Admin Only)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   16
         Top             =   1625
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F1"
         Height          =   195
         Left            =   375
         TabIndex        =   15
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F2"
         Height          =   195
         Left            =   375
         TabIndex        =   14
         Top             =   505
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F3"
         Height          =   195
         Left            =   375
         TabIndex        =   13
         Top             =   785
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F4"
         Height          =   195
         Left            =   375
         TabIndex        =   12
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "F5"
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   1345
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F6"
         Height          =   195
         Left            =   375
         TabIndex        =   10
         Top             =   1625
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "F7"
         Height          =   195
         Left            =   375
         TabIndex        =   9
         Top             =   1905
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "F8"
         Height          =   195
         Left            =   375
         TabIndex        =   8
         Top             =   2185
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "F9"
         Height          =   195
         Left            =   375
         TabIndex        =   7
         Top             =   2465
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "F12"
         Height          =   195
         Left            =   255
         TabIndex        =   6
         Top             =   3030
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   5
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Category Maintenance (Admin Only)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   4
         Top             =   505
         Width           =   3120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Maintenance  (Admin Only)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   3
         Top             =   785
         Width           =   3090
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "User Maintenace  (Admin Only)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   2
         Top             =   1065
         Width           =   2670
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Delivery  (Admin Only)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   870
         TabIndex        =   1
         Top             =   1345
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
