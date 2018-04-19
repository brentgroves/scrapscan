VERSION 5.00
Begin VB.Form Settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   6210
   ClientLeft      =   6495
   ClientTop       =   5205
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Settings.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdLocation 
      Caption         =   "Locations"
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtListType 
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox M2MPASS 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox M2MUSER 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox M2MSERVER_TEXT 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CANCEL 
      Caption         =   "CANCEL"
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox PLANT 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox VFPTABLE 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox M2MDB 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "LIST TYPE"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "M2M PASSWORD:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "M2M USERNAME:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "M2M SERVER:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "PLANT:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "VFP TABLE:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "M2M DATABASE:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CANCEL_Click()
Settings.Hide
'Scan.Show
'MDIForm1.Show
Menu.Show
LastActivity = Time()
End Sub


Private Sub cmdLocation_Click()
'    Settings.Hide
    Locations.Show
End Sub


Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub

Private Sub OK_Click()
 closedb
 init_db
 writesettings
 Settings.Hide
 Menu.Show
 LastActivity = Time()
End Sub
