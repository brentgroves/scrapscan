VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form JobGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JobList"
   ClientHeight    =   7275
   ClientLeft      =   7350
   ClientTop       =   5325
   ClientWidth     =   11730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11730
   Begin VB.CommandButton Command3 
      Caption         =   "&Part Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Accept"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2295
   End
   Begin MSComctlLib.ListView List1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9975
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2295
   End
End
Attribute VB_Name = "JobGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    List1.ListItems.Clear
    JobGroup.Hide
    Menu.Show
    LastActivity = Time()
End Sub

Private Sub Command2_Click()
    JobGrp = Trim(List1.SelectedItem.Text)
    List1.ListItems.Clear
    Populate_Jobs
    JobGroup.Hide
    LastActivity = Time()
End Sub

Private Sub Command3_Click()
    JobGrp = Trim(List1.SelectedItem.Text)
    List1.ListItems.Clear
    Populate_Parts
    JobGroup.Hide
    LastActivity = Time()

End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

