VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PartGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PartList"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11205
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Jobs"
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
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
   Begin MSComctlLib.ListView List1 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   0
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
End
Attribute VB_Name = "PartGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'PartGroup Form
'If you want this form to come up first when you are adding records then enter PartList into settings.ine
'File.  This form is used to add a scrap tag if you know the partnumber.
Private Sub Command1_Click()
    List1.ListItems.Clear
    PartGroup.Hide
    Menu.Show
    LastActivity = Time()
End Sub
' Activated when user clicks the Accept button once a part number has been selected in the listview control.
' Calls Enter_Count which brings up the OpList form.
Private Sub Command2_Click()
    gPartNumber = List1.SelectedItem.Text
    gPartRev = List1.SelectedItem.ListSubItems.Item(1).Text
    global_jobnumber = List1.SelectedItem.ListSubItems.Item(2).Text
    List1.ListItems.Clear
    Enter_Count
    PartGroup.Hide
    LastActivity = Time()
End Sub
' Goes to the form which allows the user to select the part family instead of just a part number.
Private Sub Command3_Click()
    List1.ListItems.Clear
    Populate_Groups
    JobGroup.Show
    PartGroup.Hide
    LastActivity = Time()

End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

