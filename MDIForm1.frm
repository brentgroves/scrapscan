VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H00400000&
   Caption         =   "Scrap Scan V1.10.2"
   ClientHeight    =   11145
   ClientLeft      =   3375
   ClientTop       =   1980
   ClientWidth     =   16785
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10770
      Width           =   16785
      _ExtentX        =   29607
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Object.Width           =   9657
            TextSave        =   "6/17/2015"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   9657
            TextSave        =   "8:43 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9657
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   240
      Top             =   6840
   End
   Begin VB.Menu CONFIGURE 
      Caption         =   "CONFIGURE"
   End
   Begin VB.Menu PLANTSELECTION 
      Caption         =   "PLANT"
      Begin VB.Menu P2 
         Caption         =   "PLANT 2"
         Checked         =   -1  'True
      End
      Begin VB.Menu P3 
         Caption         =   "PLANT 3"
         Checked         =   -1  'True
      End
      Begin VB.Menu P5 
         Caption         =   "PLANT 5"
         Checked         =   -1  'True
      End
      Begin VB.Menu P6 
         Caption         =   "PLANT 6"
         Checked         =   -1  'True
      End
      Begin VB.Menu P7 
         Caption         =   "PLANT 7"
         Checked         =   -1  'True
      End
      Begin VB.Menu P8A 
         Caption         =   "PLANT 8"
         Checked         =   -1  'True
      End
      Begin VB.Menu P9 
         Caption         =   "PLANT 9"
         Checked         =   -1  'True
      End
      Begin VB.Menu P11 
         Caption         =   "PLANT 11"
         Checked         =   -1  'True
      End
      Begin VB.Menu P8 
         Caption         =   "DISTRIBUTION CENTER"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CONFIGURE_Click()
    MDIForm1.ActiveForm.Hide
    LastActivity = Time
    frmLogin.Show
    
End Sub

Private Sub MDIForm_Activate()
    OpList.List1.ColumnHeaders.Add , , "OPERATION NUMBER", 2500
    OpList.List1.ColumnHeaders.Add , , "WORKCENTER", (JobList.List1.Width - 2500) / 2 - 500
    OpList.List1.ColumnHeaders.Add , , "OPERATION DESCRIPTION", (JobList.List1.Width - 2500) / 2 + 200
    JobList.List1.ColumnHeaders.Add , , "JOB NUMBER", 4500
    JobList.List1.ColumnHeaders.Add , , "PART NUMBER", (JobList.List1.Width - 2000)
    PartGroup.List1.ColumnHeaders.Add , , "PART NUMBER", 4000
    PartGroup.List1.ColumnHeaders.Add , , "REV", (JobList.List1.Width - 4000) / 6
    PartGroup.List1.ColumnHeaders.Add , , "JOB NAME", (JobList.List1.Width - 3500 - ((JobList.List1.Width - 3500) / 4))
    JobGroup.List1.ColumnHeaders.Add , , "JOB GROUP", 8000
    EmpList.List1.ColumnHeaders.Add , , "EMPLOYEE #", 2000
    EmpList.List1.ColumnHeaders.Add , , "EMPLOYEE NAME", (JobList.List1.Width - 2000)
    ScrapCodeList.List1.ColumnHeaders.Add , , "SCRAP CODE", 2000
    ScrapCodeList.List1.ColumnHeaders.Add , , "SCRAP CODE DESCRIPTION", (JobList.List1.Width - 2000)
    BOMList.List1.ColumnHeaders.Add , , "PART NUMBER", 4000
    BOMList.List1.ColumnHeaders.Add , , "REV", (JobList.List1.Width - 4000) / 6
    Me.Caption = App.ProductName + " v" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    getSettings
    init_db
    Menu.Show
    Menu.START_PICT.Visible = True
    Menu.END_PICTR.Visible = False
    Menu.CANCEL_PICT.Visible = False
    Menu.ListView1.Visible = False
    GetBCMast
    P2.Checked = False
    P3.Checked = True
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False
    
    PLANT = 10
    
    getLocationSettings
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    closedb
End Sub

Private Sub P2_Click()
    P2.Checked = True
    P3.Checked = False
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False

    PLANT = 20
End Sub

Private Sub P3_Click()
    P2.Checked = False
    P3.Checked = True
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False
        PLANT = 10
End Sub
Private Sub P5_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = True
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False
        PLANT = 50
End Sub
Private Sub P6_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = False
    P6.Checked = True
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False
    PLANT = 60
End Sub
Private Sub P7_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = False
    P6.Checked = False
    P7.Checked = True
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False
    PLANT = 77
End Sub
Private Sub P8_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = True
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = False
    PLANT = 85
End Sub
Private Sub P8A_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = False
    P8A.Checked = True
    PLANT = 88
End Sub
Private Sub P9_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = True
    P11.Checked = False
    P8A.Checked = False
    PLANT = 99
End Sub
Private Sub P11_Click()
    P2.Checked = False
    P3.Checked = False
    P5.Checked = False
    P6.Checked = False
    P7.Checked = False
    P8.Checked = False
    P9.Checked = False
    P11.Checked = True
    P8A.Checked = False
    PLANT = 30
End Sub
Private Sub Timer1_Timer()
    If Me.ActiveForm.Name <> Menu.Name Then
        If DateDiff("s", LastActivity, Time()) > 90 Then
            Me.ActiveForm.Hide
            JobList.List1.ListItems.Clear
            OpList.List1.ListItems.Clear
            Menu.Show
        End If
    Else
        If DateDiff("s", LastActivity, Time()) > 30 Then
            GetBCMast
        End If
        If DateDiff("s", RefreshDB, Time()) > 15 Then
            closedb
            init_db
            RefreshDB = Time()
        End If
        
    End If
End Sub
