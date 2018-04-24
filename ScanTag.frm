VERSION 5.00
Begin VB.Form ScanSCRAPTAG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scan Scrap Tag"
   ClientHeight    =   1410
   ClientLeft      =   3870
   ClientTop       =   3270
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
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
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SCAN  TAG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1253
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "ScanSCRAPTAG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CANCEL_Click()
    ScanSCRAPTAG.Hide
    Menu.Show
End Sub
' Validates the scrap code and chooses to bring up the JobList or PartList form based on the
' LISTTYPE global variable which was initialized by
Private Sub Command1_Click()
If Len(Text1.Text) = 7 Or Len(Text1.Text) = 8 Then
'    If Left(Text1.Text, 1) = "T" Or Left(Text1.Text, 1) = "t" Then
    If UCase(Left(Text1.Text, 1)) = "T" Or UCase(Left(Text1.Text, 1)) = "H" Then
        If LISTTYPE = "PartList" Then
            PartGroup.Show
            SCRAPTAG = Text1.Text
            Verify.TagNumber.Caption = Text1.Text
            Populate_Parts
        Else
            JobGroup.Show
            SCRAPTAG = Text1.Text
            Verify.TagNumber.Caption = Text1.Text
            Populate_Groups
        End If
        Text1.Text = ""
        ScanSCRAPTAG.Hide
    Else
        Timer1.Enabled = False
        MsgBox ("INVALID TAG - Bad Prefix")
        Timer1.Enabled = True
        Text1.Text = ""
    End If
Else
        Timer1.Enabled = False
        MsgBox ("INVALID TAG - Invalid Length")
        Timer1.Enabled = True
        Text1.Text = ""
End If
LastActivity = Time()
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub



Private Sub Timer1_Timer()
    If MDIForm1.ActiveForm.Name = ScanSCRAPTAG.Name Then
        Text1.SetFocus
    End If
End Sub
