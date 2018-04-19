VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu"
   ClientHeight    =   7665
   ClientLeft      =   7350
   ClientTop       =   5325
   ClientWidth     =   11565
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFromLocation 
      Height          =   450
      Left            =   7590
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   7125
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.ComboBox cboToLocation 
      Height          =   450
      Left            =   7590
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   3360
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9340
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee Number"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Job Number"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Operation Number"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Scrap Code"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Pieces"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Tag Number"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Material Number"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Material Rev."
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "To Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "From Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Delete_Checked 
      Caption         =   "Delete Checked Records"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "DELETE"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox add_pict 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   600
      Picture         =   "Form1.frx":031A
      ScaleHeight     =   1065
      ScaleWidth      =   3240
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.PictureBox CANCEL_PICT 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   1065
      Left            =   7680
      Picture         =   "Form1.frx":5746
      ScaleHeight     =   1065
      ScaleWidth      =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   3240
   End
   Begin VB.PictureBox END_PICTR 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   4200
      Picture         =   "Form1.frx":A567
      ScaleHeight     =   1065
      ScaleWidth      =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   3240
   End
   Begin VB.PictureBox START_PICT 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   4200
      Picture         =   "Form1.frx":FE7D
      ScaleHeight     =   1065
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   3240
   End
   Begin VB.Label lblSetFromLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Select ""From"" Location"
      Height          =   360
      Left            =   4290
      TabIndex        =   9
      Top             =   7140
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Label lblSetToLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "Select ""To"" Location"
      Height          =   360
      Left            =   4470
      TabIndex        =   7
      Top             =   6615
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub add_pict_Click()
    Menu.Hide
    LastActivity = Time()
    ScanSCRAPTAG.Show
End Sub

Private Sub add_pict_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And add_pict.Visible = True Then
        add_pict_Click
    End If
    
End Sub

Private Sub CANCEL_PICT_Click()
If ListView1.ListItems.Count > 0 Then
    If MsgBox("You have unsaved records, Do you stil want to exit?", vbYesNo, "Exit?") = vbNo Then
        Exit Sub
    End If
End If
    ListView1.ListItems.Clear
     
    START_PICT.Visible = True
    CANCEL_PICT.Visible = False
    END_PICTR.Visible = False
    add_pict.Visible = False
    ListView1.Visible = False
    Delete_Checked.Visible = False
    JobGroup.Hide
    LastActivity = Time()
    MDIForm1.P2.Enabled = True
    MDIForm1.P3.Enabled = True
    MDIForm1.P5.Enabled = True
    MDIForm1.P6.Enabled = True
    MDIForm1.P7.Enabled = True
    MDIForm1.P8.Enabled = True
    MDIForm1.P9.Enabled = True
    MDIForm1.P11.Enabled = True
    MDIForm1.P8A.Enabled = True
    cboToLocation.Visible = False
    cboFromLocation.Visible = False
    lblSetToLocation.Visible = False
    lblSetFromLocation.Visible = False
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub

Private Sub cboToLocation_Click()
    Dim Idx As Integer
    Idx = ListView1.SelectedItem.Index

'MsgBox ("Test selection index " + Str(Idx))
    '            cboLocation.Text = ListView1.ListItems(Idx).ListSubItems(7).Text
    '(cboToLocation.Text <> cboToLocation.Tag) And
    If (cboToLocation.ListIndex <> -1) Then
        cboToLocation.Tag = cboToLocation.Text
        ListView1.ListItems(Idx).SubItems(8) = cboToLocation.Text
        End If
        
End Sub

Private Sub cboFromLocation_Click()
    Dim Idx As Integer
    Idx = ListView1.SelectedItem.Index

'MsgBox ("Test selection")
    '            cboLocation.Text = ListView1.ListItems(Idx).ListSubItems(7).Text
    '(cboFromLocation.Text <> cboFromLocation.Tag) And
    If (cboFromLocation.ListIndex <> -1) Then
        cboFromLocation.Tag = cboFromLocation.Text
        ListView1.ListItems(Idx).SubItems(9) = cboFromLocation.Text
        End If
        
End Sub


Private Sub Delete_Checked_Click()
    Dim i As Integer
    If ListView1.ListItems.Count <> 0 Then
        For i = ListView1.ListItems.Count To 1 Step -1
            If ListView1.ListItems(i).Checked Then
                ListView1.ListItems.Remove (i)
            End If
        Next
    End If
    
    
End Sub

Private Sub END_PICTR_Click()
    MDIForm1.StatusBar1.Panels(3).Text = "TRANSFERRING...."
    Screen.MousePointer = vbHourglass
    WriteRecords
    Screen.MousePointer = vbNromal
    CANCEL_PICT_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And add_pict.Visible = True Then
        add_pict_Click
    End If
    
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ListView1.ListItems.Count > 0 Then
        If MsgBox("You have unsaved records, Do you really want to exit?", vbYesNo, "Exit?") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub ListView1_Click()
    Dim Idx As Integer
    Idx = ListView1.SelectedItem.Index

    If UCase(Left(Trim(ListView1.ListItems(Idx).ListSubItems(5).Text), 1)) = "H" Then
            lblSetToLocation.Visible = True
            cboToLocation.Visible = True
            cboToLocation.ListIndex = -1
            lblSetFromLocation.Visible = True
            cboFromLocation.Visible = True
            cboFromLocation.ListIndex = -1
            For i = 0 To cboToLocation.ListCount - 1
                If cboToLocation.List(i) = ListView1.ListItems(Idx).SubItems(8) Then
                    cboToLocation.ListIndex = i
                    End If
                    Next i
                    
            For i = 0 To cboFromLocation.ListCount - 1
                If cboFromLocation.List(i) = ListView1.ListItems(Idx).SubItems(9) Then
                    cboFromLocation.ListIndex = i
                    End If
                    Next i
'            cboLocation.Text = ListView1.ListItems(Idx).ListSubItems(7).Text
        Else
            cboToLocation.Visible = False
            cboFromLocation.Visible = False
            lblSetToLocation.Visible = False
            lblSetFromLocation.Visible = False
       End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And add_pict.Visible = True Then
        add_pict_Click
    End If
End Sub

Private Sub START_PICT_Click()
    START_PICT.Visible = False
    END_PICTR.Visible = True
    CANCEL_PICT.Visible = True
    ListView1.Visible = True
    add_pict.Visible = True
    Delete_Checked.Visible = True
    MDIForm1.P2.Enabled = False
    MDIForm1.P3.Enabled = False
    MDIForm1.P5.Enabled = False
    MDIForm1.P6.Enabled = False
    MDIForm1.P7.Enabled = False
    MDIForm1.P8.Enabled = False
    MDIForm1.P9.Enabled = False
    MDIForm1.P11.Enabled = False
    MDIForm1.P8A.Enabled = False
    MDIForm1.StatusBar1.Panels(3).Text = ""
'    cboLocation.Visible = True
    cboToLocation.Clear
    For i = 1 To Locations.lstSelected.ListItems.Count
        cboToLocation.AddItem Locations.lstSelected.ListItems(i)
        Next i
        
    cboFromLocation.Clear
    For i = 1 To Locations.lstSelected.ListItems.Count
        cboFromLocation.AddItem Locations.lstSelected.ListItems(i)
        Next i
        
    ListView1.SetFocus
End Sub
