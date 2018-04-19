VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Locations 
   Caption         =   "List of Locations"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstAvailable 
      Height          =   5025
      Left            =   135
      TabIndex        =   10
      Top             =   510
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   8864
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1543
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Single"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3915
      TabIndex        =   5
      Top             =   3165
      Width           =   1170
      Begin VB.CommandButton cmdSel 
         Height          =   735
         Left            =   135
         Picture         =   "Locations.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   525
         Width           =   855
      End
      Begin VB.CommandButton cmdunSel 
         Height          =   735
         Left            =   135
         Picture         =   "Locations.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1410
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3915
      TabIndex        =   4
      Top             =   690
      Width           =   1170
      Begin VB.CommandButton cmdSelAll 
         Height          =   735
         Left            =   150
         Picture         =   "Locations.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton cmdUnSelAll 
         Height          =   735
         Left            =   150
         Picture         =   "Locations.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1350
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
   Begin MSComctlLib.ListView lstSelected 
      Height          =   5100
      Left            =   5220
      TabIndex        =   11
      Top             =   600
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   8996
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1543
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
     Locations.Hide
     Settings.Show
     
End Sub

Private Sub cmdSave_Click()
    writeLocationSettings
    Locations.Hide
    Settings.Show
End Sub

Private Sub cmdSel_Click()
    Dim cnt As Integer
    Dim Sel() As Integer
    Dim itmx As ListItem

    cnt = -1
    For i = 1 To lstAvailable.ListItems.Count
        If lstAvailable.ListItems(i).Selected Then
            cnt = cnt + 1
            ReDim Preserve Sel(cnt)
            Set itmx = lstSelected.ListItems.Add(, , lstAvailable.ListItems(i).Text)
            itmx.SubItems(1) = lstAvailable.ListItems(i).ListSubItems(1).Text
            Sel(cnt) = i
            End If
        Next i
        
    Sort Sel(), cnt
    
    For j = cnt To 0 Step -1
        lstAvailable.ListItems.Remove (Sel(j))
        Next j
        End Sub

Private Sub cmdSelAll_Click()
    Dim cnt As Integer
    Dim Sel() As Integer
    Dim itmx As ListItem
    
    cnt = -1
    For i = 1 To lstAvailable.ListItems.Count
        cnt = cnt + 1
        ReDim Preserve Sel(cnt)
        Set itmx = lstSelected.ListItems.Add(, , lstAvailable.ListItems(i).Text)
        itmx.SubItems(1) = lstAvailable.ListItems(i).ListSubItems(1).Text
        Sel(cnt) = i
        Next i
        
    Sort Sel(), cnt
    
    For j = cnt To 0 Step -1
        lstAvailable.ListItems.Remove (Sel(j))
        Next j

End Sub

Private Sub cmdunSel_Click()
    Dim cnt As Integer
    Dim Sel() As Integer
    Dim itmx As ListItem
    
    cnt = -1
    For i = 1 To lstSelected.ListItems.Count
        If lstSelected.ListItems(i).Selected Then
            cnt = cnt + 1
            ReDim Preserve Sel(cnt)
            Set itmx = lstAvailable.ListItems.Add(, , lstSelected.ListItems(i).Text)
            itmx.SubItems(1) = lstSelected.ListItems(i).ListSubItems(1).Text
            Sel(cnt) = i
            End If
        Next i
        
    Sort Sel(), cnt
    
    For j = cnt To 0 Step -1
        lstSelected.ListItems.Remove (Sel(j))
        Next j
         
End Sub

Private Sub cmdUnSelAll_Click()
    Dim cnt As Integer
    Dim Sel() As Integer
    Dim itmx As ListItem
    
    cnt = -1
    For i = 1 To lstSelected.ListItems.Count
        cnt = cnt + 1
        ReDim Preserve Sel(cnt)
        Set itmx = lstAvailable.ListItems.Add(, , lstSelected.ListItems(i).Text)
        itmx.SubItems(1) = lstSelected.ListItems(i).ListSubItems(1).Text
'        lstAvailable.AddItem lstSelected.List(i)
        Sel(cnt) = i
        Next i
        
    Sort Sel(), cnt
        
    For j = cnt To 0 Step -1
        lstSelected.ListItems.Remove (Sel(j))
        Next j
        

End Sub


Private Sub Form_Load()
    '****   Populate the Available locations listbox
    Populate_Locations

End Sub

Private Sub lstAvailable_DblClick()
    cmdSel_Click
End Sub

Private Sub lstSelected_DblClick()
    cmdunSel_Click
End Sub


Private Sub Sort(ByRef Val() As Integer, cnt As Integer)
    For i = 0 To cnt
        For k = 0 To cnt
            If Val(i) < Val(k) Then
                tmp = Val(i)
                Val(i) = Val(k)
                Val(k) = tmp
                End If
            Next k
        Next i
End Sub
