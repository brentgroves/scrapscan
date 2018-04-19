VERSION 5.00
Begin VB.Form Verify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify"
   ClientHeight    =   5220
   ClientLeft      =   7350
   ClientTop       =   5325
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
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
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton OKButton 
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
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Tag Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   3600
      Width           =   2715
   End
   Begin VB.Label TagNumber 
      Caption         =   "tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   3600
      Width           =   2715
   End
   Begin VB.Label Pieces_Label 
      Caption         =   "Pieces"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   3120
      Width           =   2715
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieces:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   2715
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Function:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   2715
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   2715
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Job Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Part Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   2715
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Scrap Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   2715
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Operation Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   2715
   End
   Begin VB.Label functionLabel 
      Caption         =   "Function:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   2715
   End
   Begin VB.Label OperationNumber 
      Caption         =   "Operation Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   2715
   End
   Begin VB.Label SCRAPCODE 
      Caption         =   "scrapcode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   2715
   End
   Begin VB.Label PartNum 
      Caption         =   "Part Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   2715
   End
   Begin VB.Label JobNum 
      Caption         =   "Job Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   2715
   End
   Begin VB.Label emp_name 
      Caption         =   "Employee Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   2715
   End
End
Attribute VB_Name = "Verify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    Verify.Hide
    Menu.Show
    LastActivity = Time()
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub

Private Sub OKButton_Click()
    'WriteRecord
    Verify.Hide
    Menu.Show
    LastActivity = Time()
    Dim itmx As ListItem
    Set itmx = Menu.ListView1.ListItems.Add(, , "E" + Trim(empNumber))
    itmx.SubItems(1) = Trim(global_jobnumber)
    itmx.SubItems(2) = Trim(opNumber)
    itmx.SubItems(3) = Trim(SCRAPCODE)
    itmx.SubItems(4) = Trim(Piece_Count)
    itmx.SubItems(5) = Trim(TagNumber)
    itmx.SubItems(6) = Trim(PartNum)
    itmx.SubItems(7) = Trim(PartRev)
End Sub
