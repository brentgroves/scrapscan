VERSION 5.00
Begin VB.Form Pieces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pieces"
   ClientHeight    =   6375
   ClientLeft      =   7350
   ClientTop       =   5325
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   3825
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin ScrapScan.PieceCount PieceCount1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8281
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ENTER PIECES"
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
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Pieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim temp As String
temp = PieceCount1.PIECECOUNT
If IsNumeric(temp) Then
    Piece_Count = Val(Trim(temp))
    If Piece_Count > 50 Then
        If MsgBox("Are you sure you want to scrap " + Str(Piece_Count) + "?", vbYesNo, "?!!?") = vbNo Then
            Exit Sub
        End If
    End If
    Pieces.Hide
    Populate_ScrapCodeList
    ScrapCodeList.Show
    LastActivity = Time()
    Verify.Pieces_Label = Piece_Count
    Verify.functionLabel.Caption = "SCRAP INPUT"
Else
    MsgBox ("Invalid Quantity")
End If
End Sub

Private Sub Command2_Click()
    Pieces.Hide
    Menu.Show
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub
