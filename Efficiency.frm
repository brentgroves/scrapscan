VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Efficiency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Efficiency"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10860
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   975
      Left            =   8640
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6135
      Left            =   120
      OleObjectBlob   =   "Efficiency.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Efficiency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Menu.Show
LastActivity = Time()
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub
