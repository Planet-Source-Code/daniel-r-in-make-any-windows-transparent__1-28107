VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API : Windows ""hwnd"" Example"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSize 
      Caption         =   " Windows Effects"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Text            =   "128"
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set transparency"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Value"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Windows : "
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdListWindows 
         Caption         =   "Refresh Windows"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   3720
         Width           =   5175
      End
      Begin VB.ListBox listWindows 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdListWindows_Click()
' ---------------------------------------------------------------------------------- '
    Set TargetList = frmMain.listWindows
    TargetList.Clear
    EnumWindows AddressOf EnumWindowsProc, 0
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub cmdMinimize_Click()
' ---------------------------------------------------------------------------------- '
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Minimize
        End If
    End If
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub cmdNormal_Click()
' ---------------------------------------------------------------------------------- '
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Normal
        End If
    End If
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub cmdMaximize_Click()
' ---------------------------------------------------------------------------------- '
    If listWindows.ListCount > 0 Then
        If listWindows.ListIndex > -1 Then
            ShowWindow listWindows.ItemData(listWindows.ListIndex), SW_Maximize
        End If
    End If
' ---------------------------------------------------------------------------------- '
End Sub

Private Sub Command1_Click()
Call MakeTransparent(listWindows.ItemData(listWindows.ListIndex), Text1.Text)
End Sub

Private Sub Form_Load()
cmdListWindows_Click
End Sub
