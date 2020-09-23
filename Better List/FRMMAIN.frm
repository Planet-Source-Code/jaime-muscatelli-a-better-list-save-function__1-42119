VERSION 5.00
Begin VB.Form FRMMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Example // Jaime Muscatelli (Right Click On List)"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDCLEAR 
      Caption         =   "&Clear List"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtLog 
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Text            =   "c:\log.myls"
      ToolTipText     =   "List Location"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton CMDDELETE 
      Caption         =   "&Delete Entry"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton CMDADD 
      Caption         =   "&Add Entry"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton CMDSAVE 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton CMDLOAD 
      Caption         =   "&Load"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ListBox lstData 
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Right Click for Menu"
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton CMDNEW 
      Caption         =   "&New"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblLog 
      Caption         =   "List Location: "
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin VB.Menu mnulistfunctions 
      Caption         =   "FUNCTIONS"
      Visible         =   0   'False
      Begin VB.Menu mnulistadd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnulistremove 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnulistline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnulistclear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "FRMMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMDADD_Click()
Dim sResult As String

sResult = InputBox("List Entry:", "User Input")

If sResult = vbNullString Then
Exit Sub
End If

lstData.AddItem sResult
End Sub

Private Sub CMDCLEAR_Click()
lstData.Clear
End Sub

Private Sub CMDDELETE_Click()
On Error Resume Next
lstData.RemoveItem lstData.ListIndex

End Sub

Private Sub CMDLOAD_Click()

LoadList txtLog.Text, lstData
FRMMAIN.CMDADD.Enabled = True
FRMMAIN.CMDCLEAR.Enabled = True
FRMMAIN.CMDDELETE.Enabled = True
FRMMAIN.CMDLOAD.Enabled = True
FRMMAIN.CMDSAVE.Enabled = True



End Sub

Private Sub CMDNEW_Click()
FRMMAIN.CMDADD.Enabled = True
FRMMAIN.CMDCLEAR.Enabled = True
FRMMAIN.CMDDELETE.Enabled = True
FRMMAIN.CMDLOAD.Enabled = True
FRMMAIN.CMDSAVE.Enabled = True

lstData.Clear


End Sub

Private Sub CMDSAVE_Click()
On Error Resume Next
SaveList txtLog.Text, lstData
End Sub

Private Sub Form_Load()
MsgBox "You'll notice that it is 'log.myls', yet it can be any binary filetype." & vbCrLf & "EX - .txt .log .dll" & vbCrLf & "(and anything you create , such as myls, will default to binary (Binary stores text)"
End Sub

Private Sub lstData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CMDADD.Enabled = False Then Exit Sub
If Button = vbRightButton Then
PopupMenu mnulistfunctions
End If
End Sub

Private Sub mnulistadd_Click()
CMDADD = True
End Sub

Private Sub mnulistclear_Click()
CMDCLEAR = True
End Sub

Private Sub mnulistremove_Click()
CMDDELETE = True
End Sub
