VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Caption         =   "[ Note ]"
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "This also works with MDI Forms"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   4455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Close"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Exit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Index"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  MoveMenuLeft FORMhwnd:=frmTest.hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' ------------------------------------------------------------
' Confirmation of program exit
' ------------------------------------------------------------
  Select Case UnloadMode
    Case vbFormCode, vbFormControlMenu
      If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Confirmation to exit") Then
        Cancel = False
      Else
        Cancel = True
      End If
    Case vbAppWindows
      Call MsgBox("Close Program Command received from the Windows Environment Engine", vbOKOnly + vbInformation, "Confirmation to exit")
      Cancel = True
    Case vbAppTaskManager
      Call MsgBox("Close Program Command received from the TaskManager", vbOKOnly + vbInformation, "Confirmation to exit")
      Cancel = True
    Case Else
      Cancel = False
  End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
  Select Case Index
    Case 3
      Unload Me
    Case Else
  End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
  Select Case Index
    Case 0
      Call MsgBox("Menu index to be displayed here", vbOKOnly + vbInformation, "Help selection")
    Case Else
  End Select
End Sub
