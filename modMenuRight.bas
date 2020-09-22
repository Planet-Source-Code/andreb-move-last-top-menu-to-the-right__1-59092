Attribute VB_Name = "modMenuRight"
' ***************************************************************************
' Module       : modMenuRight.bas in project NCSPowerSofware
'
' Purpose      : Move the last known menu to the right of a form
'
' Special Logic: None
'
' ===========================================================================
' Author       : Andre Beneke
' DateTime     : 22/02/2005 10:43
' ***************************************************************************
Option Explicit

Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Private Const MF_STRING = &H0&
Private Const MF_HELP = &H4000&
Private Const MFS_DEFAULT = &H1000&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" _
    (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, _
    lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" _
    (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, _
    lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Function MoveMenuLeft(FORMhwnd As Long) As Boolean
' ***************************************************************************
' Procedure    : MoveMenuLeft in file modMenuRight.bas
'
' Purpose      : Move the last menu on a form completely to the right
'
' Parameters   :
'
' Return Values: Boolean
'
' Special Logic: Assumption is that there are not more that 100 toplevel menus
'
' ===========================================================================
' Author    : Andre Beneke
' DateTime  : 22/02/2005 10:42
' ***************************************************************************
  Dim mnuItemInfo As MENUITEMINFO
  Dim hMenu       As Long
  Dim BuffStr     As String * 255
  Dim iMenuCount  As Integer

  MoveMenuLeft = False

  hMenu = GetMenu(FORMhwnd)
  BuffStr = Space(80)

  With mnuItemInfo
    .cbSize = Len(mnuItemInfo)
    .dwTypeData = BuffStr & Chr(0)
    .fType = MF_STRING
    .cch = Len(mnuItemInfo.dwTypeData)
    .fState = MFS_DEFAULT
    .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
  End With

  For iMenuCount = 100 To 1 Step -1
    If GetMenuItemInfo(hMenu, iMenuCount, True, mnuItemInfo) <> 0 Then
      Exit For
    End If
  Next iMenuCount

  If iMenuCount > 0 Then
    mnuItemInfo.fType = mnuItemInfo.fType Or MF_HELP
    SetMenuItemInfo hMenu, iMenuCount, True, mnuItemInfo

    ' Repaint top level Menu
    DrawMenuBar (FORMhwnd)

    MoveMenuLeft = True
  End If
End Function


