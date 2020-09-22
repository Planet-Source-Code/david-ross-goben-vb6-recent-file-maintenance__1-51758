Attribute VB_Name = "modListviewMouseUp"
Option Explicit
'~modListViewMouseUp.bas;
'Returns the ListView control's Row and Column that the user raised the left mouse button over
'******************************************************************************
' modListViewMouseUp:
' The ListViewMouseUp() subroutine, provided X/Y parameters from the user's
'                       ListView_MouseUp event, this routine returns the
'                       ListView control's Row and Column that the user raised
'                       the left mouse button over.
'EXAMPLE:
' Dim LV1Row As Integer, LV1Col As Integer
'
'Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Call ListViewMouseUp(ListView1, X, Y, LV1Row, LV1Col) 'Save row and column
'End Sub
'******************************************************************************

'****************************************************
' Types, constants, API calls
'****************************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_SUBITEMHITTEST As Long = 4153

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type LVHITTESTINFO
  pt As POINTAPI
  lngFlags As Long
  lngItem As Long
  lngSubItem As Long
End Type

'****************************************************
' ListViewMouseUp(): Provided X/Y parameters from the user's ListView_MouseUp
'                    event, this routine returns the ListView control's Row
'                    and Column that the user raised a mouse button over.
'****************************************************
Public Sub ListViewMouseUp(LVW As ListView, X As Single, Y As Single, Row As Integer, Col As Integer)
  Dim hti As LVHITTESTINFO
  
  hti.pt.X = X \ Screen.TwipsPerPixelX                    'twips to pixels
  hti.pt.Y = Y \ Screen.TwipsPerPixelY
  Call SendMessage(LVW.hwnd, LVM_SUBITEMHITTEST, 0&, hti) 'get info from API
  Row = hti.lngItem + 1                                   'grab row
  Col = hti.lngSubItem                                    'grab column
End Sub


