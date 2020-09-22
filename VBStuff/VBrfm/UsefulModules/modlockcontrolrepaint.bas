Attribute VB_Name = "modLockControlRepaint"
Option Explicit
'~modLockControlRepaint.bas;
'Prevents paint refreshes from being performed on a control
'*******************************************************************************
' modLockControlRepaint: The LockControlRepaint() function prevents paint refreshes from being
'                        performed on a control. This can speed the operation of a control,
'                        such as a ListBox or TreeView by as much as 30%. Only one control at a
'                        time can be repaint locked.
'                        The UnlockControlRepaint() function turns off any locking on any control.
'*******************************************************************************

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'*******************************************************************************
' LockControlRepaint(): Prevent repaints on specified control
'*******************************************************************************
Public Sub LockControlRepaint(uControl As Control)
  On Error Resume Next
  LockWindowUpdate uControl.hwnd
  If Err.Number Then MsgBox uControl.Name & " does not have an hWnd value"
  On Error GoTo 0
End Sub

'*******************************************************************************
' UnlockControlRepaint(): turn off repaint locking
'*******************************************************************************
Public Sub UnlockControlRepaint(uControl As Control)
  On Error Resume Next
  LockWindowUpdate 0
  uControl.Refresh
  On Error GoTo 0
End Sub

