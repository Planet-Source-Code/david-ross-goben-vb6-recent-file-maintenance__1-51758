Attribute VB_Name = "modActivatePreviousInstance"
Option Explicit
'~modActivatePreviousInstance.bas;
'Activate a previous instance of the running application
'******************************************************************************
' modActivatePreviousInstance: The ActivatePrevInstance() function will check
'                              to see if a previous instance of the current
'                              application is running. If so, it will activate
'                              it and terminate the current process.
'EXAMPLE:
'Private Sub Form_Load()       'main form (could also be Main() Sub
'  If App.PrevInstance Then    'previous exists?
'    ActivatePrevInstance      'yes, activate that one, unload this one
'    Exit Sub
'  End If
'End Sub
'******************************************************************************

Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const GW_HWNDPREV = 3

'******************************************************************************
' ActivatePrevInstance(): Activate previous instance, if found
'******************************************************************************
Public Sub ActivatePrevInstance()
  Dim OldTitle As String
  Dim PrevHndl As Long
  Dim Result As Long
'
' Save the title of the application.
'
  OldTitle = App.Title
'
' Rename the title of this application so FindWindow will not find this application instance.
'
  App.Title = "unwanted instance"
'
' Attempt to get window handle using no class name (check just title).
'
  PrevHndl = FindWindow(vbNullString, OldTitle)
'
' Check if no previous instance found.
'
  If PrevHndl = 0 Then
    App.Title = OldTitle                       'reset the app title
    Exit Sub
  End If
'
' Get handle to previous window.
'
  PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
'
' Restore the program.
'
  SetActiveWindow (PrevHndl)
  Call OpenIcon(PrevHndl)
'
' Activate the application.
'
  Call SetForegroundWindow(PrevHndl)
'
'End the application.
'
  End
End Sub

