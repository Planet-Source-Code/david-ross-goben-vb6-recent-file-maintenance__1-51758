VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVBPrj 
   Caption         =   "VB Recent File Maintenance"
   ClientHeight    =   5910
   ClientLeft      =   1140
   ClientTop       =   2640
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picControls 
      Height          =   975
      Index           =   1
      Left            =   8820
      ScaleHeight     =   915
      ScaleWidth      =   1815
      TabIndex        =   10
      Top             =   4440
      Width           =   1875
      Begin VB.CommandButton cmdUndo 
         Cancel          =   -1  'True
         Caption         =   "&Undo Changes, Exit"
         Height          =   375
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Cancel changes and exit"
         Top             =   0
         Width           =   1830
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes, Exit"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Save changes and exit"
         Top             =   540
         Width           =   1830
      End
   End
   Begin VB.PictureBox picControls 
      Height          =   3855
      Index           =   0
      Left            =   8760
      ScaleHeight     =   3795
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   180
      Width           =   1875
      Begin VB.CommandButton cmdExplore 
         Caption         =   "&Explore VBP Folder"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Explore the selected project's folder"
         Top             =   3480
         Width           =   1830
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove Entry"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Remove selected item(s) from the list"
         Top             =   0
         Width           =   1830
      End
      Begin VB.CommandButton cmdUnFound 
         Caption         =   "Remove Unfound &Items"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Remove items with no reachable path  (marked in RED)"
         Top             =   540
         Width           =   1830
      End
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "&Delete All"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Remove all items in the list"
         Top             =   1080
         Width           =   1830
      End
      Begin VB.CommandButton cmdReread 
         Caption         =   "Reread &List"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Reread the Recent File list"
         Top             =   2760
         Width           =   1830
      End
      Begin VB.Label Label2 
         Caption         =   "Note: Only entires are removed, not the actual project files."
         ForeColor       =   &H80000015&
         Height          =   675
         Left            =   0
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Note: Double-Click item to launch it in VB."
         ForeColor       =   &H80000015&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   2280
         Width           =   1755
      End
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "PName"
         Text            =   "Project Name"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Project File Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "lblWidth"
      Height          =   195
      Left            =   8820
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmVBPrj.frx":0000
      ForeColor       =   &H80000015&
      Height          =   435
      Left            =   180
      TabIndex        =   8
      Top             =   5460
      Width           =   8445
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExplore 
         Caption         =   "&Explore folder of selection"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReread 
         Caption         =   "&Re-read list"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndoExit 
         Caption         =   "&Undo changes && Exit"
      End
      Begin VB.Menu mnuSaveExit 
         Caption         =   "&Save changes && Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Removal"
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove &selected item(s)"
      End
      Begin VB.Menu mnuRemoveUnfound 
         Caption         =   "Remove &unfound items"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete all entries"
      End
   End
End
Attribute VB_Name = "frmVBPrj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim colNames As cExtCollection  'list of VBP file names
Dim colPaths As cExtCollection  'paths to the VBP files
Dim colIndex As cExtCollection  'track deleted items
Dim Unfound As Integer          'track entries that no longer actually exist
Dim Dragging As Boolean         'when dragging items
'
' registry path to the recent list
'
Private Const Path As String = "Software\Microsoft\Visual Basic\6.0\RecentFiles"

'*******************************************************************************
' Subroutine Name   : Form_Initialize
' Purpose           : Set up app for XP-style buttons, if supported
'*******************************************************************************
Private Sub Form_Initialize()
  Call FormInitialize
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Application startup
'*******************************************************************************
Private Sub Form_Load()
  Dim PrjList As Variant
'
' do not lauch if it is already active. Activate the previous instance instead
'
  If App.PrevInstance Then
    ActivatePrevInstance
    Exit Sub
  End If
  Me.Caption = Me.Caption & " " & GetAppVersion()
'
' create collections that we will use
'
  Set colNames = New cExtCollection
  Set colPaths = New cExtCollection
  Set colIndex = New cExtCollection
  Me.picControls(0).BorderStyle = 0   'remove framing about container objects
  Me.picControls(1).BorderStyle = 0
'
' read the recent list and adjust the column header widths
'
  Call ReadRecentList
  AdjustListViewColumns Me.lvwProjects, True
  Me.cmdExplore.Enabled = Me.lvwProjects.ListItems(Me.lvwProjects.SelectedItem.Index).ForeColor <> vbRed
  Me.mnuExplore.Enabled = Me.cmdExplore.Enabled
  Me.cmdRemove.Enabled = Me.cmdExplore.Enabled
  Me.mnuRemove.Enabled = Me.cmdRemove.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : Form_QueryUnload
' Purpose           : If exiting through the "X" button, unload all forms
'*******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Call cmdUndo_Click
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resizing the main form
'*******************************************************************************
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  If Me.Width < 11000 Then Me.Width = 11000
  If Me.Height < 6500 Then Me.Height = 6500
  
  With Me.picControls(0)
    .Left = Me.ScaleWidth - .Width - Me.lvwProjects.Left
  End With
  
  With Me.lvwProjects
    .Width = Me.picControls(0).Left - .Left * 2
    Me.lblInfo.Width = .Width
    Me.lblWidth.Caption = Me.lblInfo.Caption
    If Me.lblWidth.Width > .Width Then
      Me.lblInfo.Height = Me.lblWidth.Height * 2
    Else
      Me.lblInfo.Height = Me.lblWidth.Height
    End If
    Me.lblInfo.Top = Me.ScaleHeight - Me.lblInfo.Height - 60
    .Height = Me.lblInfo.Top - .Top - 60
  End With
  
  With Me.picControls(1)
    .Left = Me.picControls(0).Left
    .Top = Me.lblInfo.Top - .Height - 60
  End With
  
  AdjustListViewColumns Me.lvwProjects, True
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Remove created objects when leaving
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set colNames = Nothing
  Set colPaths = Nothing
  Set colIndex = Nothing
End Sub

'*******************************************************************************
' Subroutine Name   : cmdDeleteAll_Click
' Purpose           : Delete all entries in the VB6 recent file list
'                   : This will only remove entries from the displayed list
'                   : unless the user afterwards saves the changes
'*******************************************************************************
Private Sub cmdDeleteAll_Click()
  If MsgBox("Are you sure that you want to delete ALL entries?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Delete All Confirmation") = vbYes Then
    Call DeleteAll
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : DeleteAll
' Purpose           : Remove all entries from the displayed list.
'                   : Clean out the collections
'*******************************************************************************
Private Sub DeleteAll()
  Dim i As Integer
  
  LockControlRepaint Me.lvwProjects   'make faster by ignoring screen updates
  Me.lvwProjects.ListItems.Clear      'flush listview
  colPaths.Clear                      'flush collections
  colNames.Clear
  UnlockControlRepaint Me.lvwProjects 'refresh control
  Me.cmdDeleteAll.Enabled = False     'disable some buttons
  Me.mnuDelete.Enabled = False
  Me.cmdSave.Enabled = True
  Me.mnuSaveExit.Enabled = True
  Me.cmdRemove.Enabled = False
  Me.mnuRemove.Enabled = False
  Me.cmdUnFound.Enabled = False
  Me.cmdExplore.Enabled = False
  Me.mnuExplore.Enabled = False
  Me.mnuRemoveUnfound.Enabled = False
  Me.cmdReread.Enabled = True
  Me.mnuReread.Enabled = True
End Sub

'*******************************************************************************
' Subroutine Name   : cmdRemove_Click
' Purpose           : Remove selected entries
'*******************************************************************************
Private Sub cmdRemove_Click()
  Dim itm As ListItem
  Dim Idx As Long, i As Long
  Dim S As String
'
' scan backward through the list so that deletions do not hose up
' consecutive checks
'
  With Me.lvwProjects.ListItems
    For i = .Count To 1 Step -1
      Set itm = .Item(i)                  'get a treeview item
      If itm.Selected Then                'selected?
        If itm.ForeColor = vbRed Then Unfound = Unfound - 1
        S = itm.SubItems(1)               'get path
        For Idx = 1 To colPaths.Count     'find path in collection
          If S = colPaths.Item(Idx) Then  'remove from collections
            colPaths.Remove Idx
            colNames.Remove Idx
            Exit For
          End If
        Next Idx
        .Remove itm.Index                 'finally remove from the listview
      End If
    Next i
  End With
'
' cleanup
'
  Me.cmdRemove.Enabled = False
  Me.mnuRemove.Enabled = False
  Me.cmdUnFound.Enabled = CBool(Unfound)
  Me.mnuRemoveUnfound.Enabled = Me.cmdUnFound.Enabled
  Me.cmdSave.Enabled = True
  Me.mnuSaveExit.Enabled = True
  Me.cmdDeleteAll.Enabled = CBool(colNames.Count)
  Me.mnuDelete.Enabled = Me.cmdDeleteAll.Enabled
  Me.cmdReread.Enabled = Me.cmdDeleteAll.Enabled
  Me.mnuReread.Enabled = Me.cmdDeleteAll.Enabled
  Me.cmdExplore.Enabled = Me.cmdRemove.Enabled
  Me.mnuExplore.Enabled = Me.cmdExplore.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : cmdReread_Click
' Purpose           : Re-read the recent list from the registry
'*******************************************************************************
Private Sub cmdReread_Click()
  Call DeleteAll                  'clear listview out
  Call ReadRecentList             'refresh the list
End Sub

'*******************************************************************************
' Subroutine Name   : cmdExplore_Click
' Author            : David Goben
' Purpose           : Explore the selected VBP file's folder
'*******************************************************************************
Private Sub cmdExplore_Click()
  Dim Path As String
  With Me.lvwProjects
    Path = .SelectedItem.SubItems(1)            'get the full path to the VBP file
    Path = Left$(Path, InStrRev(Path, "\") - 1) 'get the folder that contains it
  End With
  BrowsePath Me.hwnd, Path                      'browse that folder
End Sub

'*******************************************************************************
' Subroutine Name   : cmdSave_Click
' Purpose           : Save the changes to the registry
'*******************************************************************************
Private Sub cmdSave_Click()
  Dim Idx As Long
  Dim S As String
  Dim Myreg As cReadWriteEasyReg
'
' flush unfound items from the registry
'
  If colIndex.Count Then
    Set Myreg = New cReadWriteEasyReg
    Call Myreg.OpenRegistry(HKEY_CURRENT_USER, Path)
    Do While colIndex.Count                 'delete all items tagged for deletion
      Myreg.DeleteValue colIndex.Item(1)
      colIndex.Remove 1
    Loop
'
' now save the updated list to the registry
'
    With Me.lvwProjects.ListItems
      For Idx = 1 To .Count
        Myreg.CreateValue CStr(Idx), .Item(Idx).SubItems(1), REG_SZ
      Next Idx
    End With
'
' flush the collections
'
    colPaths.Clear
    colNames.Clear
    
    Myreg.CloseRegistry   'close and remove registry object
    Set Myreg = Nothing
  End If
  
  Unload frmName
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUndo_Click
' Purpose           : Throw away changes and leave
'*******************************************************************************
Private Sub cmdUndo_Click()
  Unload frmName
  Unload Me
End Sub

'*******************************************************************************
' Subroutine Name   : cmdUnFound_Click
' Purpose           : Remove entries whose VBP files no longer exist
'*******************************************************************************
Private Sub cmdUnFound_Click()
  Dim itm As ListItem
  Dim S As String
  Dim Idx As Long, i As Integer
  
  With Me.lvwProjects.ListItems
    For i = .Count To 1 Step -1           'scan the listview
      Set itm = .Item(i)                  'grab an item
      If itm.ForeColor = vbRed Then       'unfound items are tagged with RED
        S = itm.SubItems(1)               'get the path
        For Idx = 1 To colPaths.Count     'flush from collections
          If S = colPaths.Item(Idx) Then
            colPaths.Remove Idx
            colNames.Remove Idx
            Exit For
          End If
        Next Idx
        .Remove itm.Index                 'then remove from the listview
      End If
    Next i
  End With
'
' cleanup
'
  Unfound = 0
  Me.cmdRemove.Enabled = False
  Me.mnuRemove.Enabled = False
  Me.cmdUnFound.Enabled = False
  Me.mnuRemoveUnfound.Enabled = False
  Me.cmdSave.Enabled = True
  Me.mnuSaveExit.Enabled = True
  Me.cmdDeleteAll.Enabled = CBool(colNames.Count)
  Me.mnuDelete.Enabled = Me.cmdDeleteAll.Enabled
  Me.cmdReread.Enabled = Me.cmdDeleteAll.Enabled
  Me.mnuReread.Enabled = Me.cmdDeleteAll.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : ReadRecentList
' Purpose           : Read the VB recent files list from the registry
'*******************************************************************************
Private Sub ReadRecentList()
  Dim Myreg As cReadWriteEasyReg
  Dim Idx As Long, i As Long
  Dim sIdx As String, sPath As String, sName As String, S As String
  Dim itm As ListItem
  Dim PrjList As Variant
'
' flush collections
'
  colNames.Clear
  colPaths.Clear
  colIndex.Clear
  
  Unfound = 0       'init nothing unfound
'
' scan the registry for entries
'
  Set Myreg = New cReadWriteEasyReg
  If Myreg.OpenRegistry(HKEY_CURRENT_USER, Path) Then
    PrjList = Myreg.GetAllValues                    'grab recent list to variant array
    If Not IsNull(PrjList) Then                     'if something grabbed
      For Idx = LBound(PrjList) To UBound(PrjList)  'process all entries
        sIdx = PrjList(Idx)                         'get the index number
        sPath = Myreg.GetValue(sIdx)
        sName = Mid$(sPath, InStrRev(sPath, "\") + 1)
        colIndex.Add sIdx                           'add to index
        colPaths.Add sPath                          'add path to VBP file
        colNames.Add sName                          'add VBP file name
        On Error GoTo 0
      Next Idx
    End If
    Myreg.CloseRegistry                             'done with object
    
    LockControlRepaint Me.lvwProjects               'kill screen flashes
    With Me.lvwProjects.ListItems
      .Clear                                        'init listview
      For Idx = 1 To colPaths.Count                 'add all paths
        sPath = colPaths.Item(Idx)                  'grab the path
        Set itm = .Add(, , colNames.Item(Idx))      'add filename to listview
        itm.SubItems(1) = sPath                     'and its path
        On Error Resume Next
        i = Len(Dir$(sPath))                        'see if VBP file exists
        If Err.Number > 0 Or i = 0 Then             'NOPE
          itm.ForeColor = vbRed                     'mark as RED if not found
          Unfound = Unfound + 1                     'mark unfound count
        End If
        On Error GoTo 0
      Next Idx
      Me.cmdRemove.Enabled = CBool(.Count)
      Me.mnuRemove.Enabled = Me.cmdRemove.Enabled
      Me.cmdDeleteAll.Enabled = CBool(.Count)
      Me.mnuDelete.Enabled = Me.cmdDeleteAll.Enabled
    End With
    UnlockControlRepaint Me.lvwProjects
    Me.cmdUnFound.Enabled = CBool(Unfound)          'enable button if items not found
    Me.mnuRemoveUnfound.Enabled = Me.cmdUnFound.Enabled
    Me.cmdSave.Enabled = False                      'nothing to save yet
    Me.mnuSaveExit.Enabled = False
    Me.cmdRemove.Enabled = False
    Me.mnuRemove.Enabled = False
    Me.cmdExplore.Enabled = False
    Me.mnuExplore.Enabled = False
    Me.cmdReread.Enabled = False
    Me.mnuReread.Enabled = False
  End If
  Set Myreg = Nothing
'
' adjust column headers to contents and width
'
  AdjustListViewColumns Me.lvwProjects, True
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_Click
' Purpose           : Enable remove button when something selected
'*******************************************************************************
Private Sub lvwProjects_Click()
  Me.cmdRemove.Enabled = True
  Me.mnuRemove.Enabled = True
  With Me.lvwProjects
    Me.cmdExplore.Enabled = .ListItems(.SelectedItem.Index).ForeColor <> vbRed
    Me.mnuExplore.Enabled = Me.cmdExplore.Enabled
    Me.cmdRemove.Enabled = CBool(.SelectedItem.Index)
    Me.mnuRemove.Enabled = Me.cmdRemove.Enabled
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_ColumnClick
' Purpose           : Sort times based upon columns clicked
'*******************************************************************************
Private Sub lvwProjects_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With Me.lvwProjects
    .SortKey = ColumnHeader.Index - 1   'indicate which column we want sorting
    If Not .Sorted Then                 'if not yet sorted...
      .Sorted = True                    'turn on sorting
      .SortOrder = lvwAscending         'init to ascending
    Else
      If .SortOrder = lvwAscending Then 'if previously ascending...
        .SortOrder = lvwDescending      'then mark as descending
      Else
        .SortOrder = lvwAscending       'else it had ben descending, so ascend
      End If
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_DblClick
' Purpose           : Launch a project in VB if it is doubleclicked
'*******************************************************************************
Private Sub lvwProjects_DblClick()
  Dim itm As ListItem
  Dim sFile As String, sVB As String
  
  For Each itm In Me.lvwProjects.ListItems
    If itm.Selected Then                        'only one item will be selected
      If itm.ForeColor = vbRed Then             'if red, then VBP file not there
        MsgBox itm.Text & " no longer exists. Cannot open it.", _
               vbOKOnly Or vbExclamation, "VBP File Not Found"
        Exit Sub
      End If
      sFile = itm.SubItems(1)                   'get path
      Select Case FindAssociatedExecutable(sFile, sVB) 'find VB6 executable
        Case 0                                  'found exe
          Shell sVB & " " & sFile, vbNormalFocus
        Case Else                               'did not find it (uninstalled?)
          MsgBox "VBP files do not seem to be associated with VB.", _
                 vbOKOnly Or vbExclamation, "Cannot Launch VB"
          Exit Sub
      End Select
      Exit For
    End If
  Next itm
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_MouseDown
' Purpose           : Initiate dragging, in case that is user intent
'*******************************************************************************
Private Sub lvwProjects_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton And Shift = 0 Then
    With Me.lvwProjects
      If Not .SelectedItem Is Nothing Then .OLEDrag 'turn on dragging
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_OLEDragDrop
' Purpose           : Mouse released
'*******************************************************************************
Private Sub lvwProjects_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Idx As Integer, lvRow As Integer
  Dim itm As MSComctlLib.ListItem
  Dim sItm As String, itmPath As String
  
  Me.lvwProjects.OLEDragMode = ccOLEDragManual
  Call ListViewMouseUp(Me.lvwProjects, X, Y, lvRow, Idx)  'ger row dropped on
'
' if a row is actually dropped on, then shift things around
'
  If lvRow Then
    With Me.lvwProjects.ListItems
      For Idx = .Count To 1 Step -1
        Set itm = .Item(Idx)
        If itm.Selected Then
          sItm = itm.Text
          itmPath = itm.SubItems(1)
          .Remove Idx
          Set itm = .Add(lvRow, , sItm)
          itm.SubItems(1) = itmPath
        End If
      Next Idx
      .Item(lvRow).Selected = True
      Me.cmdSave.Enabled = True
      Me.mnuSaveExit.Enabled = True
    End With
  End If
  frmName.Visible = False         'turn off dragging "icon"
  Me.cmdRemove.Enabled = CBool(Me.lvwProjects.SelectedItem.Index)
  Me.mnuRemove.Enabled = Me.cmdRemove.Enabled
  Me.cmdExplore.Enabled = Me.cmdRemove.Enabled
  Me.mnuExplore.Enabled = Me.cmdExplore.Enabled
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_OLEStartDrag
' Purpose           : We are initialiting a drag
'*******************************************************************************
Private Sub lvwProjects_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  Dragging = True
  
  With frmName
    .lblName.Caption = Me.lvwProjects.SelectedItem.SubItems(1)  'stuff filepath
    .Width = .lblName.Width + 40
    .Height = .lblName.Height + 20
    Me.lvwProjects.OLEDragMode = ccOLEDragAutomatic 'turn on dragging control
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : lvwProjects_OLEDragOver
' Purpose           : If the user is actually dragging, show dragging "icon"
'*******************************************************************************
Private Sub lvwProjects_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
  With frmName
    .Top = Me.Top + Me.lvwProjects.Top + Y + 640
    .Left = Me.Left + Me.lvwProjects.Left + X + 400
    If Not .Visible Then .Visible = True            'display "icon" if not vis
  End With
End Sub

'*******************************************************************************
' Let menu options fire button stuff
'*******************************************************************************
Private Sub mnuDelete_Click()
  Me.cmdDeleteAll.Value = True
End Sub

Private Sub mnuExplore_Click()
  Me.cmdExplore.Value = True
End Sub

Private Sub mnuRemove_Click()
  Me.cmdRemove.Value = True
End Sub

Private Sub mnuRemoveUnfound_Click()
  Me.cmdUnFound.Value = True
End Sub

Private Sub mnuReread_Click()
  Me.cmdReread.Value = True
End Sub

Private Sub mnuSaveExit_Click()
  Me.cmdSave.Value = True
End Sub

Private Sub mnuUndoExit_Click()
  Me.cmdUndo.Value = True
End Sub
