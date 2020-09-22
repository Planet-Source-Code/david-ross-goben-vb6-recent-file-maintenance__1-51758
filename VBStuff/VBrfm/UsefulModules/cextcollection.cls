VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExtCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'~cExtCollection.cls;
'Extended Collection class which allows Clear and Sort functionality
'***************************************************************************
' cExtCollection - This class is like a normal collection, except that it
'                  provides a Clear method to totally erase the collection,
'                  a Sort Property to allow alphabetic sorting, and a
'                  ReverseSort property to allow descending sorts. Sort and
'                  Reverse sort will toggle each other. Set either to FALSE
'                  to turn both options off. An Error property can be checked
'                  to see if the last operation succeeded. Save and Load from
'                  Files is supported, as well as loading, and merging to and
'                  from other ExtCollection, Collection, ListBox, and ComboBox
'                  objects.
' PROPERTIES:
' ByKey       :Read only. Like the Item property, but obtain the item by its key.
' Count       :Read only. Get number of items in the collection.
' Error       :Read only. Indicate if last command had an error.
' Item        :Read/Write. Get or assign an index item (index of 0=current)
' Key         :Read only. Returns the insertion order index for the nth item
'              in the collection. Not much use, but you can use it to grab
'              a list by order of insertion using the ByKey property.
' ListIndex   :Read/Write. Get or assign the current index in the collection.
' NewIndex    :Read only. Index of last added or 'Let' item assignment. This
'              is very useful when sorting is turned on.
' Sort        :Read/Write. Toggles ascending sorts. Assignment of TRUE or
'              FALSE will also turn off the ReverseSort property.
' ReverseSort :Read/write. Toggles descending sorts. Assignment of TRUE or
'              FALSE will also turn off the ascending Sort property.
'METHODS:
' Add         :Add an item to the end of the list. Update ListIndex to the
'              end of the list, and set NewIndex to the actual location of
'              the added item. Useful when sorting enabled. Also returns
'              TRUE if it succeeds.
' Append      :Append the contents of the collection to a Collection,
'              clsExtCollection, ListBox, or ComboBox. Return TRUE if success.
'              Duplicates will be ignored.
' Clear       :Reinitializes the collection to no data. Always success.
' Copy        :Copy contents of a Collection, clsExtCollection, ListBox,
'              or ComboBox to the current collection, replacing its contents.
'              Return TRUE if success.
' Load        :Load a file to the collection. The optional Merge parameter, if
'              set to TRUE, will merge loaded data to the current collection.
'              Duplicates will be ignored.
' Merge       :Merge contents of a Collection, clsExtCollection, ListBox,
'              or ComboBox to the current collection. Return TRUE if success.
'              Duplicates will be ignored.
' Paste       :Paste the collection contents to a Collection,
'              clsExtCollection, ListBox, or ComboBox. Return TRUE if success.
' Remove      :Removes the indexed item from the list. If Index=0, then the
'              current item pointed to by ListIndex is removed. Also returns
'              TRUE if it succeeds.
' Save        :Save the collection to a file. The OverWrite flag allows existing
'              files to be overwritten if set to TRUE (default is FALSE).
' Note: Any property that requires an index can generate
'       an error if the index is out of range (1 to Count).
'       Method that add or assign items can generate errors
'       if an assigned index is out of range, or you are
'       adding an item that already exists in the collection.
' Note: The Save and Load methods have the following error codes:
'       0 = Success
'       1 = File not found
'       2 = 'File' is a directory
'       3 = Cannot write to the file. Check permissions
'       4 = Unable to open the file
'       5 = Filename is blank
'       6 = Drive error
'       7 = Filepath already exists
'       8 = Filepath error
'EXAMPLE:
'  Dim MyCol As cExtCollection          'define collection object
'  Dim idx As Long                      'used index
'
'  Set MyCol = New cExtCollection       'create object
'  With MyCol
'    .Sort = True                       'enable sorting
'    .Add "Jeff"                        'add items to the list
'    .Add "Janet"
'    .Add "David"
'    .Add "Steve"
'    .Add "Scott"
'    .Add "Gary"
'    .Save "C:\Junk", True              'save the file. Overwite
'    .Clear                             'clear collection
'    .Load "C:\Junk"                    'reload
'    Me.Show                            'force form to screen
'    .ReverseSort = True                'reset descending
'    Me.Print "Items = " & CStr(.Count) 'show items in list
'    For idx = 1 To .Count              'display sorted list
'      Me.Print .Item(idx)
'    Next idx
'  End With
'  Set MyCol = Nothing                  'remove the object
'***************************************************************************

Private m_Col() As String         'the collection object
Private m_Key() As Long           'key indexes
Private m_Count As Long           'number of items in the collection
Private m_ListIndex As Long       'the current index into the collection
Private m_Sort As Boolean         'flag indicating if sorting turned on
Private m_ReverseSort As Boolean  'Indicating that sorting is to be reversed
Private m_Error As Boolean
Private m_NewIndex As Long        'actual insertion index

'*******************************************************************************
' Subroutine Name   : Class_Initialize
' Purpose           : Fire up teh class. Init all variables
'*******************************************************************************
Private Sub Class_Initialize()
  Call Clear
  m_Sort = False                  'ensure sorting turned off
  m_ReverseSort = False
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : Leaving. Make sure string array is cleared
'*******************************************************************************
Private Sub Class_Terminate()
  Erase m_Col
  Erase m_Key
End Sub

'*******************************************************************************
' Subroutine Name   : Clear
' Purpose           : Remove all items from the collection
'*******************************************************************************
Public Sub Clear()
  Erase m_Col         'remove collection
  Erase m_Key
  m_ListIndex = 0     'reset current index and count
  m_NewIndex = 0
  m_Count = 0
  m_Error = False     'success
End Sub

'*******************************************************************************
' Get Name          : Error
' Purpose           : get the boolean error flag. If true, then the last
'                     collection command failed
'*******************************************************************************
Public Property Get Error() As Boolean
  Error = m_Error
End Property

'*******************************************************************************
' Get Name          : Count
' Purpose           : Return the number of items in the collection
'*******************************************************************************
Public Property Get Count() As Long
  Count = m_Count
  m_Error = False     'show success
End Property

'*******************************************************************************
' Get Name          : Item
' Purpose           : Get an item from the list.
'*******************************************************************************
Public Property Get Item(Optional Index As Long = 0) As String
  Dim I As Long
  
  m_Error = True                                      'default for failure
  I = Index                                           'grab the index
  If I = 0 Then I = m_ListIndex                       'if 0, get current index
  If I < 1 Or I > m_Count Then Exit Property          'index out of range
  Item = m_Col(I)                                     'else grab item
  m_Error = False                                     'and indicate success
End Property

'*******************************************************************************
' Get Name          : Key
' Purpose           : Key the key for an item in the list
'*******************************************************************************
Public Property Get Key(Optional Index As Long = 0) As Long
  Dim I As Long
  
  m_Error = True                                      'default for failure
  I = Index                                           'grab the index
  If I = 0 Then I = m_ListIndex                       'if 0, get current index
  If I < 1 Or I > m_Count Then Exit Property          'index out of range
  Key = m_Col(I)                                     'else grab item
  m_Error = False                                     'and indicate success
End Property

'*******************************************************************************
' Get Name          : ByKey
' Purpose           : Return a collection item by its key
'*******************************************************************************
Public Property Get ByKey(Key As Long) As String
  Dim Idx As Long, K As Long
  
  m_Error = True                                      'default for failure
  m_ListIndex = 0
  If m_Count < 1 Or Key = 0 Then Exit Property        'out of range
  For Idx = 1 To m_Count
    If Key = m_Key(Idx) Then Exit For                 'check for key match
  Next Idx
  If Idx > m_Count Then Exit Property                 'did not find key
  m_ListIndex = Idx                                   'set ListIndex
  ByKey = m_Col(Idx)                                  'return text
  m_Error = False                                     'and indicate success
  
End Property

'*******************************************************************************
' Let Name          : Item
' Purpose           : Set an indexted item in the list.
'*******************************************************************************
Public Property Let Item(Index As Long, Text As String)
  Dim S As String
  Dim Idx As Long, I As Long
  
  m_Error = True                                      'default to fail
  I = Index                                           'grab the index
  If I = 0 Then I = m_ListIndex                       'if 0, get current index
  If I < 1 Or I > m_Count Then Exit Property          'index out of range
  
  S = UCase$(Trim$(Text))                             'grab text to stuff
  If Len(S) = 0 Then                                  'nothing to add?
    Call Remove(I)                                    'so remove
    Exit Property
  End If
'
' scan the list for a match
'
  For m_ListIndex = 1 To m_Count
    If S = UCase$(m_Col(Idx)) Then                    'match?
      If m_ListIndex = I Then Exit For                'yes; ok if same as index?
      Exit Property                                   'error otherwise
    End If
  Next m_ListIndex
'
' stuff item to list and sort as needed
'
  m_Col(I) = Trim$(Text)
  m_NewIndex = I                                      'shere item was added
  
  Call SortCol
  m_ListIndex = m_Count
  m_Error = False                                     'report all is ok
End Property

'*******************************************************************************
' Subroutine Name   : Add
' Purpose           : Add an item to the collection
'*******************************************************************************
Public Function Add(Text As String) As Boolean
  Dim S As String
  Dim Idx As Long, I As Long
  
  m_Error = True                                      'default to fail
  S = UCase$(Trim$(Text))                             'get text to add
  If Len(S) = 0 Then Exit Function                    'error if nothing
'
' see if item already exists in teh list
'
  For m_ListIndex = 1 To m_Count
    If S = UCase$(m_Col(Idx)) Then Exit Function      'match, so error
  Next m_ListIndex
  m_Count = m_Count + 1                               'else add an item
  ReDim Preserve m_Col(m_Count) As String             'bump the collection
  ReDim Preserve m_Key(m_Count) As Long
  m_Col(m_Count) = Trim$(Text)                        'stuff text
  m_Key(m_Count) = m_Count                            'keep original index
  m_ListIndex = m_Count                               'set list index
  m_NewIndex = m_Count                                'shere item was added
  Call SortCol                                        'sort as needed
  m_Error = False                                     'show success
  Add = True
End Function

'*******************************************************************************
' Get Name          : Listindex
' Purpose           : Get current index into the collection
'*******************************************************************************
Public Property Get ListIndex() As Long
  ListIndex = m_ListIndex
  m_Error = False                                     'show success
End Property

'*******************************************************************************
' Get Name          : NewIndex
' Purpose           : Get the last added item's index in teh collection. This is
'                   : especially useful when sorting is turned on
'*******************************************************************************
Public Property Get NewIndex() As Long
  NewIndex = m_NewIndex
  m_Error = False                                     'show success
End Property

'*******************************************************************************
' Get Name          : Sort
' Purpose           : Return the Sort Flag
'*******************************************************************************
Public Property Get Sort() As Boolean
  Sort = m_Sort                                       'return Sort flag
  m_Error = False                                     'show success
End Property

'*******************************************************************************
' Let Name          : Sort
' Purpose           : Assign the Sort flag. Turn off ReverseSort
'*******************************************************************************
Public Property Let Sort(Flag As Boolean)
  m_Sort = Flag                                       'set sort flag
  m_ReverseSort = False                               'turn off reverse sort
  Call SortCol                                        'sort as needed
  m_Error = False                                     'show success
End Property

'*******************************************************************************
' Get Name          : ReverseSort
' Purpose           : Return the ReverseSort state
'*******************************************************************************
Public Property Get ReverseSort() As Boolean
  ReverseSort = m_ReverseSort                         'return reverse sort flag
  m_Error = False                                     'show success
End Property

'*******************************************************************************
' Let Name          : ReverseSort
' Purpose           : Assign the ReverseSort state. Turn off the Sort flag
'*******************************************************************************
Public Property Let ReverseSort(Flag As Boolean)
  m_ReverseSort = Flag                                'set reverse sort flag
  m_Sort = False                                      'turn off ascending sort
  Call SortCol                                        'sort as needed
  m_Error = False                                     'show success
End Property

'*******************************************************************************
' Subroutine Name   : Remove
' Purpose           : Remove an item from the list
'*******************************************************************************
Public Function Remove(Optional Index As Long = 0) As Boolean
  Dim Idx As Long, I As Long
  
  m_Error = True                                      'default to fail
  I = Index
  If I = 0 Then I = m_ListIndex                       'if 0, use current index
  If I < 1 Or I > m_Count Then Exit Function          'error if out of range
  m_Count = m_Count - 1                               'drop one item
  If m_Count = 0 Then                                 'nothing left?
    Call Clear                                        'yes, so clear collection
    Exit Function                                     'and exit with success
  End If
'
' drop higher items down in collection
'
  For Idx = Index To m_Count
    m_Col(Idx) = m_Col(Idx + 1)
    m_Key(Idx) = m_Key(Idx + 1)
  Next Idx
  ReDim Preserve m_Col(m_Count) As String             'drop higher cell
  ReDim Preserve m_Key(m_Count) As Long
  m_ListIndex = Index                                 'set index to current location
  m_Error = False                                     'show success
  If m_NewIndex > m_Count Then m_NewIndex = 0         'out of range
  If m_NewIndex = Index Then m_NewIndex = 0           'was deleted
  Remove = True
End Function

'*******************************************************************************
' Subroutine Name   : SortCol
' Purpose           : Sort the collection as needed
'*******************************************************************************
Private Sub SortCol()
  Dim Idx As Long, colCount As Long, iTemp As Long
  Dim IndexLo As Long, IncIndex As Long
  Dim HalfUp As Long, IndexHi As Long
  Dim HalfDown As Long, NumberofItems As Long
  Dim S As String
  
  If Not m_Sort And Not m_ReverseSort Then Exit Sub
  If m_Count < 2 Then Exit Sub
  HalfDown = m_Count                    'number of items to sort
  Do While HalfDown \ 2                 'while counter can be halved
    HalfDown = HalfDown \ 2             'back down by 1/2
    HalfUp = m_Count - HalfDown         'look in upper half
    IncIndex = 1                        'init index to start of array
    Do While IncIndex <= HalfUp         'do while we can index range
      IndexLo = IncIndex                'set base
      Do
        IndexHi = IndexLo + HalfDown
        If UCase$(m_Col(IndexLo)) > UCase$(m_Col(IndexHi)) Then 'check strings
          S = m_Col(IndexLo)            'swap strings
          m_Col(IndexLo) = m_Col(IndexHi)
          m_Col(IndexHi) = S
          iTemp = m_Key(IndexLo)        'swap keys
          m_Key(IndexLo) = m_Key(IndexHi)
          m_Key(IndexHi) = iTemp
          If m_NewIndex = IndexHi Then
            m_NewIndex = IndexLo 'adjust Newindex
          ElseIf m_NewIndex = IndexLo Then
            m_NewIndex = IndexHi
          End If
          IndexLo = IndexLo - HalfDown  'back up index
        Else
          IncIndex = IncIndex + 1       'else bump counter
          IndexLo = 0                   'allow busting out of 2 loops
          Exit Do
        End If
      Loop While IndexLo > 0            'while more things to check
    Loop
  Loop
'
' check for storage in ascending or decending order
'
  If m_ReverseSort Then                   'DESCENDING
    IndexHi = m_Count
    IndexLo = m_Count \ 2
    For Idx = 1 To IndexLo                'flip sorted collection
      S = m_Col(Idx)
      m_Col(Idx) = m_Col(IndexHi)
      m_Col(IndexHi) = S
      IndexHi = IndexHi - 1
    Next Idx
  End If
End Sub

'*******************************************************************************
' Function Name     : Copy
' Purpose           : Copy data to collection if the source object is a
'                   : clsExtCollection, Collection, Listbox or ComboBox
'*******************************************************************************
Public Function Copy(Col As Object) As Boolean
  If TypeOf Col Is Collection Or TypeOf Col Is cExtCollection Or _
     TypeOf Col Is ListBox Or TypeOf Col Is ComboBox Then
    Clear
    Sort = False
  End If
  If Merge(Col) Then
    If Not TypeOf Col Is Collection Then
      If Col.Sort Then Sort = True
    End If
    
    If TypeOf Col Is cExtCollection Then
      If Col.ReverseSort Then ReverseSort = True
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : Merge
' Purpose           : Merge the contents of a source collection to the
'                   : current collection.
'*******************************************************************************
Public Function Merge(Col As Object) As Boolean
  Dim Idx As Long
  
  Merge = True
  If TypeOf Col Is Collection Then
    For Idx = 1 To Col.Count
      Add Col(Idx)
    Next Idx

  ElseIf TypeOf Col Is cExtCollection Then
    For Idx = 1 To Col.Count
      Add Col(Idx)
    Next Idx

  ElseIf TypeOf Col Is ListBox Or TypeOf Col Is ComboBox Then
    For Idx = 0 To Col.ListIndex - 1
      Add Col(Idx)
    Next Idx
  Else
    Merge = False
  End If
End Function

'*******************************************************************************
' Function Name     : Paste
' Purpose           : Paste colection to target Collection, clsExtCollection,
'                   : Listbox, or ComboBox
'*******************************************************************************
Public Function Paste(Col As Object) As Boolean
  Dim Idx As Long
  
  Paste = True                                'init to true
  If TypeOf Col Is Collection Then
    Do While Col.Count                        'erase collection contents
      Col.Remove 1
    Loop
    For Idx = 1 To m_Count                    'add current collection
      Col.Add m_Col(Idx)
    Next Idx

  ElseIf TypeOf Col Is cExtCollection Then
    Col.Clear                                 'clsExtCollection
    For Idx = 1 To m_Count
      Col.Add m_Col(Idx)
    Next Idx

  ElseIf TypeOf Col Is ListBox Or TypeOf Col Is ComboBox Then
    Col.Clear                                 'listbox or combobox
    For Idx = 1 To m_Count
      Col.AddItem m_Col(Idx)
    Next Idx
  Else
    Paste = False                             'else fail
  End If
End Function

'*******************************************************************************
' Function Name     : Append
' Purpose           : Append the collection contents to a Collection,
'                   : clsExtCollection, ListBox, or ComboBox.
'*******************************************************************************
Public Function Append(Col As Object) As Boolean
  Dim Idx As Long
  
  Append = True                               'init to true
  If TypeOf Col Is Collection Then
    On Error Resume Next                      'ignore duplication
    For Idx = 1 To m_Count                    'add current collection
      Col.Add m_Col(Idx)
    Next Idx

  ElseIf TypeOf Col Is cExtCollection Then
    For Idx = 1 To m_Count
      Col.Add m_Col(Idx)
    Next Idx

  ElseIf TypeOf Col Is ListBox Or TypeOf Col Is ComboBox Then
    For Idx = 1 To m_Count
      Col.AddItem m_Col(Idx)
    Next Idx
  Else
    Append = False                             'else fail
  End If
End Function

'*******************************************************************************
' Function Name     : Save
' Purpose           : Save the contents of the collection to a disk file
'*******************************************************************************
Public Function Save(FilePath As String, Optional OverWrite As Boolean = False) As Integer
  Dim FName As String
  Dim Fnum As Integer
  Dim Idx As Long
  
  Save = 5                                    'init filename is blank
  FName = Trim$(FilePath)                     'get filename
  If Len(FName) = 0 Then Exit Function        'nothing there
  Save = 8                                    'init to filepath error
  If InStr(1, FName, "*") Or InStr(1, FName, "?") Then Exit Function
  Save = 6                                    'init drive error
  On Error Resume Next                        'check for file and dirve error
  Idx = Len(Dir$(FName, vbDirectory))
  If Err.Number Then Exit Function            'drive error
  On Error GoTo 0
  If Idx Then                                 'file exists
    Save = 7                                  'init to file already exits
    If Not OverWrite Then Exit Function       'file already exists
    Idx = GetAttr(FName)                      'get type
    Save = 2                                  'init file is a directory
    If Idx And vbDirectory Then Exit Function 'directory? oops
    If Idx Then                               'other attributes?
      Save = 3                                'init permissions error
      On Error Resume Next                    'yes, prepare for disaster
      Call SetAttr(FName, vbNormal)           'try to make normal
      If Err.Number Then Exit Function        'no can do. Abandon ship
      On Error GoTo 0
    End If
  End If
'
' write the file
'
  Fnum = FreeFile                             'get a free file number
  Save = 4                                    'init unable to open
  On Error Resume Next                        'repare for danger
  Open FName For Output Access Write As #Fnum 'open the file
  If Err.Number Then Exit Function            'failed, so skip town
  For Idx = 1 To m_Count                      'else write contents
    Print #Fnum, m_Col(Idx)                   'keep error trap in case disk full
  Next Idx
  Close #Fnum                                 'close shop
  Save = 0                                    'indicate success
End Function

'*******************************************************************************
' Function Name     : Load
' Purpose           : Load a file to the collection
'*******************************************************************************
Public Function Load(FileName As String, Optional Merge As Boolean = False) As Integer
  Dim FName As String, S As String
  Dim Fnum As Integer
  Dim Idx As Long
  Dim SaveSort As Boolean, SaveReverseSort As Boolean
  
  Load = 5                                    'init filename is blank
  FName = Trim$(FileName)                     'get filename
  If Len(FName) = 0 Then Exit Function        'nothing there
  Load = 8                                    'init to filepath error
  If InStr(1, FName, "*") Or InStr(1, FName, "?") Then Exit Function
  Load = 6                                    'init drive error
  On Error Resume Next                        'check for file and dirve error
  Idx = Len(Dir$(FName, vbDirectory))
  If Err.Number Then Exit Function            'drive error
  On Error GoTo 0
  Load = 1                                    'init file does not exist
  If Idx = 0 Then Exit Function               'file did not exist
  Load = 2                                    'init file is a directory
  If GetAttr(FName) And vbDirectory Then Exit Function 'directory? oops
'
' write the file
'
  Fnum = FreeFile                             'get a free file number
  Load = 4                                    'init unable to open
  On Error Resume Next                        'repare for danger
  Open FName For Input Access Read As #Fnum 'open the file for reading
  If Err.Number Then Exit Function            'failed, so skip town
  If Not Merge Then Clear                     'clear out collection
  SaveSort = m_Sort                           'save sort flags
  SaveReverseSort = m_ReverseSort
  Sort = False                                'turn off sorting
  Do While Not EOF(Fnum)                      'read file straight in
    Line Input #Fnum, S                       'read a line
    Add S                                     'add line to file
  Loop
  Close #Fnum                                 'close shop
  If SaveSort Then Sort = True                'set sorting, if required
  If SaveReverseSort Then ReverseSort = True
  Load = 0                                    'indicate success
End Function
