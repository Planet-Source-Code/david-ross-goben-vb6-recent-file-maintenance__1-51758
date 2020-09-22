VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReadWriteEasyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~cReadWriteEasyReg.cls;
'An easy way to READ/WRITE and DELETE from the registry
'***************************************************************
' cReadWriteEasyReg
'
'Description:An easy way to READ/WRITE and DELETE from the registry.
'            I've designed this class to be as easy as posible to use.
'            To read write and delete from the registry is actually easy
'            (if you know how). With this code you can be a pro in the regestry
'            without realy working with the nitty-gritty of the API Call's etc.
'
' Also include the following lines into your code:
'   Dim MyReg As cReadWriteEasyReg
'   Set MyReg = New cReadWriteEasyReg
'
' Remove the object using:
'   Set MyReg = Nothing
'
' This class supports the following methods:
'GetAllSubDirectories(): This function will return an array of variant with all the subkey values.
'CreateDirectory():      This function will return a true or false when it creates a key for you.
'DeleteDirectory():      This function will return a true or false when it deletes a key for you.
'GetAllValues():         This function will return an array of variant with all the value names in a key.
'CreateValue():          This function will return a true or false when it creates a value for you.
'DeleteValue():          This function will return a true or false when it deletes a value for you.
'GetValue():             This function will return a specific value from the registry. To get the
'                        "(Default)" value, provide a blank parameter: GetValue(vbnullstring).
'Get RegistryRootKey():  This property returns the current KeyValue
'Get SubDirectory():     This property returns the current 'Registry Directory' you are in
'OpenRegistry():         This function open's the registry at a specific Registry Directory.
'CloseRegistry():        This function should be called after you're done with the registry
'SortArrayAscending():   Sort a variable array into ascending order
'GetHKey():              This function returns the current HKEY value for the current registry directory
'
' NOTE: For SMPIS usage, see "modSMPISRegisteryIO" for even easier routines
'       to process data without worrying about a tree structure, using this
'       class as a dependency file
'------------------------------------------------------------------------
' Example: to obtain a list of printers on your system. We will get a list
' of Directory Names, and then debug print each printer's assigned name:
'------------------------------------------------------------------------
'  Dim MyReg As cReadWriteEasyReg
'  Dim PrtList As Variant, S As String, Idx As Integer, Path As String
'
'  Set MyReg = New cReadWriteEasyReg
'  Path = "System\CurrentControlSet\Control\Print\Printers"
'  If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Path) Then
'    MsgBox "Couldn't open the registry"
'  Else
'    PrtList = MyReg.GetAllSubDirectories   'get all directory names in 'Printers'
'    MyReg.CloseRegistry                    'close the registry
'    If Not IsNull(PrtList) Then            'do if list contains data
'      MyReg.SortArrayAscending (PrtList)   'Sort the list if you want to
'      For Idx = LBound(PrtList) To UBound(PrtList)
'        S = PrtList(Idx)                   'get a directory name
'        Call MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, Path & "\" & S)
'        Debug.Print MyReg.GetValue("Name") 'obtain value from the Name key
'        MyReg.CloseRegistry                'close the registry
'      Next Idx
'    End If
'  End If
'  Set MyReg = Nothing
'***************************************************************
Option Explicit

'***************************************************************
'Windows API/Global Declarations
'***************************************************************
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpKeyName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "advapi32.dll" (lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExL Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long

Private Type SECURITY_ATTRIBUTES
  nLength              As Long
  lpSecurityDescriptor As Variant
  bInheritHandle       As Long
End Type
'
'Enum's for the OpenRegistry function
'
Public Enum HKeys
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum
'
'Enum's for the DataTypes
'
Public Enum lDataType
  REG_NONE = 0
  REG_SZ = 1
  REG_EXPAND_SZ = 2
  REG_BINARY = 3
  REG_DWORD = 4
  REG_DWORD_LITTLE_ENDIAN = 4
  REG_DWORD_BIG_ENDIAN = 5
  REG_LINK = 6
  REG_MULTI_SZ = 7
  REG_RESOURCE_LIST = 8
  REG_FULL_RESOURCE_DESCRIPTOR = 9
  REG_RESOURCE_REQUIREMENTS_LIST = 10
  REG_QWORD = 11
End Enum
'
'Right's for the OpenRegistry
'
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = (KEY_READ And KEY_WRITE)
Private Const KEY_ALL_ACCESS = (( _
  STANDARD_RIGHTS_ALL Or _
  KEY_QUERY_VALUE Or _
  KEY_SET_VALUE Or _
  KEY_CREATE_SUB_KEY Or _
  KEY_ENUMERATE_SUB_KEYS Or _
  KEY_NOTIFY Or _
  KEY_CREATE_LINK _
) And (Not SYNCHRONIZE))
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1
'
'Local var's to keep track of things happening
'
Dim RootHKey As HKeys
Dim SubDir As String
Dim hKey As Long
Dim OpenRegOk As Boolean
Dim m_value_Type As lDataType

'*************************************************************************
' GetHKey(): This function returns the current HKEY value for the current registry directory
'*************************************************************************
Function GetHKey() As Long
  GetHKey = hKey
End Function

'*************************************************************************
'GetAllSubDirectories(): This function will return a array of variant with all the subkey values.
'*************************************************************************
'EXAMPLE:
'Dim MyVariant As Variant, MyReg As CReadWriteEasyReg, i As Integer
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'MyVariant = MyReg.GetAllSubDirectories
'For i = LBound(MyVariant) To UBound(MyVariant)
'  Debug.Print MyVariant(i)
'Next i
'MyReg.CloseRegistry
'Set MyReg = Nothing
'
Function GetAllSubDirectories() As Variant
  On Error GoTo handelgetdirvalues
  Dim SubKey_Num As Integer
  Dim SubKey_Name As String
  Dim Length As Long
  Dim ReturnArray() As Variant
    
  GetAllSubDirectories = Null
  If Not OpenRegOk Then Exit Function
  'Get the Dir List
  SubKey_Num = 0
  Do
    Length = 256
    SubKey_Name = Space$(Length)
    If RegEnumKey(hKey, SubKey_Num, SubKey_Name, Length) <> 0 Then
      Exit Do
    End If
    
    SubKey_Name = Left$(SubKey_Name, InStr(SubKey_Name, vbNullChar) - 1)

    ReDim Preserve ReturnArray(SubKey_Num) As Variant
    ReturnArray(SubKey_Num) = SubKey_Name
    SubKey_Num = SubKey_Num + 1
  Loop

  If SubKey_Num Then GetAllSubDirectories = ReturnArray
  Exit Function

handelgetdirvalues:
  GetAllSubDirectories = Null
End Function

'*************************************************************************
'CreateDirectory(): This function will return a true or false when it creates a key for you.
'*************************************************************************
'EXAMPLE:
'Dim MyReg As CReadWriteEasyReg
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'if MyReg.CreateDirectory("TestDir") then
'  Msgbox "Key created"
'else
'  msgbox "Couldn't Create key"
'end if
'MyReg.CloseRegistry
'Set MyReg = Nothing
'
Public Function CreateDirectory(ByVal sNewDirName As String) As Boolean
  Dim hNewKey As Long, lpdwDisposition As Long
  Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    
  If Not OpenRegOk Then Exit Function
  If RegCreateKeyEx(hKey, sNewDirName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, lpdwDisposition) = 0 Then
    CreateDirectory = True
  End If
End Function

'*************************************************************************
'DeleteDirectory(): This function will return a true or false when it deletes a key for you.
'*************************************************************************
'EXAMPLE:
'Dim MyReg As CReadWriteEasyReg
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'if MyReg.DeleteDirectory("MyTestDir") then
'  Msgbox "Key Deleted"
'else
'  msgbox "Couldn't Delete key"
'end if
'MyReg.CloseRegistry
'Set MyReg = Nothing
'
Public Function DeleteDirectory(ByVal sKeyName As String) As Boolean
  If Not OpenRegOk Then Exit Function
  If RegDeleteKey(hKey, sKeyName) = 0 Then DeleteDirectory = True
End Function

'*************************************************************************
'GetAllValues(): This function will return a array of variant with all the value names in a key.
'*************************************************************************
'EXAMPLE:
'Dim MyVariant As Variant, MyReg As CReadWriteEasyReg, i As Integer
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "HardWare\Description\System\CentralProcessor\0") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'MyVariant = MyReg.GetAllValues
'For i = LBound(MyVariant) To UBound(MyVariant)
'  Debug.Print MyVariant(i)
'Next i
'MyReg.CloseRegistry
'Set MyReg = Nothing
'
Function GetAllValues() As Variant

  On Error GoTo handelgetdirvalues
  Dim lpData As String, KeyType As Long
  Dim BufferLengh As Long, vname As String, vnamel As Long
  Dim ReturnArray() As Variant, Index As Integer
    
  GetAllValues = Null
  If Not OpenRegOk Then Exit Function
    
  'Get the Values List
  Index = 0
  Do
    lpData = String(250, " ")
    BufferLengh = 240
    vname = String(250, " ")
    vnamel = 240

    If RegEnumValue(ByVal hKey, ByVal Index, vname, vnamel, 0, KeyType, lpData, BufferLengh) <> 0 Then
      Exit Do
    End If

    vname = Left$(vname, InStr(vname, vbNullChar) - 1)
    ReDim Preserve ReturnArray(Index) As Variant
    ReturnArray(Index) = vname
    Index = Index + 1
  Loop

  GetAllValues = ReturnArray
  Exit Function

handelgetdirvalues:
  GetAllValues = Null
End Function

'*************************************************************************
'CreateValue(): This function will return a true or false when it creates a value for you.
'*************************************************************************
'EXAMPLE:
'Dim MyReg As CReadWriteEasyReg
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'If MyReg.CreateValue("ValName", "This is written as the value", REG_SZ) then
'  Msgbox "Value created"
'else
'  msgbox "Couldn't Create Value"
'end if
'MyReg.CloseRegistry
'Set MyReg = Nothing

Public Function CreateValue(ByVal sValueName As String, ByVal vWriteThis As Variant, ldValueDataType As lDataType, Optional Multi_SZ_AddtlStrings As Variant) As Boolean
  Dim lpData As String 'The pointer to the value written to the Registry key's value
  Dim cbData As Long 'The size of the data written to the Registry key's value, including termination characters If applicable
  Dim lReturn As Long 'The Error value returned by the Registry Function
  Dim str As Variant
  
  If Not OpenRegOk Then Exit Function
  
  Select Case ldValueDataType
    Case REG_SZ, REG_EXPAND_SZ
      lpData = vWriteThis & vbNullChar
      cbData = Len(lpData) - 1
      If RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData) Then
        CreateValue = False
      Else
        CreateValue = True
      End If

    Case REG_MULTI_SZ
      lpData = vWriteThis & vbNullChar
      If Not IsMissing(Multi_SZ_AddtlStrings) Then
        If IsArray(Multi_SZ_AddtlStrings) Then
          For Each str In Multi_SZ_AddtlStrings
            If Len(str) > 0 And str <> vbNullChar And Not IsNull(str) Then
              lpData = lpData & str & vbNullChar
            End If
          Next str
        Else
          If Len(Multi_SZ_AddtlStrings) > 0 And Multi_SZ_AddtlStrings <> vbNullChar And Not IsNull(Multi_SZ_AddtlStrings) Then
            lpData = lpData & Multi_SZ_AddtlStrings & vbNullChar
          End If
        End If
      End If

      lpData = lpData & vbNullChar
      cbData = Len(lpData) - 1
      If RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData) Then
        CreateValue = False
      Else
        CreateValue = True
      End If

    Case REG_DWORD
      Dim lpLong As Long
      lpLong = CLng(vWriteThis)
      cbData = 4
      If RegSetValueExL(hKey, sValueName, 0&, ldValueDataType, lpLong, cbData) Then
        CreateValue = False
      Else
        CreateValue = True
      End If

    Case Else
      MsgBox "Create Value: Unable to process that type of data."
      CreateValue = False
  End Select
End Function

'*************************************************************************
'DeleteValue(): This function will return a true or false when it deletes a value for you.
'*************************************************************************
'EXAMPLE:
'Dim MyReg As CReadWriteEasyReg
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'if MyReg.DeleteValue("ValName") then
'  Msgbox "Value Deleted"
'else
'  msgbox "Couldn't Delete Value"
'end if
'MyReg.CloseRegistry
'Set MyReg = Nothing

Public Function DeleteValue(ByVal sValueName As String) As Boolean
  If Not OpenRegOk Then Exit Function
  If RegDeleteValue(hKey, sValueName) = 0 Then DeleteValue = True
End Function

'*************************************************************************
'GetValue(): This function will return a specific value from the registry.
'*************************************************************************
'EXAMPLE:
'Dim MyString As String, MyReg As CReadWriteEasyReg, i As Integer
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, "HardWare\Description\System\CentralProcessor\0") Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'MyString = MyReg.GetValue("Identifier")
'Debug.Print MyString
'MyReg.CloseRegistry
'Set MyReg = Nothing
Function GetLastType() As lDataType
  GetLastType = m_value_Type
End Function

Function GetValue(ByVal VarName As String, Optional ReturnBinStr As Boolean = False) As Variant
  On Error GoTo handelgetavalue
  Dim I As Integer
  Dim SubKey_Value As String, TempStr As String, ReturnArray() As Variant
  Dim Length As Long
  Dim RtnVal As Long, value_Type As lDataType
    
  If Not OpenRegOk Then Exit Function
    
  'Read the size of the value value
  RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal 0&, Length)
  m_value_Type = value_Type
  
  Select Case RtnVal
    Case 0  'Ok so continue
    Case 2  'Not Found
            Exit Function
    Case 5  'Access Denied
            GetValue = "Access Denied"
            Exit Function
    Case Else 'What?
            GetValue = "RegQueryValueEx Returned : (" & RtnVal & ")"
            Exit Function
  End Select

'declare the size of the value and read it

  SubKey_Value = Space$(Length)

  RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal SubKey_Value, Length)

  Select Case value_Type
    Case REG_NONE
        'Not defined
        SubKey_Value = "Not defined value_type=REG_NONE"
    Case REG_SZ 'A null-terminated String
        SubKey_Value = Left$(SubKey_Value, Length)  'get data, but check for null
        If Right$(SubKey_Value, 1) = Chr$(0) Then SubKey_Value = Left$(SubKey_Value, Length - 1)
    Case REG_EXPAND_SZ
    'A null-terminated string that contains unexpanded references to
    'environment variables (for example, "%PATH%").
    'Use ExpandEnvironmentStrings to expand
        SubKey_Value = Left$(SubKey_Value, Length - 1)
    Case REG_BINARY 'Binary data in any form.
        SubKey_Value = Left$(SubKey_Value, Length)
        If Not ReturnBinStr Then
          TempStr = vbNullString
          For I = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, I, 1)))), 2) & " "
          Next I
          SubKey_Value = TempStr
        End If
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN 'A 32-bit number.
        SubKey_Value = Left$(SubKey_Value, Length)
        If Not ReturnBinStr Then
          TempStr = vbNullString
          For I = Len(SubKey_Value) To 1 Step -1
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, I, 1)))), 2)
          Next I
          SubKey_Value = TempStr
        End If
    Case REG_DWORD_BIG_ENDIAN
      'A 32-bit number in big-endian format.
      'In big-endian format, a multi-byte value is stored in memory from
      'the highest byte (the "big end") to the lowest byte. For example,
      'the value 0x12345678 is stored as (0x12 0x34 0x56 0x78) in big-endian format.
    Case REG_LINK
      'A Unicode symbolic link. Used internally; applications should not
      'use this type.
      SubKey_Value = "Not defined value_type=REG_LINK"
    Case REG_MULTI_SZ
      'Array of null-terminated string
      SubKey_Value = Left$(SubKey_Value, Length)
      I = 0
      While Len(SubKey_Value) > 0
        ReDim Preserve ReturnArray(I) As Variant
        ReturnArray(I) = Mid$(SubKey_Value, 1, InStr(1, SubKey_Value, vbNullChar) - 1)
        SubKey_Value = Mid$(SubKey_Value, InStr(1, SubKey_Value, vbNullChar) + 1)
        I = I + 1
      Wend
      GetValue = ReturnArray
      Exit Function
    Case REG_RESOURCE_LIST
      'Device driver resource list.
      SubKey_Value = "Not defined value_type=REG_RESOURCE_LIST"
    Case REG_FULL_RESOURCE_DESCRIPTOR
      'Device driver resource list.
      SubKey_Value = "Not defined value_type=REG_FULL_RESOURCE_DESCRIPTOR"
    Case REG_RESOURCE_REQUIREMENTS_LIST
      'Device driver resource list.
      SubKey_Value = "Not defined value_type=REG_RESOURCE_REQUIREMENTS_LIST"
    Case REG_QWORD
      SubKey_Value = Left$(SubKey_Value, Length)
      If Not ReturnBinStr Then
        TempStr = vbNullString
        For I = Len(SubKey_Value) To 1 Step -1
          TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, I, 1)))), 2)
        Next I
        SubKey_Value = TempStr
      End If
    Case Else
      SubKey_Value = "value_type=" & value_Type
  End Select

  GetValue = SubKey_Value
  Exit Function

handelgetavalue:
  GetValue = vbNullString
End Function

'*************************************************************************
'Get RegistryRootKey(): This property returns the current KeyValue
'*************************************************************************
Public Property Get RegistryRootKey() As HKeys
  RegistryRootKey = RootHKey
End Property

'*************************************************************************
'Get SubDirectory(): This property returns the current 'Registry Directory' you are in
'*************************************************************************
Public Property Get SubDirectory() As String
  SubDirectory = SubDir
End Property

'*************************************************************************
'OpenRegistry(): This function open's the registry at a specific 'Registry Directory.
'*************************************************************************
'EXAMPLE:
'Dim MyVariant As Variant, MyReg As CReadWriteEasyReg, i As Integer
'Set MyReg = New cReadWriteEasyReg
'If Not MyReg.OpenRegistry(HKEY_LOCAL_MACHINE, vbnullstring) Then
'  MsgBox "Couldn't open the registry"
'  Exit Sub
'End If
'MyVariant = MyReg.GetAllSubDirectories
'For i = LBound(MyVariant) To UBound(MyVariant)
'  Debug.Print MyVariant(i)
'Next i
'MyReg.CloseRegistry
'Set MyReg = Nothing
Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Boolean
  On Error GoTo OpenReg
  Dim ReturnVal As Integer
'
' close any currently open key
'
  If OpenRegOk Then
    CloseRegistry
    OpenRegOk = False
  End If
'
' if no valid HKEY value
'
  If RtHKey < HKEY_CLASSES_ROOT Or RtHKey > HKEY_DYN_DATA Then
    OpenRegistry = False
    OpenRegOk = False
    Exit Function
  End If
'
' keey local copies
'
  RootHKey = RtHKey
  SubDir = SbDr
'
' open registry path
'
  ReturnVal = RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_ALL_ACCESS, hKey)
'
' if error
'
  If ReturnVal Then
    OpenRegistry = False
    Exit Function
  End If
'
' all is OK
'
  OpenRegOk = True
  OpenRegistry = True
  Exit Function
'
' error trap
'
OpenReg:
  OpenRegOk = False
  OpenRegistry = False
End Function

'*************************************************************************
'CloseRegistry(): This function should be called after you're done with the registry
'(see other examples)
'*************************************************************************
Public Function CloseRegistry() As Boolean
  On Error Resume Next
  If RegCloseKey(hKey) Then
    CloseRegistry = False
    Exit Function
  End If

  CloseRegistry = True
  OpenRegOk = False
End Function

'*************************************************************************
'SortArrayAscending(): sort a variable array into ascending order
'*************************************************************************
Public Function SortArrayAscending(ValueList As Variant) As Variant
  On Error GoTo handelsort
  Dim RipVal As Variant
  Dim RipOrdinal As Long
  Dim RipDescent As Long
  Dim PrivateBuffer As Variant
  Dim Placed As Boolean
  Dim x As Long
  Dim y As Long

  If IsArray(ValueList) Then
    PrivateBuffer = ValueList
    'Ok, we start at the second position in the array and go
    'from there
    RipOrdinal = 1
    RipDescent = 1

    For y = 1 To UBound(PrivateBuffer)
      RipVal = PrivateBuffer(y)
      If y <> 1 Then RipDescent = y
        Do Until Placed
          If PrivateBuffer(RipDescent - 1) >= RipVal Then
            RipDescent = RipDescent - 1
            If RipDescent = 0 Then
              For x = y To RipDescent Step -1
                If x = 0 Then Exit For
                PrivateBuffer(x) = PrivateBuffer(x - 1)
              Next x
              PrivateBuffer(RipDescent) = RipVal
              Placed = True
            End If
          Else
            'shift the array to the right
            For x = y To RipDescent Step -1
              If x = 0 Then Exit For
              PrivateBuffer(x) = PrivateBuffer(x - 1)
            Next x
            'insert the ripped value
            PrivateBuffer(RipDescent) = RipVal
            Placed = True
          End If
        Loop
        Placed = False
      Next y

      SortArrayAscending = PrivateBuffer
   Else
      SortArrayAscending = ValueList
  End If
  Exit Function

handelsort:
  SortArrayAscending = ValueList
End Function

'*************************************************************************
' internal support. find last backslash in a string
'*************************************************************************
Private Function FindLastBackSlash(VarValue As String) As Integer
  FindLastBackSlash = InStrRev(VarValue, "\")
End Function

'*************************************************************************
' internal support
'*************************************************************************
Private Sub Class_Initialize()
  RootHKey = &H0
  SubDir = vbNullString
  hKey = 0
  OpenRegOk = False
End Sub

'*************************************************************************
' internal support
'*************************************************************************
Private Sub Class_Terminate()
  On Error Resume Next
  Call RegCloseKey(hKey)
End Sub

'*************************************************************************
' internal support
'*************************************************************************
Public Function OneBackOnKey()
  SubDir = Mid$(SubDir, 1, FindLastBackSlash(SubDir) - 1)
  CloseRegistry
  OpenRegistry RootHKey, SubDir
End Function

