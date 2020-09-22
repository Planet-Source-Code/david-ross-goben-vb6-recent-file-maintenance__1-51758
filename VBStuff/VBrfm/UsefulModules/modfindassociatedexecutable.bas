Attribute VB_Name = "modFindAssociatedExecutable"
Option Explicit
'~modFindAssociatedExecutable.bas;
'Find the executable file associated with a provided filepath
'*******************************************************************************
' modFindAssociatedExecutable: The FindAssociatedExecutable() function will find
'                              find the executable file associated with a provided
'                              filepath. The OpenWithDialog() subroutine will
'                              display the "Open With..." dialog box. This should
'                              be used for files that do not have an application
'                              associated with then. When an association (or open
'                              file if the user disables the always open with checkbox,
'                              the newly associated application opens the file.
'
' FindAssociatedExecutable() return values:
'                 1: test file does not exist. EXEPath will be blank.
'                 0: EXEPath will contain the path to the executable file
'                -1: file exisits, but has no associations. EXEPath will be blank.
' EXAMPLE:
' Dim Path As String, Flg As Integer, TestFile As String
'
' TestFile = "D:\Readme.txt"                                 'file to check
' Select Case FindAssociatedExecutable(TestFile, Path)       'find association
'   Case 0                                                   'is associated
'     Debug.Print TestFile & "is associated with: " & Path
'   Case 1                                                   'does not exist
'     Debug.Print TestFile & " does not exist"
'   Case Else
'     Debug.Print TestFile & " is not associated with any file"
' End Select
'*******************************************************************************

'*************************************************
' API call used by FindAssociatedExecutable
'*************************************************
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Public Function FindAssociatedExecutable(FilePath As String, EXEPath As String) As Integer
  Dim S As String, Temp As String, lint As Long
  
  FindAssociatedExecutable = 1          'default to file exisits, but no associations
  EXEPath = vbNullString                'init blank return
  Temp = Trim$(FilePath)                'trim up the provited test file  If FileExists(Temp) Then              'if file exists, then we will check for associations
  S = String(260, " ") & vbNullChar     'init result string
  lint = FindExecutable(Temp, vbNullString, S)    'get data
  If lint > 32 Then                     'we are OK
    lint = InStr(S, vbNullChar)         'so, find null terminator
    EXEPath = Left$(S, lint - 1)        'set path
    FindAssociatedExecutable = 0        'flag as success
  Else
    If Len(Dir(Temp)) Then
      FindAssociatedExecutable = -1     'did not find the file. Return error code
    End If
  End If
End Function
