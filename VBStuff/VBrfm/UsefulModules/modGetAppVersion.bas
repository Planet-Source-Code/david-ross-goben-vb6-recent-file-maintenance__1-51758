Attribute VB_Name = "modGetAppVersion"
Option Explicit
'~modGetAppVersion.bas;
'Return the application version numbers in the format "major.minor.revision"
'********************************************************************************
' modGetAppVersion - The GetAppVersion() returns a string containing the application
'                    version information in the form "major.minor.revision".
'EXAMPLE:
'  Debug.Print "This app's verison is v" & GetAppVersion
'********************************************************************************

Public Function GetAppVersion() As String
  GetAppVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function
