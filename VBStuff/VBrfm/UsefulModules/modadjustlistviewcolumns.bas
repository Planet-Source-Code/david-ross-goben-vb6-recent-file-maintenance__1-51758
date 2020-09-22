Attribute VB_Name = "modAdjustListViewColumns"
Option Explicit
'~modAdjustListViewColumns.bas;
'Modify ListView control to auto-adjust column widths
'********************************************************************************
' modAdjustListViewColumns - The AdjustListViewColumns() function modifies a
'                            ListView control to auto-adjust column widths to fit
'                            The contents of the columns. If you specify a specific
'                            column with the optional SpecificColumn parameter,
'                            then only that column is modified, otherwise all
'                            columns are modified. If you specify True for the
'                            AccountForHeaders parameter, then the subroutine will
'                            automatically size the columns to fit the Header text.
'                            When this is applied to the last column, the last
'                            column's width will be automatically sized to fill
'                            the remaining width of the ListView control.
'
'EXAMPLE: Assuming a ListView1 control with 4 (0-3) columns:
'  Call AdjustListViewColumns(ListView1, False)    'adjust all columns for contents
'  Call AdjustListViewColumns(ListView1, True)     'adust all columns for header text,
'                                                  'except last column, which will fill
'                                                  'the remainder of the control width.
'  Call AdjustListViewColumns(ListView1, True, 3)  'Adjust only the last column to fill
'                                                  'the remainder of the control width.
'********************************************************************************
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Sub AdjustListViewColumns(LVW As ListView, AccountForHeaders As Boolean, Optional SpecificColumn As Integer = -1)
  Dim col As Long, lParam As Long
'
' determine how to handle headers
'
  If AccountForHeaders Then
    lParam = LVSCW_AUTOSIZE_USEHEADER
  Else
    lParam = LVSCW_AUTOSIZE
  End If
'
' autosize all columns
'
  If SpecificColumn = -1 Then
    For col = 0 To LVW.ColumnHeaders.Count - 1
      SendMessageByLong LVW.hWnd, LVM_SETCOLUMNWIDTH, col, ByVal lParam
    Next col
  Else
    col = SpecificColumn
    If col >= 0 And col < LVW.ColumnHeaders.Count Then
      SendMessageByLong LVW.hWnd, LVM_SETCOLUMNWIDTH, col, ByVal lParam
    End If
  End If
End Sub
