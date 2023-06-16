VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_RangeArrayProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE, do not change or remove
'============================================================================================
'   NAME:     f_C_RangeArrayProcessor
'============================================================================================
'   Purpose:  convenient array-based processing for range data
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.11.0    05.08.2022    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "f_C_RangeArrayProcessor"

Private va_m_Data() As Variant
Private rng_m_DataRange() As Range
Private b_m_FirstRowIsHeader As Boolean
Private b_m_DataHavePrimaryKey As Boolean
Private b_m_ReadValuesInsteadOfFormulas As Boolean
Private b_m_WriteValuesInsteadOfFormulas As Boolean

'TODO: Write function to construct class instance
'TODO: Establish dictionary with column keys - col numbers or header in range data is the key
'TODO: Establish dictionary with row keys, based on (composite) primary key or on row numbers
'TODO: Add support for a name, if none is set, the address of the range is used

' Purpose: adds ' to all array items starting with 0
' 0.11.0    05.08.2022    gueleh    Initially created
Public Sub SanitizeLeadingZeroItems(ByRef vaData() As Variant)
   Dim lRow As Long
   Dim lColumn As Long
   On Error Resume Next 'make sure to process all that can be processed
   For lRow = LBound(vaData, 1) To UBound(vaData, 1)
      For lColumn = LBound(vaData, 2) To UBound(vaData, 2)
         If Left$(vaData(lRow, lColumn), 1) = "0" Then vaData(lRow, lColumn) = "'" & vaData(lRow, lColumn)
      Next lColumn
   Next lRow
End Sub
