VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "i_C_StringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   NAME:     i_C_StringFormatter
'============================================================================================
'   Purpose:  this class allows for inserting values of variables into strings based on placeholder chars
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.0.0    02.06.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Public Function sFormatted( _
   ByVal sINPUT As String, _
   ByVal sPLACEHOLDER As String, _
   ParamArray vaVariables() _
) As String

   Const sERROR As String = "<ERROR> when applying i_C_StringFormatter.sFormatted()"

   Dim lIterator As Long
   Dim lPlaceholderCount As Long
   Dim lVariableCount As Long
   Dim sReturn As String
   
   On Error GoTo Catch
    ' Count the placeholders
    lPlaceholderCount = _
      (Len(sINPUT) - Len(Replace(sINPUT, sPLACEHOLDER, ""))) / Len(sPLACEHOLDER)
    
    ' Count the variables
    lVariableCount = UBound(vaVariables) - LBound(vaVariables) + 1
    
    ' If the count of placeholders doesn't match the count of variables, return an error
    If lPlaceholderCount <> lVariableCount Then Err.Raise 9999
    
    sReturn = sINPUT
    
    ' Replace each placeholder with its corresponding variable
    For lIterator = LBound(vaVariables) To UBound(vaVariables)
        sReturn = Replace(sReturn, sPLACEHOLDER, vaVariables(lIterator), 1, 1)
    Next lIterator

Finally:
   sFormatted = sReturn
   Exit Function

Catch:
   sReturn = sERROR
   Resume Finally
End Function

