VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_C_VersionControlRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' core, dev - can be removed from production instance
'============================================================================================
'   NAME:     DEV_f_C_VersionControlRanges
'============================================================================================
'   Purpose:  build a collection with ranges subject to version control monitoring and provide this collection
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@wts.de
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.6.0    20.10.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "DEV_f_C_VersionControlRanges"

Private oCol_m_Ranges As New Collection
Private oWks_m_RangeInfos As Worksheet
Private l_m_RowStart As Long
Private l_m_ColName As Long
Private l_m_ColDefinedName As Long

Public Sub Construct _
( _
   ByRef oWks_arg_RangeInfos As Worksheet, _
   ByVal l_arg_RowStart As Long, _
   ByVal l_arg_ColName As Long, _
   ByVal l_arg_ColDefinedName As Long _
)
   Set oWks_m_RangeInfos = oWks_arg_RangeInfos
   l_m_RowStart = l_arg_RowStart
   l_m_ColDefinedName = l_arg_ColDefinedName
   l_m_ColName = l_arg_ColName
End Sub

Public Property Get oCol_prop_r_VersionControlRanges() As Collection
   Set oCol_prop_r_VersionControlRanges = oCol_m_Ranges
End Property

Public Function bBuildCollectionWithRanges() As Boolean

'Fixed, don't change
   Dim oC_Me As New f_C_CallParams
   oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "bBuildCollectionWithRanges" 'Name of the function
      .b_prop_rw_SilentError = True 'False will display a message box - you should only do this on entry level
      .s_prop_rw_ErrorMessage = "Your message here." 'A message that properly informs the user and the devs (silent errors will be logged nonetheless)
      .SetCallArgs "No args" 'If the sub takes args put the here like ("sExample:=" & sExample, "lExample:=" & lExample)
   End With
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
Try:
   On Error GoTo Catch

'>>>>>>> Your code here
   Dim oC As DEV_f_C_VersionControlRange
   Dim lRow As Long
   Dim bOneNotValidEntryFound As Boolean
   Dim sDefinedName As String
   Dim oRng As Range
   
   If oWks_m_RangeInfos Is Nothing Then
      Err.Raise _
         e_f_p_HandledError_AppSpecificError, , _
         "Construct has not been called, thus no worksheet object set."
   End If
      
   With oWks_m_RangeInfos
      For lRow = l_m_RowStart To .UsedRange.Rows.Count + l_m_RowStart
         sDefinedName = .Cells(lRow, l_m_ColDefinedName).Value2
         If Len(sDefinedName) > 0 Then
            On Error Resume Next
            Set oRng = ThisWorkbook.Names(sDefinedName).RefersToRange
            If Err.Number > 0 Then
               Err.Clear
               Err.Raise e_f_p_HandledError_AppSpecificError, , "Defined Name " & sDefinedName & " is not valid."
               f_p_RegisterError oC_Me, Err.Number, Err.Description
               f_p_HandleError oC_Me
               Err.Clear
               Set oRng = Nothing
               bOneNotValidEntryFound = True
            End If
            On Error GoTo Catch
            If Not oRng Is Nothing Then
               Set oC = New DEV_f_C_VersionControlRange
               oC.Construct .Cells(lRow, l_m_ColName).Value2, sDefinedName, oRng
               oCol_m_Ranges.Add oC
            End If
         End If
      Next lRow
   End With
      
'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next

'>>>>>>> Your code here
   If bOneNotValidEntryFound Then
      MsgBox "At least one entry for version control of range contents is not valid, " & _
         "please look into the error log after processing to find out more.", vbExclamation
   End If
'End of your code <<<<<<<

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Your custom settings here
'MAKE SURE TO REPLACE fbTemplateLowerLevel WITH THE NAME OF YOUR FUNCTION
   If oC_Me.oC_prop_r_Error Is Nothing Then bBuildCollectionWithRanges = True 'reports execution as successful to caller
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Fixed, don't change
   Exit Function
   
HandleError:
   af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel
'>>>>>>> Your code here
   'everything that must be executed in case of an error

'End of your code <<<<<<<

'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing _
   Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If

End Function

