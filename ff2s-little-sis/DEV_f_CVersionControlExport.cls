VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_f_CVersionControlExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------------------
' CORE-DEV, do not change
'============================================================================================
'   NAME:     DEV_f_CVersionControlExport
'============================================================================================
'   Purpose:  exports all data relevant for convenient version control
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
' 0.1.0    17.03.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "DEV_f_CVersionControlExport"

' Purpose: print project reference data to text file for version control
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function bExportReferenceData() As Boolean

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_FUNCTION As String = "bExportReferenceData"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here

   'Change this is necessary
   bIsSilentError = True 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = ""

   Dim oC As VBIDE.Reference
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   Dim sReference As String
   
   sFileFullPath = ThisWorkbook.Path & "\" & "References.txt"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oC In ThisWorkbook.VBProject.References
      With oC
         sReference = .Name & "|" & .Description & "|" _
            & .GUID & "|" & .Major & "|" & .Minor & "|" & .FullPath
      End With
      Print #iFileNumber, sReference
   Next oC

'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Finally:
   On Error Resume Next
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not
      
   Close #iFileNumber
   
   'change this to meet the name of your function
   bExportReferenceData = bExecutedSuccessfully


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Exit Function
Catch:
   'for informing the caller that the execution failed
   bExecutedSuccessfully = False
   
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_FUNCTION, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed

End Function


' Purpose: exports names and codenames of worksheets into text file for version control
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function bExportWorksheetNameData() As Boolean
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_FUNCTION As String = "bExportWorksheetNameData"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here

   'Change this is necessary
   bIsSilentError = True 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = ""

   
   
   Dim oWks As Worksheet
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   
   sFileFullPath = ThisWorkbook.Path & "\" & "WorksheetNames.txt"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oWks In ThisWorkbook.Worksheets
      Print #iFileNumber, oWks.CodeName, oWks.Name
   Next oWks

'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Finally:
   On Error Resume Next
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not
      
   Close #iFileNumber
   
   'change this to meet the name of your function
   bExportWorksheetNameData = bExecutedSuccessfully


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Exit Function
Catch:
   'for informing the caller that the execution failed
   bExecutedSuccessfully = False
   
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_FUNCTION, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed


End Function

' Purpose: exports data of all names to a text file
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function bExportNameData() As Boolean
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_FUNCTION As String = "bExportNameData"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here

   'Change this is necessary
   bIsSilentError = True 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = ""
   
   
   Dim oName As Name
   Dim iFileNumber As Integer
   Dim sFileFullPath As String
   
   sFileFullPath = ThisWorkbook.Path & "\" & "Names.txt"
   On Error Resume Next
   Kill sFileFullPath
   On Error GoTo Catch
   iFileNumber = FreeFile
   
   Open sFileFullPath For Output As iFileNumber
   
   For Each oName In ThisWorkbook.Names
      If oName.MacroType <> xlNone Then
         Print #iFileNumber, oName.Name, oName.Category, oName.Parent.Name, oName.Comment
      Else
         Print #iFileNumber, oName.Name, oName.RefersTo, oName.Parent.Name, oName.Comment
      End If
   Next oName

'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Finally:
   On Error Resume Next
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not
      
   Close #iFileNumber
   
   'change this to meet the name of your function
   bExportNameData = bExecutedSuccessfully


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Exit Function
Catch:
   'for informing the caller that the execution failed
   bExecutedSuccessfully = False
   
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_FUNCTION, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed

End Function

' Purpose: exports all VB components of types worksheet, module, class module and form to the project folder
' 0.1.0    17.03.2023    gueleh    Initially created
Public Function bExportAllComponents() As Boolean
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Change name of the sub if you want to have this information in the error log
   Const sNAME_OF_FUNCTION As String = "bExportAllComponents"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean
   Dim oCError As f_C_Error
   Dim bIsSilentError As Boolean
   Dim sErrorMessage As String

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
'End Fixed
   
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here

   'Change this is necessary
   bIsSilentError = True 'if False, a MsgBox will be displayed. Otherwise only an entry in the error log.
   sErrorMessage = ""

   
   
   Dim ocmpComponent As VBIDE.VBComponent
   Dim bExport As Boolean
   Dim sFileName As String
   
   For Each ocmpComponent In ThisWorkbook.VBProject.VBComponents
      bExport = True
      sFileName = ocmpComponent.Name
      Select Case ocmpComponent.Type
         Case vbext_ct_ClassModule, 100 '100 is a worksheet
            sFileName = sFileName & ".cls"
         Case vbext_ct_MSForm
            sFileName = sFileName & ".frm"
         Case vbext_ct_StdModule
            sFileName = sFileName & ".bas"
         Case vbext_ct_Document
            bExport = False
      End Select
      
      If bExport Then
         ocmpComponent.Export ThisWorkbook.Path & "\" & sFileName
      End If
      
   Next ocmpComponent

'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Finally:
   On Error Resume Next
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>> Your code here
'>>>>>>> everything that MUST BE EXECUTED regardless of an error or not
      
   'change this to meet the name of your function
   bExportAllComponents = bExecutedSuccessfully


'End of your code <<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Fixed, don't change
   Exit Function
Catch:
   'for informing the caller that the execution failed
   bExecutedSuccessfully = False
   
   'Set error data for logging if not already existent
   If oCError Is Nothing Then
      Set oCError = New f_C_Error
      oCError.SetErrorData _
         Err.Number, Err.Description, sNAME_OF_FUNCTION, _
         s_m_COMPONENT_NAME, bIsSilentError, sErrorMessage
   End If
   
   'If in debug mode, then the code will stop once so that you can step into
   '  the row which caused the error.
   '  If already stopped and resumed once, the error data are sent to the error handler
   If b_p_DEBUG_MODE_IS_ON _
   And Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      f_p_LogError oCError: Resume Finally
   End If
'End Fixed

End Function


