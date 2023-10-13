VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_i_C_VersionControlExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================================================================================
'   NAME:     DEV_i_C_VersionControlExport
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
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

' Purpose: print project reference data to text file for version control
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
Public Function bExportReferenceData() As Boolean

   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
   
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

Finally:
   On Error Resume Next
   Close #iFileNumber
   bExportReferenceData = bExecutedSuccessfully
   Exit Function

Catch:
   bExecutedSuccessfully = False
   
   If Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      Debug.Print Err.Number, Err.Description
      Resume Finally
   End If

End Function


' Purpose: exports names and codenames of worksheets into text file for version control
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
Public Function bExportWorksheetNameData() As Boolean

   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
   
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

Finally:
   On Error Resume Next
   Close #iFileNumber
   bExportWorksheetNameData = bExecutedSuccessfully
   Exit Function

Catch:
   bExecutedSuccessfully = False
   
   If Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      Debug.Print Err.Number, Err.Description
      Resume Finally
   End If

End Function

' Purpose: exports data of all names to a text file
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
Public Function bExportNameData() As Boolean

   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
   
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

Finally:
   On Error Resume Next
   Close #iFileNumber
   bExportNameData = bExecutedSuccessfully
   Exit Function

Catch:
   bExecutedSuccessfully = False
   
   If Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      Debug.Print Err.Number, Err.Description
      Resume Finally
   End If

End Function

' Purpose: exports all VB components of types worksheet, module, class module and form to the project folder
' 1.0.0  01.06.2023  gueleh   Carved out from framework and removed dependency
Public Function bExportAllComponents() As Boolean

   Dim bResumedOnce As Boolean
   Dim bExecutedSuccessfully As Boolean

Try:
   On Error GoTo Catch
   bExecutedSuccessfully = True
   
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

Finally:
   On Error Resume Next
   bExportAllComponents = bExecutedSuccessfully
   Exit Function

Catch:
   bExecutedSuccessfully = False
   
   If Not bResumedOnce Then
      bResumedOnce = True: Stop: Resume
   Else
      Debug.Print Err.Number, Err.Description
      Resume Finally
   End If

End Function
