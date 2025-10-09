Attribute VB_Name = "f_pM_FloFra2Marker"
' CORE, do not change
'============================================================================================
'   NAME:     f_pM_FloFra2Marker
'============================================================================================
'   Purpose:  reads and sets a Custom XML marker to identify Flow Framework 2 workbooks
'   Access:   Private
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guenther.lehner@protonmail.com
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "f_pM_FloFra2Marker"
Private Const s_m_MarkerNamespace As String = "urn:flowframework2:marker"
Private Const s_m_MarkerXPath As String = "/*[local-name()='FlowFramework2' and namespace-uri()='urn:flowframework2:marker']"


Public Function b_f_p_WorkbookHasFlowFrameworkMarker(Optional ByRef oWkb As Workbook) As Boolean
    
   Dim wb As Workbook
   Dim parts As CustomXMLParts
   Dim part As CustomXMLPart
   Dim node As CustomXMLNode
   Dim i As Long
    
   On Error Resume Next
   If oWkb Is Nothing Then
      Set wb = ThisWorkbook
   Else
      Set wb = oWkb
   End If
   If Err.Number > 0 Then
      Set wb = ThisWorkbook
      Err.Clear
   End If
   
   On Error GoTo Catch
   

   Set parts = wb.CustomXMLParts.SelectByNamespace(s_m_MarkerNamespace)
   If parts Is Nothing Or parts.Count = 0 Then
      b_f_p_WorkbookHasFlowFrameworkMarker = False
      Exit Function
   End If

   For i = 1 To parts.Count
      Set part = parts(i)
      Set node = part.SelectSingleNode(s_m_MarkerXPath)
   
      If Not node Is Nothing Then
         b_f_p_WorkbookHasFlowFrameworkMarker = True
         Exit Function
      End If
   Next i

Catch:

End Function

Public Function b_f_p_AddFlowFrameworkMarker(Optional ByRef oWkb As Workbook) As Boolean
   Dim wb As Workbook
   Dim xml As String
   Dim newPart As CustomXMLPart

   On Error Resume Next
   If oWkb Is Nothing Then
      Set wb = ThisWorkbook
   Else
      Set wb = oWkb
   End If
   If Err.Number > 0 Then
      Set wb = ThisWorkbook
      Err.Clear
   End If
   On Error GoTo Catch

   ' Bereits vorhanden?
   If b_f_p_WorkbookHasFlowFrameworkMarker(wb) Then
      b_f_p_AddFlowFrameworkMarker = False
      Exit Function
   End If

   ' XML für neuen Marker
   xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
         "<FlowFramework2 xmlns=""" & s_m_MarkerNamespace & """>" & vbCrLf & _
         "   <Created>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</Created>" & vbCrLf & _
         "</FlowFramework2>"

   ' Marker hinzufügen
   Set newPart = wb.CustomXMLParts.Add(xml)

   b_f_p_AddFlowFrameworkMarker = Not newPart Is Nothing

Catch:
End Function

