Attribute VB_Name = "devfpMSandBox"
' Purpose: Sandbox for dev experiments
' 0.1.0    20220709    gueleh    Initially created
Option Explicit
Option Private Module

' Purpose: tests zero sanitation manually, if the code executes without stopping the tests were successful
' 0.11.0    05.08.2022    gueleh    Initially created
Private Sub mManualTest_RangeArrayProcessorZeroSanitation()
   Dim oC As New fCRangeArrayProcessor
   Dim va() As Variant
   DEV_Reset_devfwksTestCanvas
   With devfwksTestCanvas
      .Range("A1").Value = "ID"
      .Range("B1").Value = "ID2"
      .Range("C1").Value = "Value"
      .Range("A2").Value = "'01"
      .Range("B2").Value = "AgA"
      .Range("C2").Value = "What is AgA?"
      .Columns.AutoFit
      va = .Range("A1").CurrentRegion.Formula
      .Range("A1").CurrentRegion.Formula = va
      Debug.Assert .Range("A2").Value = "1"
      oC.SanitizeLeadingZeroItems va
      .Range("A1").CurrentRegion.Formula = va
      Debug.Assert .Range("A2").Value = "01"
   End With
   DEV_Reset_devfwksTestCanvas
End Sub
