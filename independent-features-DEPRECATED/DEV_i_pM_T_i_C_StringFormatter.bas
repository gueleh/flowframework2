Attribute VB_Name = "DEV_i_pM_T_i_C_StringFormatter"
Option Explicit
Option Private Module

Private Sub mTestStringFormatter()
   Const sINPUT As String = "My name is ${s} and I am ${s} years old."
   Const sEXPECTED_OUTPUT As String = "My name is Anders and I am 7 years old."
   Const sPLACEHOLDER As String = "${s}"
   Const sPLACEHOLDER2 As String = "${t}"
   Const sNAME As String = "Anders"
   Const lAGE As Long = 7
   Const sERROR As String = "<ERROR> when applying i_C_StringFormatter.sFormatted()"
   Dim oCTest As i_C_Test
   Dim oColTests As New Collection
   Dim oCLog As New i_C_TestLogger
   Dim oC As New i_C_StringFormatter
   Dim sAllInputs As String
   
   Set oCTest = New i_C_Test
   sAllInputs = "Input String: " & sINPUT & vbLf _
      & "Placeholder: " & sPLACEHOLDER & vbLf _
      & "Variable 1: " & sNAME & vbLf _
      & "Variable 2: " & lAGE
   
   
   
   oCTest.AddTest "1", "Return correct string", _
      sAllInputs, sEXPECTED_OUTPUT, "Instance created", _
      oC.sFormatted(sINPUT, sPLACEHOLDER, sNAME, lAGE) = sEXPECTED_OUTPUT
   
   oColTests.Add oCTest
   
  
   Set oCTest = New i_C_Test
   sAllInputs = "Input String: " & sINPUT & vbLf _
      & "Placeholder: empty string" & sPLACEHOLDER & vbLf _
      & "Variable 1: " & sNAME & vbLf _
      & "Variable 2: " & lAGE
   oCTest.AddTest "2", "Return error message as string when placeholder is empty", _
      sAllInputs, sERROR, "Instance created", _
      oC.sFormatted(sINPUT, vbNullString, sNAME, lAGE) = sERROR
   
   oColTests.Add oCTest
   
   Set oCTest = New i_C_Test
   sAllInputs = "Input String: " & sINPUT & vbLf _
      & "Placeholder: " & sPLACEHOLDER & vbLf _
      & "Variable 1: " & sNAME & vbLf _
      & "Variable 2: " & sNAME & vbLf _
      & "Variable 3: " & lAGE
   oCTest.AddTest "3", "Return error message when too few placeholders are in the input", _
      sAllInputs, sERROR, "Instance created", _
      oC.sFormatted(sINPUT, sPLACEHOLDER, sNAME, sNAME, lAGE) = sERROR
   
   oColTests.Add oCTest
   
   Set oCTest = New i_C_Test
   sAllInputs = "Input String: " & sINPUT & vbLf _
      & "Placeholder: " & sPLACEHOLDER & vbLf _
      & "Variable 1: " & sNAME
   oCTest.AddTest "4", "Return error message when too few variables are passed in", _
      sAllInputs, sERROR, "Instance created", _
      oC.sFormatted(sINPUT, sPLACEHOLDER, sNAME) = sERROR
   
   oColTests.Add oCTest
   
   
   Set oCTest = New i_C_Test
   sAllInputs = "Input String: " & sINPUT & vbLf _
      & "Placeholder: " & sPLACEHOLDER2 & vbLf _
      & "Variable 1: " & sNAME
   oCTest.AddTest "5", "Return error message when wrong placeholder is passed in", _
      sAllInputs, sERROR, "Instance created", _
      oC.sFormatted(sINPUT, sPLACEHOLDER2, sNAME, lAGE) = sERROR
   
   oColTests.Add oCTest
   
   
   oCLog.LogTestResults wksTestCanvas, oColTests
End Sub
