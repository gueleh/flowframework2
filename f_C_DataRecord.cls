VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_DataRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE, do not change
' Purpose: standard class for data records, based on a dictionary for field name and value pairs.
'     based on interface for polymorphism and extension if needed in app
' 0.15.0    15.05.2023    gueleh    Initially created
Option Explicit
Implements f_I_DataRecord

Private s_m_PrimaryKey As String
Private oDict_m_ValuesByFieldName As Scripting.Dictionary

' Purpose: returns value of a field or false in case of an error
Private Function f_I_DataRecord_bGetFieldValue _
( _
   ByVal sFieldName As String, _
   vFieldValue As Variant _
) As Boolean
   
   On Error GoTo Catch
   If Not oDict_m_ValuesByFieldName.Exists(sFieldName) Then Err.Raise 9999, , "Invalid field key"
   vFieldValue = oDict_m_ValuesByFieldName(sFieldName)
   f_I_DataRecord_bGetFieldValue = True
   Exit Function
Catch:
End Function

' Purpose: sets value of a field and returns false in case of an error, setting values of error object
Private Function f_I_DataRecord_bSetFieldValue _
( _
   ByVal sFieldName As String, _
   ByVal vFieldValue As Variant _
) As Boolean
   
   On Error GoTo Catch
   oDict_m_ValuesByFieldName(sFieldName) = vFieldValue
   f_I_DataRecord_bSetFieldValue = True
   Exit Function
Catch:
End Function

Private Property Set f_I_DataRecord_oDictValuesByFieldName _
( _
   ByVal oDictValuesByFieldName As Scripting.Dictionary _
)
   Set oDict_m_ValuesByFieldName = oDictValuesByFieldName
End Property

Private Property Get f_I_DataRecord_oDictValuesByFieldName() As Scripting.Dictionary
   Set f_I_DataRecord_oDictValuesByFieldName = oDict_m_ValuesByFieldName
End Property

Private Property Let f_I_DataRecord_sPrimaryKey(ByVal sPrimaryKey As String)
   s_m_PrimaryKey = sPrimaryKey
End Property

Private Property Get f_I_DataRecord_sPrimaryKey() As String
   f_I_DataRecord_sPrimaryKey = s_m_PrimaryKey
End Property

Private Sub Class_Initialize()
   Set oDict_m_ValuesByFieldName = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
   Set oDict_m_ValuesByFieldName = Nothing
End Sub
