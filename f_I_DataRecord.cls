VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_I_DataRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CORE, do not change
' Purpose: interface for standard class for data records,
'     based on a dictionary for field name and value pairs.
Option Explicit

Public sPrimaryKey As String
Public oDictValuesByFieldName As Scripting.Dictionary

Public Function bGetFieldValue(ByVal sFieldName As String, ByRef vFieldValue As Variant) As Boolean

End Function

Public Function bSetFieldValue(ByVal sFieldName As String, ByVal vFieldValue As Variant) As Boolean

End Function

