VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fCError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose: CORE, do not change - trivial class to store error object properties
' 0.1.0    20220709    gueleh    Initially created
Option Explicit

Public lErrorNumber As Long
Public sErrorDescription As String

