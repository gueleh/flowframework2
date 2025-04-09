VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f_C_DriveMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   NAME:     i_C_DriveMapper
'============================================================================================
'   Purpose:  maps and unmaps a SharePoint path as a drive in the file system
'   Access:   Public
'   Type:     Klassenmodul
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'       Microsoft Scripting Runtime
'       Windows Script Host Object Model
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.0.0    29.07.2024    gueleh    Initial creation
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "i_C_DriveMapper"

Private o_m_MappedDrive As Scripting.Drive
Private o_m_FSO As Scripting.FileSystemObject
Private o_m_Network As WshNetwork
Private o_m_RootFolder As Folder

Private Sub Class_Initialize()
    Set o_m_FSO = New Scripting.FileSystemObject
    Set o_m_Network = New WshNetwork
End Sub

' Purpose: makes sure to unmount the drive once the class instance is terminated
Private Sub Class_Terminate()
    mUnmapDrive
End Sub

Public Property Get oRootFolder() As Scripting.Folder
    Set oRootFolder = o_m_RootFolder
End Property

Public Function sPathSanitized( _
    ByVal sPath As String, _
    Optional ByVal bWithFolderSeparator As Boolean = False _
    ) As String

    Const sSEP As String = "\"
    Dim sReturnValue As String
    
    On Error GoTo Catch

    If Left$(sPath, 4) = "http" Then
        If o_m_RootFolder Is Nothing Then SetRootFolderMappedDrive (sPath)
        sReturnValue = o_m_RootFolder.Path
    Else
        sReturnValue = sPath
    End If

    If bWithFolderSeparator _
    And Right$(sReturnValue, 1) <> sSEP _
    Then sReturnValue = sReturnValue & "\"

    sPathSanitized = sReturnValue
Finally:


    Exit Function
Catch:
    Resume Finally
   
End Function

' Purpose: gets and returns the root folder of a mapped drive for passed in network path
Public Sub SetRootFolderMappedDrive(ByVal sNetworkPath As String)
    Dim sDriveLetter As String
    Dim i As Integer
    Dim bHasError As Boolean

    On Error GoTo Catch
    mUnmapDrive

    For i = Asc("Z") To Asc("A") Step -1
        sDriveLetter = Chr(i)
        If Not o_m_FSO.DriveExists(sDriveLetter) Then
            o_m_Network.MapNetworkDrive sDriveLetter & ":", sNetworkPath
            Set o_m_MappedDrive = o_m_FSO.GetDrive(sDriveLetter)
            Set o_m_RootFolder = o_m_MappedDrive.RootFolder
            Exit For
        End If
    Next i

Finally:
    On Error Resume Next
    If bHasError Then mUnmapDrive
    Exit Sub

Catch:
    bHasError = True
    Resume Finally

End Sub

' Purpose: unmaps the drive, i.e. it is unmounted
Private Sub mUnmapDrive(Optional ByVal bForceRemoval As Boolean = True)
    If Not o_m_MappedDrive Is Nothing Then
        If o_m_MappedDrive.IsReady Then
            o_m_Network.RemoveNetworkDrive o_m_MappedDrive.DriveLetter & ":", bForceRemoval
        End If
        Set o_m_MappedDrive = Nothing
    End If
End Sub

