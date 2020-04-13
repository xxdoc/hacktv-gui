Attribute VB_Name = "mSharedFunctions"
' This module contains functions which are shared across forms


' Declare a function for the DeleteUrlCacheEntryA API.
' This ensures that we clear the requested file from the IE browser cache before downloading the file.
' Otherwise we'd just get the cached copy.
Public Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

' Declare a function for the PathIsDirectoryEmptyA API (determines if a directory is empty or not)
Public Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long

' Declare a function for the ShellExecute API (used for running external applications)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hWnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWMINIMIZED As Long = 2

' Declare functions for GetProcAddress and GetModuleName (used by the IsWine property below for Wine detection)
' This will allow us to invoke the native Unix versions of hacktv than a Windows port
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

' This function uses the GetProcAddress and GetModuleName APIs to query kernel32.dll for
' the "wine_get_unix_file_name" export. This determines if we're running on Wine or not.
Public RunningOnWine As Boolean
Property Get IsWine() As Boolean
    IsWine = (GetProcAddress(GetModuleHandle("kernel32"), "wine_get_unix_file_name") <> 0)
End Property

' This function tests whether a file is locked or not
Public Function IsFileInUse(ByVal FileToTest) As Boolean
    Dim intFile As Integer
    intFile = FreeFile(0)
    On Error Resume Next
    Open FileToTest For Input Lock Read Write As #intFile
    If Err Then
    Close #intFile
        IsFileInUse = True
    Else
    Close #intFile
        IsFileInUse = False
    End If
End Function

' This function tests whether a file exists or not
' We use this instead of Dir() to avoid runtime errors in situations such as checking empty drives
Public Function DoesFileExist(ByRef FileName As String) As Boolean
    On Error GoTo ExistError
    If Len(Dir$(FileName)) > 0 Then DoesFileExist = True
    Exit Function
ExistError:
    DoesFileExist = False
End Function

' This function tests whether a folder exists or not
Public Function FolderExists(sFullPath As String) As Boolean
' This function is used to determine if a folder exists before deleting or creating it
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
End Function

' This function returns the text that exists between two specified strings
' From http://www.vbforums.com/showthread.php?570460-get-text-between-text
Public Function Between(Text As String, Before As String, After As String, Output() As String) As Long
    Dim lngA As Long, strBetween() As String
    Output = Split(Text, Before)
    For lngA = 1 To UBound(Output)
        strBetween = Split(Output(lngA), After, 2)
        If UBound(strBetween) = 1 Then
            Output(Between) = strBetween(0)
            Between = Between + 1
        End If
    Next lngA
    If Between > 0 Then
        ReDim Preserve Output(Between - 1)
    Else
        Output = Split(vbNullString)
    End If
End Function

