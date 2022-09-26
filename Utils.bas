Attribute VB_Name = "Utils"
Option Explicit

Private map As New Collection
Private mapPath As String

Public Function GetAppPath(Optional ByVal app As String) As String
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    If app = vbNullString Then
        app = "OUTLOOK.EXE"
    End If
    
    GetAppPath = WSHShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & app & "\")
End Function

Public Sub InitMap(ByVal path As String)
    Dim i As Long
    Dim s As String
    Dim lines() As String
    
    If map.Count > 0 Then
        Exit Sub
    End If
    
    On Error GoTo ErrKill
    
    s = ReadBinary(path)
    lines = Split(s, vbCrLf)
    
    For i = 0 To UBound(lines)
        If LenB(lines(i)) Then
            Dim data() As String
            
            data = Split(lines(i), ",")
            map.Add data(1), data(0)
        End If
    Next
    Exit Sub
    
ErrKill:
    MsgBox Err.Description, vbCritical
End Sub

Public Function LookupFolderByAddress(ByVal Address As String) As String
    On Error GoTo ErrKill
    
    If map.Count = 0 Then
        InitMap
    End If
    
    LookupFolderByAddress = map(LCase$(Address))
    Exit Function
    
ErrKill:
    LookupFolderByAddress = vbNullString
End Function

'Name                               Value   Description
'
'olFolderCalendar                   9       Calendar folder
'olFolderContacts                   10      Contacts folder
'olFolderDeletedItems               3       Deleted Items folder
'olFolderDrafts                     16      Drafts folder
'olFolderInbox                      6       Inbox folder
'olFolderJournal                    11      Journal folder
'olFolderJunk                       23      Junk E-Mail folder
'olFolderNotes                      12      Notes folder
'olFolderOutbox                     4       Outbox folder
'olFolderSentMail                   5       Sent Mail folder
'olFolderSuggestedContacts          30      Suggested Contacts folder
'olFolderTasks                      13      Tasks folder
'olFolderToDo                       28      To Do folder
'olPublicFoldersAllPublicFolders    18      All Public Folders folder in Exchange Public Folders store (Exchange only)
'olFolderRssFeeds                   25      RSS Feeds folder

Public Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
        
    On Error GoTo GetFolderPath_Error
    
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
        
GetFolderPath_Error:

    Set GetFolderPath = Nothing
End Function

Public Function PickFolder() As Outlook.Folder
    Dim objNS As NameSpace
    Dim objFolder As Folder

    Set objNS = Application.GetNamespace("MAPI")
    Set objFolder = objNS.PickFolder

    If TypeName(objFolder) <> "Nothing" Then
        Set PickFolder = objFolder
    Else
        Set PickFolder = Nothing
    End If

    Set objNS = Nothing
End Function

Public Function ReadBinary(ByVal directory As String) As String
    Dim FF As Integer
    
    FF = FreeFile
    
    If Dir$(directory) = vbNullString Then
        ReadBinary = vbNullString
        Exit Function
    End If
    
    Open directory For Binary Access Read As #FF
        ReadBinary = Space$(LOF(FF))
        Get #FF, , ReadBinary
    Close #FF
End Function
