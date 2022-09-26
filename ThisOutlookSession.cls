VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const mapPath As String = "E:\Programming\VBA\Outlook\utils\data.txt"
Private WithEvents SentItems As Outlook.Items
Attribute SentItems.VB_VarHelpID = -1

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    'MsgBox "Application_ItemSend"
    SetSentItems
End Sub

Private Sub Application_NewMail()
    'MsgBox "Application_NewMail"
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    'MsgBox "Application_NewMailEx " & EntryIDCollection
End Sub

Private Sub Application_Quit()
    'MsgBox "Application_Quit"
End Sub

Private Sub Application_Reminder(ByVal Item As Object)
    'MsgBox "Application_Reminder"
End Sub

Private Sub Application_Startup()
    SetSentItems
End Sub

Private Sub SentItems_ItemAdd(ByVal Item As Object)
    Dim newLocation As Folder
    Dim str As String

    Utils.InitMap mapPath
    
    str = Utils.LookupFolderByAddress(Item.SendUsingAccount.CurrentUser.Address)
    
    If str <> vbNullString Then
        Set newLocation = Utils.GetFolderPath("Outlook\Sorted\" & str & "\Sent")
        
        If Not newLocation Is Nothing Then
            Item.Move newLocation
            Debug.Print "Moved [" & Item.Subject & "] from sent folder to " & str
        End If
    End If
End Sub

Private Sub SetSentItems()
    If SentItems Is Nothing Then
        Set SentItems = Session.GetDefaultFolder(olFolderSentMail).Items
    End If
End Sub
