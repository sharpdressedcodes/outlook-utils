Attribute VB_Name = "Macros"
Option Explicit

Public Sub MoveEmailsToFolder()
    Dim Address As String
    Dim isSender As Boolean
    Dim r As VbMsgBoxResult
    Dim newFolder As Folder
    Dim i As Long
    Dim j As Long
    Dim Recipients As Outlook.Recipients
    Dim Recipient As Outlook.Recipient
    Dim max As Long
    Dim itemCount As Long
    
    itemCount = Explorers.Item(1).CurrentFolder.Items.Count
    
    If itemCount = 0 Then
        MsgBox "No items found to move, please select a folder that contains items", vbCritical, "No items"
        Exit Sub
    End If
                    
    Do While Address = vbNullString
        Address = InputBox("Which email address would you like to scan for?", "Enter email address")
        
        If StrPtr(Address) = 0 Then
            ' user clicked cancel
            Exit Sub
        ElseIf Address = vbNullString Then
            MsgBox "Invalid input, please enter something and try again.", vbCritical, "Invalid"
        End If
    Loop
    
    Address = LCase$(Address)
    r = MsgBox("Should we consider this address as the sender?", vbYesNoCancel + vbQuestion, "Sender or Receiver")
    
    Select Case r
        Case vbCancel
            Exit Sub
        Case Else
            isSender = (r = vbYes)
    End Select
    
    Set newFolder = Utils.PickFolder
    
    If newFolder Is Nothing Then
        ' user clicked cancel
        Exit Sub
    End If
    
    Dim col As New Collection
        
    For i = 1 To itemCount
        Dim Item As Outlook.MailItem
        Set Item = Explorers.Item(1).CurrentFolder.Items.Item(i)
        
        If isSender Then
            If LCase$(Item.SenderEmailAddress) = Address Then
                col.Add Item
            End If
        Else
            max = Item.Recipients.Count
            
            If max > 0 Then
                For j = 1 To max
                    If LCase$(Item.Recipients.Item(j).Address) = Address Then
                        Set Recipient = Item.Recipients.Item(j)
                        Exit For
                    End If
                Next
            End If
    
            If Not Recipient Is Nothing Then
                col.Add Item
            End If
        End If
    Next
    
    max = col.Count
    
    If max > 0 Then
        For i = 1 To max
            col(i).Move newFolder
            Debug.Print "Moved [" & col(i).Subject & "]"
        Next
    End If
    
    MsgBox "Moved " & CStr(max) & " item" & IIf(max = 1, vbNullString, "s"), vbInformation
End Sub
