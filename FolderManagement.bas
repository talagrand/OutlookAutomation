Function FolderGetChildren(Folder As Outlook.MAPIFolder) As String()
    Dim Folders() As String
    Dim MoreFolders() As String
    Dim SizeMarker As Integer
    ReDim Folders(0 To 0) As String
      
    For i = 1 To Folder.Folders.Count
        ReDim Preserve Folders(0 To UBound(Folders) + 1) As String
        
        Folders(UBound(Folders)) = Folder.Folders(i).FolderPath
        If Folder.Folders(i).Folders Is Nothing Then
        Else
            MoreFolders = FolderGetChildren(Folder.Folders(i))
            SizeMarker = UBound(Folders)
            ReDim Preserve Folders(0 To UBound(Folders) + UBound(MoreFolders)) As String
            For j = 1 To UBound(MoreFolders)
                Folders(SizeMarker + j) = MoreFolders(j)
            Next
        End If
    Next
    
    FolderGetChildren = Folders
End Function

Function FolderPopup(StartFolder As String, DefaultFolder As String) As Outlook.MAPIFolder
    Dim objFolder As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace, objItem As Outlook.MailItem
    Dim FolderList() As String
    Dim MatchingFolderList() As String
    
    
    Set objNS = Application.GetNamespace("MAPI")
    If StartFolder <> "" Then
        Set objFolder = objNS.Folders(StartFolder)
        FolderList = FolderGetChildren(objFolder)
    Else
        ReDim FolderList(0 To 0) As String
        ' XXX fix here
        ' To objNS.Folders.Count
        For i = 1 To objNS.Folders.Count
            Dim tempFoo As String
            'tempFoo = InputBox(objNS.Folders(i).Name, "Toplevels", DefaultFolder)
            If (objNS.Folders(i).name <> "Public Folders") And (objNS.Folders(i).name <> "Archives") Then
                MatchingFolderList = FolderGetChildren(objNS.Folders(i))
                SizeMarker = UBound(FolderList)
                ReDim Preserve FolderList(0 To UBound(FolderList) + UBound(MatchingFolderList)) As String
                For j = 1 To UBound(MatchingFolderList)
                    'tempFoo = InputBox(MatchingFolderList(j), "Sub", DefaultFolder)
                    FolderList(SizeMarker + j) = MatchingFolderList(j)
                Next
            End If
        Next i
    End If
    
    MatchingFolderList = FolderList
    
    While 1:
        Statement = "Enter a part of the Folder Name:"
        If (UBound(MatchingFolderList) < 15) Then
            Statement = Statement + Chr(10) + Chr(10)
            For i = 1 To UBound(MatchingFolderList)
                Statement = Statement + "  - " + MatchingFolderList(i) + Chr(10)
            Next i
        Else
            Statement = Statement + "(" + Str(UBound(MatchingFolderList)) + " matching)"
        End If
        origfragment = InputBox(Statement, "Folder Name", DefaultFolder)
        fragment = UCase(origfragment)
        If fragment = "" Then
            Set FolderPopup = Nothing
            Exit Function
        End If
        DefaultFolder = origfragment
        
        ReDim MatchingFolderList(0 To 0) As String
        For i = 1 To UBound(FolderList)
            EndPartStart = InStrRev(FolderList(i), "\")
            EndPart = Mid(FolderList(i), EndPartStart)
            
            If InStr(1, UCase(EndPart), fragment) > 0 Then
                ' Partial match on a Folder Name
                ReDim Preserve MatchingFolderList(0 To UBound(MatchingFolderList) + 1) As String
                MatchingFolderList(UBound(MatchingFolderList)) = FolderList(i)
            End If
            If UCase(Right(FolderList(i), Len(fragment) + 1)) = "\" + fragment Then
                ' THey gave an exact folder name
                ReDim MatchingFolderList(0 To 1) As String
                MatchingFolderList(1) = FolderList(i)
                Exit For
            End If
        Next
        
        If UBound(MatchingFolderList) = 1 Then
            FolderPath = Split(MatchingFolderList(1), "\")
            Set objFolder = objNS.Folders(FolderPath(2))
            For i = 3 To UBound(FolderPath)
                Set objFolder = objFolder.Folders(FolderPath(i))
            Next i

            Set FolderPopup = objFolder
            Exit Function
        End If
                
    Wend
     
End Function

Function ItemFilter(obj) As Boolean
    ItemFilter = obj.Class = olMail Or obj.Class = olReport Or (TypeOf obj Is Outlook.MeetingItem)
End Function


Function AddMessagesFromSelection(ByRef objItems, ByRef rgMailItems(), ByRef AllOneConvo As Boolean)
    Dim objItem
    AllOneConvo = True
    Dim ConvoId As String
    Dim Convo As Outlook.Conversation
    
    
    For Each objItem In objItems
        If ItemFilter(objItem) Then
            ReDim Preserve rgMailItems(0 To UBound(rgMailItems) + 1)
            Set rgMailItems(UBound(rgMailItems)) = objItem
            If ConvoId = "" Then
                ConvoId = objItem.ConversationID
            Else
                If ConvoId <> objItem.ConversationID Then
                    AllOneConvo = False
                End If
            End If
        End If
    Next
End Function


Sub MoveToSomeFolder()
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem
    
    Dim objCurrentFolder As Outlook.MAPIFolder
    Set objCurrentFolder = Application.ActiveExplorer.CurrentFolder
    Dim rgMailItems()
    ReDim rgMailItems(0 To 0)
    Dim AllOneConvo As Boolean
    AllOneConvo = False
    
    If Application.ActiveExplorer.Selection.Count = 0 Then
        'Require that this procedure be called only when a message is selected
        'Maybe a conversation is selected, check that first
        Dim convHeaders As Outlook.Selection
        Set convHeaders = Application.ActiveExplorer.Selection.GetSelection(olConversationHeaders)
        
        If convHeaders.Count >= 1 Then
            If convHeaders.Count = 1 Then
                AllOneConvo = True
            End If
            
            Dim objConv As Outlook.ConversationHeader
            For Each objConv In convHeaders
                Dim objSimpleItems As Outlook.SimpleItems
                Set objSimpleItems = objConv.GetItems()
                Dim DummyAllOneConvoInThisConvo As Boolean
                AddMessagesFromSelection objItems:=objSimpleItems, rgMailItems:=rgMailItems, AllOneConvo:=DummyAllOneConvoInThisConvo
            Next
        Else
            MsgBox "Must select at least 1 message"
            Exit Sub
        End If
    Else
        AddMessagesFromSelection objItems:=Application.ActiveExplorer.Selection, rgMailItems:=rgMailItems, AllOneConvo:=AllOneConvo
    End If
    
    Dim DefaultFolder As String

    If AllOneConvo And UBound(rgMailItems) >= 1 Then
        Dim Convo As Outlook.Conversation
        Set Convo = rgMailItems(1).GetConversation()
        Dim ParentMsg
        On Error Resume Next
        ' This can sometimes fail?
        Set ParentMsg = Convo.GetParent(rgMailItems(1))
        If Not (ParentMsg Is Nothing) Then
            DefaultFolder = ParentMsg.Parent.name
        End If
    End If
        
        
    'Assume this is a mail folder
    Set objFolder = FolderPopup("", DefaultFolder)

    If objFolder Is Nothing Then
        'Call CustomStatus(3, "No valid folder", "Error")
        Exit Sub
    End If
    
    Dim sResults As String
    sResults = ""
    
    If objFolder.DefaultItemType = olMailItem Then
        For i = 1 To UBound(rgMailItems)
            Set objItem = rgMailItems(i)
            sResults = sResults & objItem.Subject & " From: " & objItem.SenderName & vbCrLf & vbCrLf
            If ItemFilter(objItem) Then
                objItem.UnRead = False
                If objItem.Parent <> objFolder Then
                    objItem.Move objFolder
                End If
            End If
        Next
    End If
    
  
    Call CustomStatus(1, "Moved to " & objFolder.name, sResults)
    
    Set objItem = Nothing
    Set objFolder = Nothing
End Sub

Sub OpenSomeFolder()
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem As Outlook.MailItem

    'Assume this is a mail folder
    Set objFolder = FolderPopup("", "")
     
    If objFolder Is Nothing Then
        MsgBox "This folder doesn't exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
        Exit Sub
    End If
    
    Set Application.ActiveExplorer.CurrentFolder = objFolder
End Sub
