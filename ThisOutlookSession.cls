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

Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1

Private WithEvents snItems As Outlook.Items
Attribute snItems.VB_VarHelpID = -1

Dim olNameSpace As Outlook.NameSpace

Dim FolderName As String

FolderName = [folder name here e.g. yourname@domain.com]

Public Sub Application_Startup()
    Set olNameSpace = GetNS(GetOutlookApp)
    Set Items = GetItems(olNameSpace, olFolderInbox)
    Set snItems = Session.GetDefaultFolder(olFolderSentMail).Items
End Sub



Private Sub Items_ItemAdd(ByVal Item As Object)

        Dim olStartFolder As MAPIFolder
        Dim olCaseFolder As MAPIFolder
        Dim olEmail As MailItem
        If TypeName(Item) = "MailItem" Then
        
            Set olEmail = Item
            
            Dim caseNo As String
            

            'caseNo = Get_Match(olEmail.subject, "(\b\d{5}\b)") 'captures things that aren't file numbers
            
            If Get_Match(olEmail.subject) <> "" Then
                caseNo = Get_Match(olEmail.subject)
            'ElseIf Get_Match(olEmail.Body) <> "" Then
            '    caseNo = Get_Match(olEmail.Body)
            Else
                caseNo = ""
            End If
                
            
            If caseNo <> "" Then
            
                ' MsgBox caseNo
             
                 'Search for folder represented by the case number, starting at Personal Folders/Matters
                 
                Set olStartFolder = olNameSpace.Folders(FolderName).Folders("Matters")
                Set olCaseFolder = Find_Folder(olStartFolder, caseNo)
                 
                 'If case number folder found, move this email to it
                 
                If Not olCaseFolder Is Nothing Then
                    With olEmail
                        .FlagStatus = olFlagMarked
                        .FlagRequest = "Follow up"
                        .MarkAsTask olMarkToday
                        .TaskDueDate = Now
                        
                        .Save
                        .Move olCaseFolder
                    End With
                        
                        'olEmail.Move olCaseFolder

                Else
                    ' MsgBox caseNo
                    Dim createFolder As String
                    createFolder = InputBox("Could not find Folder related to File Number: " + caseNo + ". Should we create one now?", "Could not find Folder", caseNo + " - ")
                    If Len(createFolder) = 0 Then
                        Exit Sub
                    Else
                        With olStartFolder
                            .Folders.Add (createFolder)
                        End With
                        
                        With olEmail
                        .FlagStatus = olFlagMarked
                        .FlagRequest = "Follow up"
                        .MarkAsTask olMarkToday
                        .TaskDueDate = Now
                        
                        .Save
                        .Move olNameSpace.Folders(FolderName).Folders("Matters").Folders(createFolder)
                        End With
                    End If
                End If
             
            End If

            
        End If
End Sub
Private Function GetItems(olNS As Outlook.NameSpace, Folder As OlDefaultFolders) As Outlook.Items
    Set GetItems = olNS.GetDefaultFolder(Folder).Items
End Function

Private Function GetNS(ByRef app As Outlook.Application) As Outlook.NameSpace
    Set GetNS = app.GetNamespace("MAPI")
End Function
 
Private Function GetOutlookApp() As Outlook.Application
    Set GetOutlookApp = Outlook.Application
End Function

 'Search for an Outlook folder name recursively, starting at the specified folder.  If found, return the folder as
 'a MAPIfolder object, otherwise return Nothing
 
Public Function Find_Folder(olStartFolder As Outlook.MAPIFolder, findFolderName As String) As Outlook.MAPIFolder
     
    Dim i As Integer
    Dim thisFolder As MAPIFolder
     
    Set Find_Folder = Nothing
    i = 1
    While Find_Folder Is Nothing And i <= olStartFolder.Folders.Count
        Set thisFolder = olStartFolder.Folders(i)
         
         'Debug.Print thisFolder.Name
        If InStr(thisFolder.Name, findFolderName) Then
            Set Find_Folder = thisFolder
       ' ElseIf thisFolder.Folders.Count > 0 Then
             'Search subfolders
         '   Set Find_Folder = Find_Folder(thisFolder, findFolderName)
        End If
         
        i = i + 1
    Wend
     
End Function

 'Return the substring which matches the first capture group of a regular expression pattern
 
Private Function Get_Match(subject As String) As String

    Dim Re As RegExp
    Dim ReMatches As MatchCollection
     
    Get_Match = ""
     
    Set Re = New RegExp
    

        Re.pattern = "(\b[^0]\d{4}\b)"
       ' Re.pattern = "([(])(\d{5})([)])"
    
    Set ReMatches = Re.Execute(subject)
    

    

        If ReMatches.Count > 0 Then Get_Match = ReMatches(0)

    
End Function



'''''
' on send, warn for external emails
'''''

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    Dim prompt As String
    Dim strMsg As String
    Dim domain As String
    
    domain = "[domain goes here]"

    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    Set recips = Item.Recipients
    For Each recip In recips
        Set pa = recip.PropertyAccessor
        If InStr(LCase(pa.GetProperty(PR_SMTP_ADDRESS)), domain) = 0 Then
            strMsg = strMsg & "   " & pa.GetProperty(PR_SMTP_ADDRESS) & vbNewLine
        End If
    Next

    If strMsg <> "" Then
        prompt = "This email will be sent outside of " & domain & " to:" & vbNewLine & strMsg & "Do you want to proceed?"
        If MsgBox(prompt, vbYesNo + vbExclamation + vbMsgBoxSetForeground, "External Email") = vbNo Then
            Cancel = True
        End If
    End If
End Sub


'''''
' on send move to correct folder
'''''

Private Sub snItems_ItemAdd(ByVal Item As Object)
'Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
            Dim olStartFolder As MAPIFolder
             Dim olCaseFolder As MAPIFolder
             Dim olEmail As MailItem

            Set olEmail = Item
            
            Dim caseNo As String
            
            If Get_Match(olEmail.subject) <> "" Then
                caseNo = Get_Match(olEmail.subject)
            'ElseIf Get_Match(olEmail.Body) <> "" Then
             '   caseNo = Get_Match(olEmail.Body)
            Else
                caseNo = ""
            End If
            
            If caseNo <> "" Then
            
                ' MsgBox caseNo
             
                 'Search for folder represented by the case number, starting at Personal Folders/Cases
                 
                Set olStartFolder = olNameSpace.Folders(FolderName).Folders("Matters")
                Set olCaseFolder = Find_Folder(olStartFolder, caseNo)
                 
                 'If case number folder found, move this email to it
                 
                If Not olCaseFolder Is Nothing Then
                
                
                'Dim olapp As New Outlook.Application
        'Dim MyActiveFolder As MAPIFolder
        'Set MyActiveFolder = olapp.ActiveExplorer.CurrentFolder
        
    
    
                    With olEmail   'Item is your e-mail

                            If MsgBox("Save sent email in " + olCaseFolder + " folder?", vbYesNo) = vbYes Then
                                With olEmail
                    
                                  .FlagStatus = olFlagMarked
                                  .FlagRequest = "Follow up"
                                  .MarkAsTask olMarkToday
                                  .TaskDueDate = Now
                                  
                                  .Save
                                  
                                  .Move olCaseFolder
                              End With
                   
                        End If
                    End With
                            
                              
                        
                        'olEmail.Move olCaseFolder

                Else
                    ' MsgBox caseNo
                    Dim createFolder As String
                    createFolder = InputBox("Could not find Folder related to File Number: " + caseNo + ". Should we create one now?", "Could not find Folder", caseNo + " - ")
                    If Len(createFolder) = 0 Then
                        Exit Sub
                    Else
                        With olStartFolder
                            .Folders.Add (createFolder)
                        End With
                        
                        With olEmail
                        .FlagStatus = olFlagMarked
                        .FlagRequest = "Follow up"
                        .MarkAsTask olMarkToday
                        .TaskDueDate = Now
                        
                        .Save
                        .Move olNameSpace.Folders(FolderName).Folders("Matters").Folders(createFolder)
                        End With
                    End If
                End If
             
            End If





        
    
 

End Sub

'''''
' TODO: on save move to correct folder
'''''
